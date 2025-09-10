import xlwings as xw
# Prefer rapidfuzz if available for faster similarity; fallback to fuzzywuzzy
try:
    from rapidfuzz import fuzz as _rf_fuzz  # type: ignore
    _USE_RAPIDFUZZ = True
except Exception:  # rapidfuzz not installed
    from fuzzywuzzy import fuzz as _fw_fuzz
    _USE_RAPIDFUZZ = False
import logging
from typing import List, Dict, Tuple, Optional


def _similarity(a: str, b: str) -> float:
    """Compute string similarity ratio (0-100). Uses rapidfuzz if available."""
    if _USE_RAPIDFUZZ:
        return _rf_fuzz.ratio(a, b)
    else:
        return _fw_fuzz.ratio(a, b)

class TrialBalanceProcessor:
    """Processes Excel trial balance updates using fuzzy matching logic"""
    
    def __init__(self, fuzzy_threshold: int = 80):
        self.fuzzy_threshold = fuzzy_threshold
        self.logger = logging.getLogger(__name__)
        
    def is_cell_bold(self, sheet, cell_address: str) -> bool:
        """Check if a cell is bold.
        
        Args:
            sheet: xlwings sheet object.
            cell_address (str): Cell address (e.g., 'A1').
        
        Returns:
            bool: True if cell is bold, False otherwise.
        """
        try:
            return sheet.range(cell_address).api.Font.Bold
        except:
            return False
    
    def get_excel_status(self) -> Dict[str, any]:
        """Get current Excel application status"""
        try:
            app = xw.apps.active
            if not app or not app.books:
                return {
                    'status': 'no_excel',
                    'message': 'Excel is not open or no workbook is active'
                }
            
            wb = app.books.active
            sheets = [sheet.name for sheet in wb.sheets]
            
            return {
                'status': 'active',
                'workbook': wb.name,
                'sheets': sheets,
                'message': f'Excel is open with workbook "{wb.name}" containing {len(sheets)} sheets'
            }
        except Exception as e:
            return {
                'status': 'error',
                'message': f'Error accessing Excel: {str(e)}'
            }
    
    def get_non_empty_non_bold_data(self, sheet_name: str, column: str, start_row: int = 2) -> List[Tuple[int, any]]:
        """Get non-empty, non-bold data from a column.
        
        Args:
            sheet_name (str): Name of the sheet.
            column (str): Column letter.
            start_row (int): Starting row number.
        
        Returns:
            List[Tuple[int, Any]]: List of (row_number, value) tuples.
        """
        try:
            app = xw.apps.active
            wb = app.books.active
            sheet = wb.sheets[sheet_name]
            data = []
            
            # Find the last used row in the column
            last_row = sheet.range(f"{column}1").end('down').row
            
            for row in range(start_row, last_row + 1):
                cell_address = f"{column}{row}"
                value = sheet.range(cell_address).value
                
                # Skip empty cells and bold cells
                if value is not None and str(value).strip() and not self.is_cell_bold(sheet, cell_address):
                    data.append((row, value))
            
            return data
        except Exception as e:
            self.logger.error(f"Error getting data from column {column}: {str(e)}")
            return []
    
    def analyze_sheet_structure(self, sheet_name: str) -> Dict[str, any]:
        """Analyze the structure of a specific sheet"""
        try:
            app = xw.apps.active
            wb = app.books.active
            sheet = wb.sheets[sheet_name]
            
            # Get all data from the sheet
            data = sheet.used_range.value
            if not data:
                return {
                    'status': 'empty',
                    'message': f'Sheet "{sheet_name}" appears to be empty'
                }
            
            # Analyze the data structure
            accounts = []
            for i, row in enumerate(data):
                if row and len(row) > 0 and row[0]:
                    # Check if this looks like an account row
                    if isinstance(row[0], str) and len(row[0]) > 5 and not row[0].startswith('^'):
                        account_info = {
                            'row_index': i + 1,  # Excel is 1-indexed
                            'account_name': row[0],
                            'row_data': row[:min(6, len(row))]  # First 6 columns
                        }
                        accounts.append(account_info)
            
            return {
                'status': 'success',
                'sheet_name': sheet_name,
                'total_rows': len(data),
                'account_count': len(accounts),
                'accounts': accounts[:10],  # First 10 accounts for preview
                'sample_data': data[:5] if len(data) >= 5 else data,
                'message': f'Found {len(accounts)} potential account entries in "{sheet_name}"'
            }
            
        except Exception as e:
            return {
                'status': 'error',
                'message': f'Error analyzing sheet "{sheet_name}": {str(e)}'
            }
    
    def extract_accounts_from_sheet(self, sheet_name: str, account_col: int = 0, 
                                  amount_cols: List[int] = None, start_row: int | None = None, end_row: int | None = None) -> List[Dict]:
        """Extract account data from a sheet based on the template logic, with bold and empty cell filtering"""
        if amount_cols is None:
            amount_cols = [1, 3]  # Default columns B and D (0-indexed)
            
        try:
            app = xw.apps.active
            wb = app.books.active
            sheet = wb.sheets[sheet_name]
            used = sheet.used_range
            if not used:
                return []
            last_row = used.last_cell.row
            
            # Determine iteration bounds (Excel rows are 1-indexed). Default start at 2 to skip header.
            row_start = max(2, start_row) if start_row is not None else 2
            row_end = min(last_row, end_row) if end_row is not None else last_row
            if row_end < row_start:
                return []
            
            # Helper function to convert column index to letter
            def col_index_to_letter(col_idx: int) -> str:
                """Convert 0-based column index to Excel column letter"""
                result = ""
                while col_idx >= 0:
                    result = chr(col_idx % 26 + ord('A')) + result
                    col_idx = col_idx // 26 - 1
                return result
            
            accounts: List[Dict] = []
            
            # Process each row within the specified range (default from 2 to last_row)
            for row_num in range(row_start, row_end + 1):
                account_col_letter = col_index_to_letter(account_col)
                account_cell_address = f"{account_col_letter}{row_num}"
                
                # Get account name and check if cell is bold or empty
                name = sheet.range(account_cell_address).value
                
                # Skip if cell is empty, bold, or doesn't meet criteria
                if (name is None or 
                    not str(name).strip() or 
                    self.is_cell_bold(sheet, account_cell_address) or
                    not isinstance(name, str) or 
                    len(name) <= 5 or 
                    name.startswith('^')):
                    continue
                
                # Extract amounts from specified columns
                amounts = {}
                for j, amount_col_idx in enumerate(amount_cols):
                    amount_col_letter = col_index_to_letter(amount_col_idx)
                    amount_cell_address = f"{amount_col_letter}{row_num}"
                    
                    # Get amount value, skip if bold (but allow empty amounts)
                    amount_val = sheet.range(amount_cell_address).value
                    if not self.is_cell_bold(sheet, amount_cell_address):
                        amounts[f'amount_{j+1}'] = amount_val
                    else:
                        amounts[f'amount_{j+1}'] = None
                
                accounts.append({
                    'row_index': row_num - 1,  # 0-based for compatibility
                    'excel_row': row_num,  # Excel 1-indexed
                    'account_name': name,
                    **amounts
                })
            
            return accounts
            
        except Exception as e:
            self.logger.error(f"Error extracting accounts from {sheet_name}: {str(e)}")
            return []
    
    def perform_fuzzy_matching(self, source_accounts: List[Dict], 
                             target_accounts: List[Dict]) -> List[Dict]:
        """Perform fuzzy matching between two sets of accounts with exact-match fast path"""
        matches: List[Dict] = []
        
        # Pre-clean and index target accounts for exact matches
        target_index: Dict[str, Dict] = {}
        cleaned_targets: List[Tuple[str, Dict]] = []  # (clean_name, account)
        for t in target_accounts:
            clean_t = t['account_name'].replace('|', '').strip().lower()
            target_index[clean_t] = t
            cleaned_targets.append((clean_t, t))
        
        for s in source_accounts:
            clean_s = s['account_name'].replace('|', '').strip().lower()
            # Exact match shortcut
            if clean_s in target_index:
                t_acc = target_index[clean_s]
                matches.append({
                    'source_account': s,
                    'target_account': t_acc,
                    'match_score': 100.0,
                    'source_name_cleaned': clean_s,
                    'target_name_cleaned': clean_s,
                })
                continue
            
            # Fuzzy search across targets
            best_score = -1.0
            best_target: Optional[Dict] = None
            for clean_t, t in cleaned_targets:
                score = _similarity(clean_s, clean_t)
                if score > best_score:
                    best_score = score
                    best_target = t
            
            if best_target is not None and best_score >= self.fuzzy_threshold:
                matches.append({
                    'source_account': s,
                    'target_account': best_target,
                    'match_score': best_score,
                    'source_name_cleaned': clean_s,
                    'target_name_cleaned': best_target['account_name'].replace('|', '').strip(),
                })
        
        return matches
    
    def update_trial_balance(self, to_update_sheet: str, correct_sheet: str,
                           to_update_cols: Dict[str, str], correct_cols: Dict[str, str],
                           to_update_row_range: Dict[str, int] = None, correct_row_range: Dict[str, int] = None) -> Dict[str, any]:
        """Main function to update trial balance using the template logic with Excel perf tweaks"""
        try:
            # Convert column letters to indices for extract_accounts_from_sheet
            to_update_cols_idx = {
                'account': self.column_letter_to_index(to_update_cols['account']),
                'current_year': self.column_letter_to_index(to_update_cols['current_year']),
                'prior_year': self.column_letter_to_index(to_update_cols['prior_year'])
            }
            correct_cols_idx = {
                'account': self.column_letter_to_index(correct_cols['account']),
                'current_year': self.column_letter_to_index(correct_cols['current_year']),
                'prior_year': self.column_letter_to_index(correct_cols['prior_year'])
            }
            
            # Extract row range parameters
            to_update_start_row = to_update_row_range.get('start_row') if to_update_row_range else None
            to_update_end_row = to_update_row_range.get('end_row') if to_update_row_range else None
            correct_start_row = correct_row_range.get('start_row') if correct_row_range else None
            correct_end_row = correct_row_range.get('end_row') if correct_row_range else None
            
            # Convert 0 to None for auto-detection
            if to_update_end_row == 0:
                to_update_end_row = None
            if correct_end_row == 0:
                correct_end_row = None
            
            # Extract accounts from both sheets
            self.logger.info(f"Extracting accounts from {to_update_sheet}...")
            to_update_accounts = self.extract_accounts_from_sheet(
                to_update_sheet, 
                to_update_cols_idx['account'],
                [to_update_cols_idx['current_year'], to_update_cols_idx['prior_year']],
                start_row=to_update_start_row,
                end_row=to_update_end_row
            )
            
            self.logger.info(f"Extracting accounts from {correct_sheet}...")
            correct_accounts = self.extract_accounts_from_sheet(
                correct_sheet,
                correct_cols_idx['account'],
                [correct_cols_idx['current_year'], correct_cols_idx['prior_year']],
                start_row=correct_start_row,
                end_row=correct_end_row
            )
            
            self.logger.info(f"Found {len(to_update_accounts)} accounts in {to_update_sheet}")
            self.logger.info(f"Found {len(correct_accounts)} accounts in {correct_sheet}")
            
            # Perform fuzzy matching (now includes exact-match fast path)
            matches = self.perform_fuzzy_matching(to_update_accounts, correct_accounts)
            self.logger.info(f"Found {len(matches)} matches above {self.fuzzy_threshold}% threshold")
            
            # Prepare Excel and update amounts with performance settings
            app = xw.apps.active
            wb = app.books.active
            update_sheet = wb.sheets[to_update_sheet]
            
            prev_screen = getattr(app, 'screen_updating', None)
            prev_calc = getattr(app, 'calculation', None)
            prev_alerts = getattr(app, 'display_alerts', None)
            
            updates_made = 0
            try:
                if prev_screen is not None:
                    app.screen_updating = False
                if prev_alerts is not None:
                    app.display_alerts = False
                if prev_calc is not None:
                    app.calculation = 'manual'
                
                current_col_num = to_update_cols_idx['current_year'] + 1  # 1-indexed
                prior_col_num = to_update_cols_idx['prior_year'] + 1
                
                for match in matches:
                    source_row = match['source_account']['excel_row']  # 1-indexed row
                    target_amounts = match['target_account']
                    
                    amt1 = target_amounts.get('amount_1')
                    if amt1 is not None:
                        update_sheet.cells(source_row, current_col_num).value = amt1
                    
                    amt2 = target_amounts.get('amount_2')
                    if amt2 is not None:
                        update_sheet.cells(source_row, prior_col_num).value = amt2
                    
                    updates_made += 1
            finally:
                if prev_screen is not None:
                    app.screen_updating = prev_screen
                if prev_alerts is not None:
                    app.display_alerts = prev_alerts
                if prev_calc is not None:
                    app.calculation = prev_calc
            
            # Identify new accounts (in correct sheet but not in to_update sheet)
            matched_target_names = {match['target_account']['account_name'].lower().strip() 
                                  for match in matches}
            
            new_accounts = []
            for correct_acc in correct_accounts:
                clean_name = correct_acc['account_name'].lower().strip()
                if clean_name not in matched_target_names:
                    new_accounts.append(correct_acc)
            
            # Add verification for the update process
            verification_result = {'verified': True, 'message': 'Update verification completed'}
            if updates_made > 0:
                verification_result = self.verify_updates_made(to_update_sheet, matches, to_update_cols)
            
            return {
                'status': 'success',
                'updates_made': updates_made,
                'matches_found': len(matches),
                'new_accounts_found': len(new_accounts),
                'matches': matches,
                'new_accounts': new_accounts,
                'message': f"Successfully updated {updates_made} accounts. Found {len(new_accounts)} new accounts.",
                'verification': verification_result
            }
            
        except Exception as e:
            error_msg = f"Error updating trial balance: {str(e)}"
            self.logger.error(error_msg)
            return {
                'status': 'error',
                'message': error_msg,
                'verification': {'verified': False, 'message': 'Update failed, no verification performed'}
            }
    
    def verify_updates_made(self, sheet_name: str, matches: List[Dict], 
                          column_mapping: Dict[str, str]) -> Dict[str, any]:
        """Verify that the updates were actually applied to the target sheet"""
        try:
            app = xw.apps.active
            wb = app.books.active
            sheet = wb.sheets[sheet_name]
            
            verified_updates = 0
            failed_updates = []
            
            current_col_idx = self.column_letter_to_index(column_mapping['current_year'])
            prior_col_idx = self.column_letter_to_index(column_mapping['prior_year'])
            
            for match in matches:
                source_row = match['source_account']['excel_row']
                target_amounts = match['target_account']
                
                # Check current year amount
                expected_amt1 = target_amounts.get('amount_1')
                if expected_amt1 is not None:
                    actual_value = sheet.cells(source_row, current_col_idx + 1).value
                    if actual_value == expected_amt1:
                        verified_updates += 1
                    else:
                        failed_updates.append({
                            'account': match['source_account']['account_name'],
                            'row': source_row,
                            'column': 'current_year',
                            'expected': expected_amt1,
                            'actual': actual_value
                        })
                
                # Check prior year amount
                expected_amt2 = target_amounts.get('amount_2')
                if expected_amt2 is not None:
                    actual_value = sheet.cells(source_row, prior_col_idx + 1).value
                    if actual_value == expected_amt2:
                        verified_updates += 1
                    else:
                        failed_updates.append({
                            'account': match['source_account']['account_name'],
                            'row': source_row,
                            'column': 'prior_year',
                            'expected': expected_amt2,
                            'actual': actual_value
                        })
            
            success = len(failed_updates) == 0
            
            if success:
                message = f"Update verification PASSED: All {verified_updates} updates confirmed"
                self.logger.info(f"✅ {message}")
            else:
                message = f"Update verification FAILED: {len(failed_updates)} updates not applied correctly"
                self.logger.warning(f"❌ {message}")
            
            return {
                'verified': success,
                'verified_updates': verified_updates,
                'failed_updates': failed_updates,
                'message': message
            }
            
        except Exception as e:
            error_msg = f"Update verification failed due to error: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            return {
                'verified': False,
                'error': error_msg,
                'message': error_msg
            }
    
    def column_letter_to_index(self, col_letter: str) -> int:
        """Convert Excel column letter (A, B, C, etc.) to 0-based index"""
        result = 0
        for char in col_letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def add_new_accounts(self, sheet_name: str, new_accounts: List[Dict], 
                        column_mapping: Dict[str, str], row_range: Dict[str, int] = None) -> Dict[str, any]:
        """Add new accounts to the specified sheet with highlighting"""
        try:
            app = xw.apps.active
            wb = app.books.active
            sheet = wb.sheets[sheet_name]
            
            # Debug: Log the operation
            self.logger.info(f"Adding {len(new_accounts)} new accounts to sheet '{sheet_name}'")
            self.logger.info(f"Column mapping: {column_mapping}")
            
            # Find the last row with data
            last_row = sheet.used_range.last_cell.row
            self.logger.info(f"Starting from row {last_row + 1}")
            
            accounts_added = 0
            highlighted_rows = []
            
            for account in new_accounts:
                last_row += 1
                highlighted_rows.append(last_row)
                
                # Add account name
                account_col = column_mapping['account']
                account_cell = sheet.range(f"{account_col}{last_row}")
                account_cell.value = account['account_name']
                
                # Add amounts
                if account.get('amount_1') is not None:
                    current_col = column_mapping['current_year']
                    current_cell = sheet.range(f"{current_col}{last_row}")
                    current_cell.value = account['amount_1']
                
                if account.get('amount_2') is not None:
                    prior_col = column_mapping['prior_year']
                    prior_cell = sheet.range(f"{prior_col}{last_row}")
                    prior_cell.value = account['amount_2']
                
                accounts_added += 1
            
            # Highlight all newly added rows with light yellow background
            if highlighted_rows:
                try:
                    # Get all columns that have data to highlight the entire row
                    all_cols = [column_mapping['account'], column_mapping['current_year'], column_mapping['prior_year']]
                    
                    for row in highlighted_rows:
                        for col in all_cols:
                            cell = sheet.range(f"{col}{row}")
                            # Set bright yellow background using Excel color index
                            cell.color = 65535  # Bright yellow color
                            # Make text bold to emphasize new accounts
                            cell.api.Font.Bold = True
                            
                    # Force Excel to refresh/recalculate
                    wb.save()
                    
                except Exception as highlight_error:
                    error_msg = f"Could not highlight new accounts: {str(highlight_error)}"
                    self.logger.warning(error_msg)
                    # Return the error in the result so GUI can show it
                    return {
                        'status': 'partial_success',
                        'accounts_added': accounts_added,
                        'message': f"Added {accounts_added} new accounts to {sheet_name} but highlighting failed: {str(highlight_error)}",
                        'highlighted_rows': highlighted_rows
                    }
            
            # Verify accounts were actually added
            verification_result = self.verify_accounts_added(sheet_name, new_accounts, column_mapping, row_range)
            
            return {
                'status': 'success',
                'accounts_added': accounts_added,
                'message': f"Successfully added {accounts_added} new accounts to {sheet_name} with highlighting",
                'highlighted_rows': highlighted_rows,
                'verification': verification_result
            }
            
        except Exception as e:
            error_msg = f"Error adding new accounts: {str(e)}"
            self.logger.error(error_msg)
            return {
                'status': 'error',
                'message': error_msg
            }
    
    def verify_accounts_added(self, sheet_name: str, expected_accounts: List[Dict], 
                            column_mapping: Dict[str, str], row_range: Dict[str, int] = None) -> Dict[str, any]:
        """Verify that the accounts were actually added to the target sheet"""
        try:
            # Convert column letters to indices for extract_accounts_from_sheet
            account_col_idx = self.column_letter_to_index(column_mapping['account'])
            current_col_idx = self.column_letter_to_index(column_mapping['current_year'])
            prior_col_idx = self.column_letter_to_index(column_mapping['prior_year'])
            
            # Extract row range parameters
            start_row = row_range.get('start_row') if row_range else None
            end_row = row_range.get('end_row') if row_range else None
            if end_row == 0:
                end_row = None
            
            # Re-extract accounts from the target sheet
            current_accounts = self.extract_accounts_from_sheet(
                sheet_name, 
                account_col_idx,
                [current_col_idx, prior_col_idx],
                start_row=start_row,
                end_row=end_row
            )
            
            # Check if each expected account is now present
            current_names = {acc['account_name'].replace('|', '').strip().lower() 
                           for acc in current_accounts}
            
            verified_count = 0
            missing_accounts = []
            
            for expected_acc in expected_accounts:
                expected_name = expected_acc['account_name'].replace('|', '').strip().lower()
                if expected_name in current_names:
                    verified_count += 1
                else:
                    missing_accounts.append(expected_acc['account_name'])
            
            success = verified_count == len(expected_accounts)
            
            if success:
                message = f"Verification PASSED: All {len(expected_accounts)} accounts successfully added"
                self.logger.info(f"✅ {message}")
            else:
                message = f"Verification FAILED: {verified_count}/{len(expected_accounts)} accounts found. Missing: {missing_accounts}"
                self.logger.warning(f"❌ {message}")
            
            return {
                'verified': success,
                'expected_count': len(expected_accounts),
                'verified_count': verified_count,
                'missing_accounts': missing_accounts,
                'message': message
            }
            
        except Exception as e:
            error_msg = f"Verification failed due to error: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            return {
                'verified': False,
                'error': error_msg,
                'message': error_msg
            }
    
    def get_column_preview(self, sheet_name=None, column_name=None, max_rows=10):
        """Get preview of data in columns or all sheets if no specific column specified"""
        try:
            wb = xw.books.active
            if not wb:
                return "No active workbook found"
            
            if sheet_name and column_name:
                # Preview specific column
                ws = wb.sheets[sheet_name]
                
                # Find the column by header
                header_row = ws.range('A1:Z1').value
                if column_name not in header_row:
                    return f"Column '{column_name}' not found"
                
                col_index = header_row.index(column_name) + 1
                col_letter = chr(64 + col_index)  # Convert to letter (A, B, C, etc.)
                
                # Get data from the column
                data_range = ws.range(f'{col_letter}2:{col_letter}{max_rows + 1}')
                values = data_range.value
                
                if isinstance(values, list):
                    preview = [str(v) for v in values if v is not None]
                else:
                    preview = [str(values)] if values is not None else []
                
                return f"Preview of '{column_name}' column:\n" + "\n".join(preview[:max_rows])
            else:
                # Preview all sheets and their columns
                preview_text = f"📊 Workbook: {wb.name}\n\n"
                
                for sheet in wb.sheets:
                    try:
                        # Get data from the sheet (first 10 rows, up to column Z)
                        data_range = sheet.range('A1:Z10').value
                        
                        preview_text += f"📋 Sheet: {sheet.name}\n"
                        preview_text += "=" * 50 + "\n"
                        
                        if data_range:
                            # Handle single row case
                            if not isinstance(data_range, list):
                                data_range = [[data_range]]
                            elif len(data_range) > 0 and not isinstance(data_range[0], list):
                                data_range = [data_range]
                            
                            # Create table format
                            for row_idx, row in enumerate(data_range[:10]):
                                if row and any(cell is not None for cell in row):
                                    # Convert all cells to strings and handle None values
                                    row_data = [str(cell) if cell is not None else "" for cell in row[:10]]
                                    # Only show non-empty rows
                                    if any(cell.strip() for cell in row_data if isinstance(cell, str)):
                                        preview_text += f"Row {row_idx + 1:2d}: {' | '.join(f'{cell:15s}' for cell in row_data)}\n"
                            
                            preview_text += "\n"
                        else:
                            preview_text += "No data found in this sheet\n\n"
                    
                    except Exception as sheet_error:
                        preview_text += f"📋 Sheet: {sheet.name} (Error: {str(sheet_error)})\n\n"
                
                return preview_text
            
        except Exception as e:
            return f"Error getting column preview: {str(e)}"
    
    def get_column_headers(self, sheet_name=None):
        """Get simple column labels (Column A, Column B, etc.) up to Column F.
        
        Returns a list of formatted strings like 'A: Column A', 'B: Column B', etc.
        """
        try:
            # Helper function to convert column numbers to Excel letters
            def column_number_to_letter(col_num):
                result = ""
                while col_num > 0:
                    col_num -= 1
                    result = chr(col_num % 26 + ord('A')) + result
                    col_num //= 26
                return result
            
            # Simply return generic column labels A through F
            column_info = []
            for col_idx in range(6):  # A through F (0-5)
                col_letter = column_number_to_letter(col_idx + 1)
                column_info.append(f"{col_letter}: Column {col_letter}")
            
            return column_info
            
        except Exception as e:
            print(f"Error getting column headers: {e}")
            return []
    
    def analyze_workbook_structure(self):
        """Analyze the entire workbook structure"""
        try:
            wb = xw.books.active
            if not wb:
                return "No active workbook found"
            
            analysis = f"📊 Workbook Analysis: {wb.name}\n\n"
            
            for sheet in wb.sheets:
                try:
                    # Get basic sheet info
                    used_range = sheet.used_range
                    if used_range:
                        rows = used_range.last_cell.row
                        cols = used_range.last_cell.column
                    else:
                        rows = cols = 0
                    
                    analysis += f"📋 Sheet: {sheet.name}\n"
                    analysis += f"   Size: {rows} rows × {cols} columns\n"
                    
                    # Get column headers
                    headers = sheet.range('A1:Z1').value
                    if isinstance(headers, list):
                        headers = [h for h in headers if h is not None]
                    else:
                        headers = [headers] if headers is not None else []
                    
                    analysis += f"   Headers: {', '.join(headers[:10])}{'...' if len(headers) > 10 else ''}\n"
                    
                    # Check for potential account columns
                    account_keywords = ['account', 'name', 'description', 'code']
                    amount_keywords = ['amount', 'balance', 'total', 'value', 'current', 'prior']
                    
                    potential_accounts = [h for h in headers if any(keyword in h.lower() for keyword in account_keywords)]
                    potential_amounts = [h for h in headers if any(keyword in h.lower() for keyword in amount_keywords)]
                    
                    if potential_accounts:
                        analysis += f"   🏷️ Potential Account Columns: {', '.join(potential_accounts)}\n"
                    if potential_amounts:
                        analysis += f"   💰 Potential Amount Columns: {', '.join(potential_amounts)}\n"
                    
                    analysis += "\n"
                    
                except Exception as e:
                    analysis += f"📋 Sheet: {sheet.name} (Error: {str(e)})\n\n"
            
            # Add recommendations
            analysis += "💡 Recommendations:\n"
            analysis += "• Use sheets with similar account structures for trial balance updates\n"
            analysis += "• Ensure account name columns are clearly identifiable\n"
            analysis += "• Check that amount columns contain numeric data\n"
            
            return analysis
            
        except Exception as e:
            return f"Error analyzing workbook: {str(e)}"
