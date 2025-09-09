import os
import sys
from pathlib import Path
import xlwings as xw
from fuzzywuzzy import fuzz
from dotenv import load_dotenv
from loguru import logger
from colorama import Fore, Style, init

# Initialize colorama
init(autoreset=True)

# Load environment variables
load_dotenv()

class ExcelTrialBalanceAgent:
    """
    An autonomous agent for updating Excel trial balance sheets.
    """

    def __init__(self):
        """
        Initialize the agent with settings from environment variables.
        """
        self.app = None
        self.workbook = None
        self.setup_logging()
        logger.info("ExcelTrialBalanceAgent initialized.")

    def setup_logging(self):
        """
        Configure Loguru logger.
        """
        log_level = os.getenv("LOG_LEVEL", "INFO")
        logger.add("excel_agent.log", level=log_level, rotation="1 MB",
                   format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}")

    def run(self):
        """
        Execute the main logic of the agent.
        """
        try:
            self.connect_to_excel()
            self.interactive_setup()
            self.process_sheets()
        except Exception as e:
            logger.error(f"An error occurred: {e}")
            print(f"{Fore.RED}An error occurred: {e}")
        finally:
            logger.info("Agent run finished.")

    def connect_to_excel(self):
        """
        Connect to an active Excel instance and workbook.
        """
        logger.info("Connecting to Excel...")
        try:
            self.app = xw.apps.active
            if self.app is None:
                raise Exception("No active Excel instance found.")
            self.workbook = self.app.books.active
            if self.workbook is None:
                raise Exception("No active workbook found.")
            logger.info(f"Connected to workbook: {self.workbook.name}")
            print(f"{Fore.GREEN}Connected to Excel file: {self.workbook.name}")
        except Exception as e:
            logger.error(f"Failed to connect to Excel: {e}")
            print(f"{Fore.RED}Failed to connect to Excel: {e}")
            sys.exit(1)

    def interactive_setup(self):
        """
        Guide the user through selecting sheets and columns.
        """
        print("\n=== Excel Trial Balance Agent - Interactive Setup ===")
        # Sheet selection
        sheet_names = [sheet.name for sheet in self.workbook.sheets]
        print(f"Available sheets: {', '.join(sheet_names)}")

        self.to_update_sheet_name = self._get_sheet_input("to be updated (the one with incorrect amounts)")
        self.reference_sheet_name = self._get_sheet_input("with correct amounts (reference data)")

        # Column mapping
        self.to_update_cols = self._get_column_mapping(self.to_update_sheet_name)
        self.reference_cols = self._get_column_mapping(self.reference_sheet_name)

    def _get_sheet_input(self, purpose):
        """
        Prompt the user to select a sheet for a specific purpose.
        """
        while True:
            sheet_name = input(f"✅ Which sheet contains the data {purpose}?\nEnter sheet name: ")
            if sheet_name in [s.name for s in self.workbook.sheets]:
                return sheet_name
            print(f"{Fore.YELLOW}Sheet '{sheet_name}' not found. Please try again.")

    def _get_column_mapping(self, sheet_name):
        """
        Prompt the user to map columns for a given sheet.
        """
        print(f"\n🟩 Setting up '{sheet_name}' sheet mappings:")
        sheet = self.workbook.sheets[sheet_name]
        # A simple way to show available columns, can be improved
        headers = sheet.range('A1').expand('right').value
        print(f"Available columns: {headers}")

        account_col = self._get_column_input("account names")
        current_year_col = self._get_column_input("current year amounts")
        prior_year_col = self._get_column_input("prior year amounts")

        return {
            "account": account_col,
            "current_year": current_year_col,
            "prior_year": prior_year_col
        }

    def _get_column_input(self, purpose):
        """
        Prompt the user to enter a column letter.
        """
        return input(f"Which column contains the {purpose}?\nEnter column letter (e.g., A, B, C): ").upper()

    def process_sheets(self):
        """
        Extract data, match accounts, update amounts, and add new accounts.
        """
        print("\n=== Starting Update Process ===")
        # Extract data
        to_update_data = self._extract_data(self.to_update_sheet_name, self.to_update_cols)
        reference_data = self._extract_data(self.reference_sheet_name, self.reference_cols)

        # Match and update
        self._match_and_update(to_update_data, reference_data)

        # Add new accounts
        self._add_new_accounts(to_update_data, reference_data)

    def _extract_data(self, sheet_name, cols):
        """
        Extract account data from a sheet.
        """
        sheet = self.workbook.sheets[sheet_name]
        data = {}
        # Simplified data extraction logic
        for row in range(2, sheet.range(f'{cols["account"]}1').end('down').row + 1):
            account_name = sheet.range(f'{cols["account"]}{row}').value
            if account_name:
                data[account_name] = {
                    "current_year": sheet.range(f'{cols["current_year"]}{row}').value,
                    "prior_year": sheet.range(f'{cols["prior_year"]}{row}').value,
                    "row": row
                }
        logger.info(f"Extracted {len(data)} accounts from '{sheet_name}'.")
        print(f"Found {len(data)} accounts in '{sheet_name}' sheet")
        return data

    def _match_and_update(self, to_update_data, reference_data):
        """
        Match accounts and update amounts in the to-update sheet.
        """
        logger.info("Performing fuzzy matching and updating amounts...")
        updated_count = 0
        for ref_name, ref_values in reference_data.items():
            best_match = None
            highest_score = 0
            for upd_name in to_update_data.keys():
                score = fuzz.ratio(ref_name.lower(), upd_name.lower())
                if score > highest_score:
                    highest_score = score
                    best_match = upd_name

            if highest_score >= 80: # 80% threshold from README
                upd_row = to_update_data[best_match]["row"]
                sheet = self.workbook.sheets[self.to_update_sheet_name]
                sheet.range(f'{self.to_update_cols["current_year"]}{upd_row}').value = ref_values["current_year"]
                sheet.range(f'{self.to_update_cols["prior_year"]}{upd_row}').value = ref_values["prior_year"]
                logger.info(f"Updated account '{best_match}' with data from '{ref_name}'.")
                updated_count += 1

        print(f"Successfully updated {updated_count} accounts.")

    def _add_new_accounts(self, to_update_data, reference_data):
        """
        Add new accounts from the reference sheet to the to-update sheet.
        """
        logger.info("Finding and adding new accounts...")
        new_accounts_added = 0
        to_update_sheet = self.workbook.sheets[self.to_update_sheet_name]
        last_row = to_update_sheet.range(f'{self.to_update_cols["account"]}1').end('down').row

        ref_account_names = set(reference_data.keys())
        upd_account_names = set(to_update_data.keys())

        # Find accounts in reference but not in to-update
        new_accounts = ref_account_names - upd_account_names

        for account_name in new_accounts:
            # A simple check to see if it's already matched fuzzily
            is_fuzzily_matched = False
            for upd_name in upd_account_names:
                if fuzz.ratio(account_name.lower(), upd_name.lower()) >= 80:
                    is_fuzzily_matched = True
                    break

            if not is_fuzzily_matched:
                last_row += 1
                ref_values = reference_data[account_name]
                to_update_sheet.range(f'{self.to_update_cols["account"]}{last_row}').value = account_name
                to_update_sheet.range(f'{self.to_update_cols["current_year"]}{last_row}').value = ref_values["current_year"]
                to_update_sheet.range(f'{self.to_update_cols["prior_year"]}{last_row}').value = ref_values["prior_year"]
                logger.info(f"Added new account '{account_name}' at row {last_row}.")
                new_accounts_added += 1

        if new_accounts_added > 0:
            print(f"Successfully added {new_accounts_added} new accounts.")
        else:
            print("No new accounts to add.")
