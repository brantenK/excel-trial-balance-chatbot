import sys
import json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QTextEdit, QLineEdit, QPushButton, QScrollArea, QFrame, QLabel,
    QMessageBox, QProgressBar, QDialog, QDialogButtonBox, QComboBox,
    QCheckBox, QSpinBox, QGroupBox, QGridLayout, QSplitter, QTabWidget,
    QFileDialog, QListWidget, QListWidgetItem, QTextBrowser, QSizePolicy
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QTimer, QSize
from PyQt6.QtGui import QFont, QPixmap, QIcon, QPalette, QColor, QTextCursor
import xlwings as xw
from fuzzywuzzy import fuzz
import requests
import os
from dotenv import load_dotenv
import logging
from datetime import datetime
import traceback
from excel_processor import TrialBalanceProcessor

# Load environment variables
load_dotenv()

# --- Start Rebuilt UI Components ---

class ColumnMappingDialog(QDialog):
    """A dialog to map columns for a sheet."""
    def __init__(self, sheet_name, headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Column Mapping for '{sheet_name}'")
        self.layout = QGridLayout(self)

        # Fallback if headers are not found
        if not headers:
            headers = [f"Column {chr(ord('A') + i)}" for i in range(10)] # Default to 10 columns

        self.headers = headers
        self.column_letters = [chr(ord('A') + i) for i in range(len(headers))]

        # Create widgets
        self.layout.addWidget(QLabel(f"Configure columns for sheet '{sheet_name}':"), 0, 0, 1, 2)

        self.layout.addWidget(QLabel("Account Name Column:"), 1, 0)
        self.account_combo = QComboBox()
        self.account_combo.addItems([f"{letter}: {name}" for letter, name in zip(self.column_letters, self.headers)])
        self.layout.addWidget(self.account_combo, 1, 1)

        self.layout.addWidget(QLabel("Current Year Column:"), 2, 0)
        self.current_year_combo = QComboBox()
        self.current_year_combo.addItems([f"{letter}: {name}" for letter, name in zip(self.column_letters, self.headers)])
        self.layout.addWidget(self.current_year_combo, 2, 1)

        self.layout.addWidget(QLabel("Prior Year Column:"), 3, 0)
        self.prior_year_combo = QComboBox()
        self.prior_year_combo.addItems([f"{letter}: {name}" for letter, name in zip(self.column_letters, self.headers)])
        self.layout.addWidget(self.prior_year_combo, 3, 1)

        # Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box, 4, 0, 1, 2)

    def get_mapping(self):
        """Return the selected column mapping."""
        return {
            "account": self.account_combo.currentText().split(':')[0],
            "current_year": self.current_year_combo.currentText().split(':')[0],
            "prior_year": self.prior_year_combo.currentText().split(':')[0],
        }

class UpdateSetupDialog(QDialog):
    """A dialog to select sheets for the update process."""
    def __init__(self, sheet_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Setup Trial Balance Update")
        self.layout = QGridLayout(self)

        self.sheet_names = sheet_names

        # Create widgets
        self.layout.addWidget(QLabel("Select the sheets for the update process:"), 0, 0, 1, 2)

        self.layout.addWidget(QLabel("Sheet to Update:"), 1, 0)
        self.to_update_combo = QComboBox()
        self.to_update_combo.addItems(sheet_names)
        self.layout.addWidget(self.to_update_combo, 1, 1)

        self.layout.addWidget(QLabel("Reference Sheet (with correct values):"), 2, 0)
        self.reference_combo = QComboBox()
        self.reference_combo.addItems(sheet_names)
        self.layout.addWidget(self.reference_combo, 2, 1)

        # Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box, 3, 0, 1, 2)

    def get_selection(self):
        """Return the selected sheet names."""
        to_update = self.to_update_combo.currentText()
        reference = self.reference_combo.currentText()
        if to_update == reference:
            QMessageBox.warning(self, "Selection Error", "The 'to-update' sheet and the 'reference' sheet cannot be the same.")
            return None
        return {
            "to_update_sheet": to_update,
            "reference_sheet": reference,
        }

# --- End Rebuilt UI Components ---

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_chatbot.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ExcelChatBot(QThread):
    """Background thread for handling Excel operations and API calls"""
    
    message_received = pyqtSignal(str, str)  # message, sender
    error_occurred = pyqtSignal(str)
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.processor = TrialBalanceProcessor()
        self.api_key = os.getenv('OPENROUTER_API_KEY')
        self.api_url = "https://openrouter.ai/api/v1/chat/completions"
        self.conversation_history = []
        self.current_request = None
        self.is_processing = False
        
    def set_request(self, request_type, data=None):
        """Set the current request to be processed"""
        self.current_request = {
            'type': request_type,
            'data': data or {}
        }
        
    def run(self):
        """Main thread execution"""
        if not self.current_request:
            return
            
        try:
            self.is_processing = True
            request_type = self.current_request['type']
            data = self.current_request['data']
            
            if request_type == 'analyze_structure':
                self._analyze_excel_structure()
            elif request_type == 'guide_update':
                self._guide_trial_balance_update(data.get('user_message', ''))
            elif request_type == 'chat':
                self._handle_chat_message(data.get('message', ''))
            elif request_type == 'perform_update':
                self._perform_trial_balance_update(data)
                
        except Exception as e:
            logger.error(f"Error in thread execution: {str(e)}")
            self.error_occurred.emit(f"An error occurred: {str(e)}")
        finally:
            self.is_processing = False
            
    def _analyze_excel_structure(self):
        """Analyze the structure of the active Excel workbook"""
        try:
            self.status_updated.emit("Analyzing Excel structure...")
            
            # Get Excel status
            excel_status = self.processor.get_excel_status()
            
            if not excel_status['has_excel']:
                self.message_received.emit(
                    "❌ No Excel application found. Please make sure Excel is installed and running.",
                    "assistant"
                )
                return
                
            if not excel_status['has_workbook']:
                self.message_received.emit(
                    "📋 No Excel workbook is currently open. Please open a workbook and try again.",
                    "assistant"
                )
                return
                
            # Analyze structure
            structure = self.processor.analyze_structure()
            
            # Format the analysis message
            message = f"📊 **Excel Workbook Analysis**\n\n"
            message += f"**Workbook:** {structure['workbook_name']}\n"
            message += f"**Active Sheet:** {structure['active_sheet']}\n\n"
            
            message += "**Available Sheets:**\n"
            for sheet in structure['sheets']:
                message += f"• {sheet}\n"
                
            message += f"\n**Data Range:** {structure['data_range']}\n"
            message += f"**Total Rows:** {structure['total_rows']}\n"
            message += f"**Total Columns:** {structure['total_columns']}\n\n"
            
            if structure['headers']:
                message += "**Column Headers:**\n"
                for i, header in enumerate(structure['headers'], 1):
                    message += f"{i}. {header}\n"
            
            self.message_received.emit(message, "assistant")
            self.status_updated.emit("Analysis complete")
            
        except Exception as e:
            logger.error(f"Error analyzing Excel structure: {str(e)}")
            self.error_occurred.emit(f"Failed to analyze Excel structure: {str(e)}")
            
    def _guide_trial_balance_update(self, user_message):
        """Guide the user through trial balance update process"""
        try:
            self.status_updated.emit("Processing your request...")
            
            # Get Excel status first
            excel_status = self.processor.get_excel_status()
            
            if not excel_status['has_excel'] or not excel_status['has_workbook']:
                self.message_received.emit(
                    "Please ensure Excel is running with a workbook open before proceeding.",
                    "assistant"
                )
                return
                
            # Prepare context for AI
            context = {
                'user_message': user_message,
                'excel_status': excel_status,
                'conversation_history': self.conversation_history[-5:]  # Last 5 messages for context
            }
            
            # Call OpenRouter API
            response = self._call_openrouter_api(context)
            
            if response:
                self.message_received.emit(response, "assistant")
                # Add to conversation history
                self.conversation_history.append({
                    'role': 'user',
                    'content': user_message
                })
                self.conversation_history.append({
                    'role': 'assistant', 
                    'content': response
                })
            else:
                self.message_received.emit(
                    "I'm having trouble connecting to the AI service. Please try again later.",
                    "assistant"
                )
                
            self.status_updated.emit("Ready")
            
        except Exception as e:
            logger.error(f"Error in guide update: {str(e)}")
            self.error_occurred.emit(f"Failed to process request: {str(e)}")
            
    def _handle_chat_message(self, message):
        """Handle general chat messages"""
        try:
            self.status_updated.emit("Thinking...")
            
            # Simple keyword-based responses for common queries
            message_lower = message.lower()
            
            if any(word in message_lower for word in ['help', 'what can you do', 'commands']):
                response = """🤖 **Excel Trial Balance Assistant**

I can help you with:

**📊 Analysis:**
• Analyze Excel workbook structure
• Identify trial balance data
• Review column mappings

**🔄 Updates:**
• Guide you through trial balance updates
• Perform automated updates with your approval
• Verify update results

**💬 Chat:**
• Answer questions about Excel operations
• Provide guidance on trial balance processes
• Help troubleshoot issues

**Commands:**
• Type 'analyze' to analyze current workbook
• Type 'update' to start update process
• Ask any questions about your Excel data!"""
                
            elif 'analyze' in message_lower:
                self.set_request('analyze_structure')
                self.start()
                return
                
            elif 'update' in message_lower:
                response = """🔄 **Trial Balance Update Process**

To update your trial balance, I'll need to:

1. **Analyze** your current Excel structure
2. **Identify** trial balance columns (Account, Debit, Credit)
3. **Map** your data to standard format
4. **Preview** proposed changes
5. **Execute** updates with your approval

Would you like me to start by analyzing your current workbook structure?"""
                
            else:
                # For other messages, try to use AI if available
                if self.api_key:
                    context = {
                        'user_message': message,
                        'conversation_history': self.conversation_history[-3:]
                    }
                    response = self._call_openrouter_api(context)
                    if not response:
                        response = "I'm here to help with Excel trial balance operations. Try asking about 'analyze', 'update', or 'help'."
                else:
                    response = "I'm here to help with Excel trial balance operations. Try asking about 'analyze', 'update', or 'help'."
            
            self.message_received.emit(response, "assistant")
            self.status_updated.emit("Ready")
            
        except Exception as e:
            logger.error(f"Error handling chat message: {str(e)}")
            self.error_occurred.emit(f"Failed to process message: {str(e)}")
            
    def _call_openrouter_api(self, context):
        """Call OpenRouter API for AI responses"""
        if not self.api_key:
            return None
            
        try:
            # Prepare the prompt
            system_prompt = """You are an Excel Trial Balance Assistant. You help users analyze and update Excel trial balance data.
            
Your capabilities include:
- Analyzing Excel workbook structure
- Identifying trial balance data patterns
- Guiding users through update processes
- Providing Excel-related advice

Be helpful, concise, and focus on Excel trial balance operations. Use emojis and formatting to make responses clear and engaging."""
            
            messages = [
                {"role": "system", "content": system_prompt}
            ]
            
            # Add conversation history
            if 'conversation_history' in context:
                messages.extend(context['conversation_history'])
                
            # Add current message
            messages.append({
                "role": "user",
                "content": context['user_message']
            })
            
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            data = {
                "model": "anthropic/claude-3.5-sonnet",
                "messages": messages,
                "max_tokens": 1000,
                "temperature": 0.7
            }
            
            response = requests.post(self.api_url, headers=headers, json=data, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content']
            else:
                logger.error(f"API call failed: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"Error calling OpenRouter API: {str(e)}")
            return None
            
    def _perform_trial_balance_update(self, update_data):
        """Perform the trial balance update and add new accounts."""
        try:
            self.status_updated.emit("Performing trial balance update...")
            self.progress_updated.emit(10)

            # 1. Extract parameters from the new data structure
            to_update_sheet = update_data.get('to_update_sheet')
            reference_sheet = update_data.get('reference_sheet')
            to_update_cols = update_data.get('to_update_cols')
            reference_cols = update_data.get('reference_cols')

            if not all([to_update_sheet, reference_sheet, to_update_cols, reference_cols]):
                self.error_occurred.emit("Missing data for update process. Please start over.")
                return

            # 2. Perform the update using the processor
            self.status_updated.emit("Matching and updating existing accounts...")
            self.progress_updated.emit(30)
            update_result = self.processor.update_trial_balance(
                to_update_sheet=to_update_sheet,
                correct_sheet=reference_sheet,
                to_update_cols=to_update_cols,
                correct_cols=reference_cols
            )

            self.progress_updated.emit(70)

            if update_result.get('status') != 'success':
                self.error_occurred.emit(f"Update failed: {update_result.get('message', 'Unknown error')}")
                return

            # 3. Add new accounts if any were found
            summary_message = f"✅ **Update Process Complete!**\n\n"
            summary_message += f"**Updated {update_result.get('updates_made', 0)} existing accounts.**\n"

            new_accounts = update_result.get('new_accounts', [])
            if new_accounts:
                self.status_updated.emit(f"Adding {len(new_accounts)} new accounts...")
                self.progress_updated.emit(90)
                
                add_result = self.processor.add_new_accounts(
                    sheet_name=to_update_sheet,
                    new_accounts=new_accounts,
                    column_mapping=to_update_cols
                )

                if add_result.get('status') == 'success':
                    summary_message += f"**Successfully added {add_result.get('accounts_added', 0)} new accounts.**\n"
                    summary_message += "New accounts added:\n"
                    for acc in new_accounts[:5]: # Preview first 5
                         summary_message += f"• {acc.get('account_name')}\n"
                    if len(new_accounts) > 5:
                        summary_message += "...\n"
                else:
                    summary_message += f"**⚠️ Failed to add new accounts:** {add_result.get('message', 'Unknown error')}\n"
            else:
                summary_message += "**No new accounts found to add.**\n"

            self.message_received.emit(summary_message, "assistant")
            self.progress_updated.emit(100)
            self.status_updated.emit("Update complete")

        except Exception as e:
            logger.error(f"Error performing update: {str(e)}\n{traceback.format_exc()}")
            self.error_occurred.emit(f"Update failed: {str(e)}")

class ChatMessage(QFrame):
    """Individual chat message widget"""
    
    def __init__(self, message, sender, timestamp=None):
        super().__init__()
        self.message = message
        self.sender = sender
        self.timestamp = timestamp or datetime.now().strftime("%H:%M")
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the message UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 5, 10, 5)
        
        # Message header
        header_layout = QHBoxLayout()
        
        sender_label = QLabel(f"{'🤖 Assistant' if self.sender == 'assistant' else '👤 You'}")
        sender_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        
        time_label = QLabel(self.timestamp)
        time_label.setFont(QFont("Arial", 8))
        time_label.setStyleSheet("color: #666;")
        
        header_layout.addWidget(sender_label)
        header_layout.addStretch()
        header_layout.addWidget(time_label)
        
        # Message content
        content_label = QTextBrowser()
        content_label.setMarkdown(self.message)
        content_label.setMaximumHeight(200)
        content_label.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content_label.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # Style the message based on sender
        if self.sender == "assistant":
            self.setStyleSheet("""
                QFrame {
                    background-color: #f0f8ff;
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    margin: 2px;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame {
                    background-color: #f5f5f5;
                    border: 1px solid #d0d0d0;
                    border-radius: 8px;
                    margin: 2px;
                }
            """)
            
        layout.addLayout(header_layout)
        layout.addWidget(content_label)
        self.setLayout(layout)

class ExcelChatBotGUI(QWidget):
    """Main GUI application for Excel ChatBot, designed as a side pane."""
    
    def __init__(self):
        super().__init__()
        self.chatbot = ExcelChatBot()
        self.setup_ui()
        self.setup_connections()
        self.setup_styling()
        
        # Welcome message
        self.add_message(
            "👋 Welcome to Excel Trial Balance Assistant!\n\n" +
            "I can help you analyze and update Excel trial balance data. " +
            "Type 'help' to see what I can do, or 'analyze' to start analyzing your current workbook.",
            "assistant"
        )
        
    def setup_ui(self):
        """Setup the main user interface as a compact widget."""
        self.setWindowTitle("Excel Assistant")
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setGeometry(100, 100, 400, 600)

        # Since we are a QWidget now, we set the layout directly on self
        main_layout = QVBoxLayout(self)

        # The chat panel is now the main and only component
        chat_widget = self.create_chat_panel()
        main_layout.addWidget(chat_widget)

        # Create a custom status area
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #666;")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(15)
        self.progress_bar.setMaximumWidth(120)

        status_layout.addWidget(self.status_label)
        status_layout.addStretch()
        status_layout.addWidget(self.progress_bar)

        main_layout.addLayout(status_layout)
        
    def create_chat_panel(self):
        """Create the chat panel"""
        chat_widget = QWidget()
        layout = QVBoxLayout(chat_widget)
        
        # Chat history
        self.chat_scroll = QScrollArea()
        self.chat_scroll.setWidgetResizable(True)
        self.chat_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.addStretch()
        
        self.chat_scroll.setWidget(self.chat_container)
        
        # Input area
        input_layout = QHBoxLayout()
        
        self.message_input = QLineEdit()
        self.message_input.setPlaceholderText("Type your message here...")
        self.message_input.setMinimumHeight(40)
        
        self.send_button = QPushButton("Send")
        self.send_button.setMinimumHeight(40)
        self.send_button.setMinimumWidth(80)
        
        input_layout.addWidget(self.message_input)
        input_layout.addWidget(self.send_button)
        
        layout.addWidget(self.chat_scroll)
        layout.addLayout(input_layout)
        
        return chat_widget
        
    # The create_control_panel method is removed as it's no longer needed for the compact UI.
        
    def setup_connections(self):
        """Setup signal connections"""
        # UI connections
        self.send_button.clicked.connect(self.send_message)
        self.message_input.returnPressed.connect(self.send_message)
        
        # ChatBot connections
        self.chatbot.message_received.connect(self.add_message)
        self.chatbot.error_occurred.connect(self.show_error)
        self.chatbot.progress_updated.connect(self.update_progress)
        self.chatbot.status_updated.connect(self.update_status)
        
    def setup_styling(self):
        """Setup application styling"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #45a049;
            }
            
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            
            QLineEdit {
                border: 2px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                font-size: 12px;
            }
            
            QLineEdit:focus {
                border-color: #4CAF50;
            }
            
            QScrollArea {
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        
    def send_message(self):
        """Send a message to the chatbot"""
        message = self.message_input.text().strip()
        if not message:
            return
            
        # Add user message to chat
        self.add_message(message, "user")
        self.message_input.clear()
        
        # Process message based on content
        message_lower = message.lower()
        
        if message_lower in ['analyze', 'analyze workbook', 'structure']:
            self.analyze_workbook()
        elif message_lower in ['update', 'update trial balance', 'perform update']:
            self.start_update_process()
        elif message_lower in ['help', 'what can you do', 'commands']:
            self.show_help()
        else:
            # Send to chatbot for processing
            self.chatbot.set_request('chat', {'message': message})
            if not self.chatbot.isRunning():
                self.chatbot.start()
                
    def add_message(self, message, sender):
        """Add a message to the chat"""
        message_widget = ChatMessage(message, sender)
        
        # Insert before the stretch
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, message_widget)
        
        # Always auto-scroll
        QTimer.singleShot(100, self.scroll_to_bottom)
            
    def scroll_to_bottom(self):
        """Scroll chat to bottom"""
        scrollbar = self.chat_scroll.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
    # analyze_workbook, show_help, refresh_excel_status, and clear_chat are removed.
    # Their functionality is now handled exclusively through chat commands.
    def start_update_process(self):
        """Start the redesigned trial balance update process."""
        try:
            # 1. Check Excel status
            processor = TrialBalanceProcessor()
            excel_status = processor.get_excel_status()
            if not excel_status.get('workbook'):
                self.show_error("Excel is not running or no workbook is open.")
                return

            app = xw.apps.active
            wb = app.books.active
            sheet_names = [sheet.name for sheet in wb.sheets]

            # 2. Show dual sheet selection dialog
            setup_dialog = UpdateSetupDialog(sheet_names, self)
            if setup_dialog.exec() != QDialog.DialogCode.Accepted:
                return
            
            selection = setup_dialog.get_selection()
            if not selection:
                return # User selected same sheet for both, warning is shown in dialog

            to_update_sheet_name = selection['to_update_sheet']
            reference_sheet_name = selection['reference_sheet']

            # 3. Get column mapping for the 'to-update' sheet
            to_update_sheet = wb.sheets[to_update_sheet_name]
            to_update_headers = to_update_sheet.range('A1').expand('right').value
            mapping_dialog1 = ColumnMappingDialog(to_update_sheet_name, to_update_headers, self)
            if mapping_dialog1.exec() != QDialog.DialogCode.Accepted:
                return
            to_update_cols = mapping_dialog1.get_mapping()

            # 4. Get column mapping for the 'reference' sheet
            reference_sheet = wb.sheets[reference_sheet_name]
            reference_headers = reference_sheet.range('A1').expand('right').value
            mapping_dialog2 = ColumnMappingDialog(reference_sheet_name, reference_headers, self)
            if mapping_dialog2.exec() != QDialog.DialogCode.Accepted:
                return
            reference_cols = mapping_dialog2.get_mapping()

            # 5. Package data and start background task
            update_data = {
                'to_update_sheet': to_update_sheet_name,
                'reference_sheet': reference_sheet_name,
                'to_update_cols': to_update_cols,
                'reference_cols': reference_cols
            }

            if self.chatbot.isRunning():
                self.show_error("An update process is already running.")
                return

            self.chatbot.set_request('perform_update', update_data)
            self.add_message(f"Starting update process. Comparing '{to_update_sheet_name}' with '{reference_sheet_name}'.", "assistant")
            self.chatbot.start()

        except Exception as e:
            logger.error(f"Error starting update process: {str(e)}\n{traceback.format_exc()}")
            self.show_error(f"Failed to start update process: {str(e)}")
        
    def show_error(self, error_message):
        """Show error message"""
        self.add_message(f"❌ **Error:** {error_message}", "assistant")
        logger.error(error_message)
        
    def update_progress(self, value):
        """Update progress bar"""
        if value > 0 and value < 100:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(value)
        else:
            self.progress_bar.setVisible(False)
            
    def update_status(self, status):
        """Update status label"""
        self.status_label.setText(status)
        
        if status.lower() in ['ready', 'update complete', 'analysis complete']:
            self.progress_bar.setVisible(False)
            
    def closeEvent(self, event):
        """Handle application close"""
        if self.chatbot.isRunning():
            self.chatbot.terminate()
            self.chatbot.wait()
        event.accept()

def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Excel Trial Balance ChatBot")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("Excel Automation Tools")
    
    # Create and show main window
    window = ExcelChatBotGUI()
    window.show()
    
    # Initial Excel status check
    QTimer.singleShot(1000, window.refresh_excel_status)
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()