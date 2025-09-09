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
from interactive_dialogs import SheetSelectionDialog, ColumnMappingDialog, ProgressDialog, PreviewDialog

# Load environment variables
load_dotenv()

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
        """Perform the actual trial balance update"""
        try:
            self.status_updated.emit("Performing trial balance update...")
            self.progress_updated.emit(10)
            
            # Extract update parameters
            sheet_name = update_data.get('sheet_name')
            column_mapping = update_data.get('column_mapping', {})
            updates = update_data.get('updates', [])
            
            if not updates:
                self.error_occurred.emit("No updates to perform")
                return
                
            self.progress_updated.emit(30)
            
            # Perform the update using the processor
            result = self.processor.update_trial_balance(
                sheet_name=sheet_name,
                column_mapping=column_mapping,
                updates=updates
            )
            
            self.progress_updated.emit(80)
            
            if result['success']:
                message = f"✅ **Update Successful!**\n\n"
                message += f"**Updated {result['updated_count']} accounts:**\n"
                for account in result['updated_accounts']:
                    message += f"• {account}\n"
                    
                if result['failed_accounts']:
                    message += f"\n**⚠️ Failed to update {len(result['failed_accounts'])} accounts:**\n"
                    for account in result['failed_accounts']:
                        message += f"• {account}\n"
                        
                self.message_received.emit(message, "assistant")
            else:
                self.error_occurred.emit(f"Update failed: {result.get('error', 'Unknown error')}")
                
            self.progress_updated.emit(100)
            self.status_updated.emit("Update complete")
            
        except Exception as e:
            logger.error(f"Error performing update: {str(e)}")
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

class ExcelChatBotGUI(QMainWindow):
    """Main GUI application for Excel ChatBot"""
    
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
        """Setup the main user interface"""
        self.setWindowTitle("Excel Trial Balance ChatBot")
        self.setGeometry(100, 100, 1000, 700)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel - Chat
        chat_widget = self.create_chat_panel()
        splitter.addWidget(chat_widget)
        
        # Right panel - Controls and info
        control_widget = self.create_control_panel()
        splitter.addWidget(control_widget)
        
        # Set splitter proportions
        splitter.setSizes([700, 300])
        
        main_layout.addWidget(splitter)
        
        # Status bar
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready")
        
        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
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
        
    def create_control_panel(self):
        """Create the control panel"""
        control_widget = QWidget()
        layout = QVBoxLayout(control_widget)
        
        # Quick actions
        actions_group = QGroupBox("Quick Actions")
        actions_layout = QVBoxLayout(actions_group)
        
        self.analyze_button = QPushButton("📊 Analyze Workbook")
        self.update_button = QPushButton("🔄 Update Trial Balance")
        self.help_button = QPushButton("❓ Help")
        
        actions_layout.addWidget(self.analyze_button)
        actions_layout.addWidget(self.update_button)
        actions_layout.addWidget(self.help_button)
        
        # Excel status
        status_group = QGroupBox("Excel Status")
        status_layout = QVBoxLayout(status_group)
        
        self.excel_status_label = QLabel("Checking Excel status...")
        self.excel_status_label.setWordWrap(True)
        
        self.refresh_status_button = QPushButton("🔄 Refresh Status")
        
        status_layout.addWidget(self.excel_status_label)
        status_layout.addWidget(self.refresh_status_button)
        
        # Settings
        settings_group = QGroupBox("Settings")
        settings_layout = QVBoxLayout(settings_group)
        
        self.auto_scroll_checkbox = QCheckBox("Auto-scroll chat")
        self.auto_scroll_checkbox.setChecked(True)
        
        self.clear_chat_button = QPushButton("🗑️ Clear Chat")
        
        settings_layout.addWidget(self.auto_scroll_checkbox)
        settings_layout.addWidget(self.clear_chat_button)
        
        # Add all groups to layout
        layout.addWidget(actions_group)
        layout.addWidget(status_group)
        layout.addWidget(settings_group)
        layout.addStretch()
        
        return control_widget
        
    def setup_connections(self):
        """Setup signal connections"""
        # UI connections
        self.send_button.clicked.connect(self.send_message)
        self.message_input.returnPressed.connect(self.send_message)
        self.analyze_button.clicked.connect(self.analyze_workbook)
        self.update_button.clicked.connect(self.start_update_process)
        self.help_button.clicked.connect(self.show_help)
        self.refresh_status_button.clicked.connect(self.refresh_excel_status)
        self.clear_chat_button.clicked.connect(self.clear_chat)
        
        # ChatBot connections
        self.chatbot.message_received.connect(self.add_message)
        self.chatbot.error_occurred.connect(self.show_error)
        self.chatbot.progress_updated.connect(self.update_progress)
        self.chatbot.status_updated.connect(self.update_status)
        
        # Timer for periodic status updates
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.refresh_excel_status)
        self.status_timer.start(10000)  # Update every 10 seconds
        
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
        
        # Auto-scroll if enabled
        if self.auto_scroll_checkbox.isChecked():
            QTimer.singleShot(100, self.scroll_to_bottom)
            
    def scroll_to_bottom(self):
        """Scroll chat to bottom"""
        scrollbar = self.chat_scroll.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
    def analyze_workbook(self):
        """Analyze the current Excel workbook"""
        if self.chatbot.isRunning():
            self.show_error("Please wait for the current operation to complete.")
            return
            
        self.chatbot.set_request('analyze_structure')
        self.chatbot.start()
        
    def start_update_process(self):
        """Start the trial balance update process"""
        try:
            # First check Excel status
            processor = TrialBalanceProcessor()
            excel_status = processor.get_excel_status()
            
            if not excel_status['has_excel']:
                self.show_error("Excel is not running. Please open Excel first.")
                return
                
            if not excel_status['has_workbook']:
                self.show_error("No Excel workbook is open. Please open a workbook first.")
                return
                
            # Show sheet selection dialog
            app = xw.apps.active
            wb = app.books.active
            sheet_names = [sheet.name for sheet in wb.sheets]
            
            dialog = SheetSelectionDialog(sheet_names, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected_sheet = dialog.get_selected_sheet()
                
                # Analyze the selected sheet
                structure = processor.analyze_structure(selected_sheet)
                
                # Show column mapping dialog
                mapping_dialog = ColumnMappingDialog(structure['headers'], self)
                if mapping_dialog.exec() == QDialog.DialogCode.Accepted:
                    column_mapping = mapping_dialog.get_mapping()
                    
                    # Extract trial balance data
                    trial_balance_data = processor.extract_trial_balance_data(
                        selected_sheet, column_mapping
                    )
                    
                    if not trial_balance_data:
                        self.show_error("No trial balance data found in the selected sheet.")
                        return
                        
                    # Show preview dialog
                    preview_dialog = PreviewDialog(trial_balance_data, self)
                    if preview_dialog.exec() == QDialog.DialogCode.Accepted:
                        updates = preview_dialog.get_updates()
                        
                        if updates:
                            # Perform the update
                            update_data = {
                                'sheet_name': selected_sheet,
                                'column_mapping': column_mapping,
                                'updates': updates
                            }
                            
                            self.chatbot.set_request('perform_update', update_data)
                            if not self.chatbot.isRunning():
                                self.chatbot.start()
                        else:
                            self.add_message("No updates were selected.", "assistant")
                            
        except Exception as e:
            logger.error(f"Error starting update process: {str(e)}")
            self.show_error(f"Failed to start update process: {str(e)}")
            
    def show_help(self):
        """Show help information"""
        help_message = """🤖 **Excel Trial Balance Assistant Help**

**Quick Actions:**
• **Analyze Workbook** - Analyze the structure of your Excel workbook
• **Update Trial Balance** - Start the guided update process
• **Help** - Show this help message

**Chat Commands:**
• Type `analyze` to analyze your workbook
• Type `update` to start updating trial balance
• Type `help` to see available commands
• Ask questions about Excel operations

**Update Process:**
1. Select the sheet containing trial balance data
2. Map columns (Account, Debit, Credit)
3. Preview proposed changes
4. Confirm and execute updates

**Tips:**
• Make sure Excel is running with your workbook open
• Ensure your trial balance data has clear column headers
• Review all changes before confirming updates
• Use the refresh button to update Excel status

**Troubleshooting:**
• If Excel status shows as disconnected, try refreshing
• Ensure your workbook has the expected trial balance format
• Check that column headers match expected patterns"""
        
        self.add_message(help_message, "assistant")
        
    def refresh_excel_status(self):
        """Refresh Excel connection status"""
        try:
            processor = TrialBalanceProcessor()
            status = processor.get_excel_status()
            
            status_text = "📊 **Excel Status**\n\n"
            
            if status['has_excel']:
                status_text += "✅ Excel: Connected\n"
                
                if status['has_workbook']:
                    status_text += f"✅ Workbook: {status['workbook_name']}\n"
                    status_text += f"📄 Active Sheet: {status['active_sheet']}\n"
                else:
                    status_text += "❌ Workbook: None open\n"
            else:
                status_text += "❌ Excel: Not running\n"
                
            self.excel_status_label.setText(status_text)
            
        except Exception as e:
            self.excel_status_label.setText(f"❌ Status check failed: {str(e)}")
            
    def clear_chat(self):
        """Clear the chat history"""
        # Remove all message widgets except the stretch
        while self.chat_layout.count() > 1:
            child = self.chat_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
                
        # Add welcome message back
        self.add_message(
            "👋 Chat cleared! I'm ready to help with your Excel trial balance operations.",
            "assistant"
        )
        
    def show_error(self, error_message):
        """Show error message"""
        self.add_message(f"❌ **Error:** {error_message}", "assistant")
        logger.error(error_message)
        
    def update_progress(self, value):
        """Update progress bar"""
        if value > 0:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(value)
        else:
            self.progress_bar.setVisible(False)
            
    def update_status(self, status):
        """Update status bar"""
        self.status_bar.showMessage(status)
        
        if status.lower() in ['ready', 'complete']:
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