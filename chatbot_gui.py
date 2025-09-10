import sys
import json
import os
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QTextEdit, QLineEdit, QPushButton, QScrollArea, QFrame, QLabel,
    QMessageBox, QProgressBar, QDialog, QDialogButtonBox, QComboBox,
    QCheckBox, QSpinBox, QGroupBox, QGridLayout, QSplitter, QTabWidget,
    QFileDialog, QListWidget, QListWidgetItem, QTextBrowser
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt6.QtGui import QFont, QPixmap, QIcon
import xlwings as xw
from fuzzywuzzy import fuzz
from datetime import datetime

class ExcelChatBot(QThread):
    message_received = pyqtSignal(str, str)  # message, sender
    error_occurred = pyqtSignal(str)
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.api_key = os.getenv('OPENROUTER_API_KEY')
        self.api_url = "https://openrouter.ai/api/v1/chat/completions"
        self.conversation_history = []
        self.current_request = None
        self.is_processing = False
        
    def handle_excel_request(self, request_type, data=None):
        """Handle different types of Excel requests"""
        self.current_request = {
            'type': request_type,
            'data': data or {}
        }
        if not self.isRunning():
            self.start()
    
    def run(self):
        """Main thread execution"""
        if not self.current_request:
            return
            
        try:
            self.is_processing = True
            request_type = self.current_request['type']
            data = self.current_request['data']
            
            if request_type == 'analyze_structure':
                self.analyze_excel_structure()
            elif request_type == 'guide_update':
                self.guide_trial_balance_update(data.get('user_message', ''))
            elif request_type == 'chat':
                self.handle_chat_message(data.get('message', ''))
            elif request_type == 'perform_update':
                self.perform_trial_balance_update(data)
                
        except Exception as e:
            self.error_occurred.emit(f"An error occurred: {str(e)}")
        finally:
            self.is_processing = False
    
    def analyze_excel_structure(self):
        """Analyze the structure of the active Excel workbook"""
        try:
            self.status_updated.emit("Analyzing Excel structure...")
            
            # Check if Excel is running
            try:
                app = xw.App.active
                if not app.books:
                    self.message_received.emit(
                        "❌ No Excel workbook is currently open. Please open a workbook and try again.",
                        "assistant"
                    )
                    return
                    
                wb = app.books.active
                ws = wb.sheets.active
                
                # Get basic info
                workbook_name = wb.name
                sheet_name = ws.name
                
                # Get all sheet names
                sheet_names = [sheet.name for sheet in wb.sheets]
                
                # Get data range
                used_range = ws.used_range
                if used_range:
                    rows = used_range.shape[0]
                    cols = used_range.shape[1]
                    
                    # Get headers (first row)
                    headers = []
                    if rows > 0:
                        first_row = ws.range(f"A1:{chr(64 + cols)}1").value
                        if isinstance(first_row, list):
                            headers = [str(cell) if cell is not None else f"Column {i+1}" for i, cell in enumerate(first_row)]
                        else:
                            headers = [str(first_row) if first_row is not None else "Column 1"]
                else:
                    rows = cols = 0
                    headers = []
                
                # Format the analysis message
                message = f"📊 **Excel Workbook Analysis**\n\n"
                message += f"**Workbook:** {workbook_name}\n"
                message += f"**Active Sheet:** {sheet_name}\n\n"
                
                message += "**Available Sheets:**\n"
                for sheet in sheet_names:
                    message += f"• {sheet}\n"
                    
                message += f"\n**Data Range:** {rows} rows × {cols} columns\n"
                
                if headers:
                    message += "\n**Column Headers:**\n"
                    for i, header in enumerate(headers, 1):
                        message += f"{i}. {header}\n"
                
                self.message_received.emit(message, "assistant")
                self.status_updated.emit("Analysis complete")
                
            except Exception as e:
                self.message_received.emit(
                    f"❌ Error accessing Excel: {str(e)}\n\nPlease make sure Excel is running with a workbook open.",
                    "assistant"
                )
                
        except Exception as e:
            self.error_occurred.emit(f"Failed to analyze Excel structure: {str(e)}")
    
    def guide_trial_balance_update(self, user_message):
        """Guide the user through trial balance update process"""
        try:
            self.status_updated.emit("Processing your request...")
            
            # Check Excel status first
            excel_info = self.get_excel_status()
            
            # Prepare context for AI
            context = {
                'user_message': user_message,
                'excel_info': excel_info,
                'conversation_history': self.conversation_history[-5:]  # Last 5 messages for context
            }
            
            # Call OpenRouter API
            response = self.call_openrouter_api(context)
            
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
            self.error_occurred.emit(f"Failed to process request: {str(e)}")
    
    def handle_chat_message(self, message):
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
                self.handle_excel_request('analyze_structure')
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
                    response = self.call_openrouter_api(context)
                    if not response:
                        response = "I'm here to help with Excel trial balance operations. Try asking about 'analyze', 'update', or 'help'."
                else:
                    response = "I'm here to help with Excel trial balance operations. Try asking about 'analyze', 'update', or 'help'."
            
            self.message_received.emit(response, "assistant")
            self.status_updated.emit("Ready")
            
        except Exception as e:
            self.error_occurred.emit(f"Failed to process message: {str(e)}")
    
    def get_excel_status(self):
        """Get current Excel application status"""
        try:
            app = xw.App.active
            if not app.books:
                return {
                    'has_excel': True,
                    'has_workbook': False,
                    'workbook_name': None,
                    'sheet_names': [],
                    'active_sheet': None
                }
                
            wb = app.books.active
            return {
                'has_excel': True,
                'has_workbook': True,
                'workbook_name': wb.name,
                'sheet_names': [sheet.name for sheet in wb.sheets],
                'active_sheet': wb.sheets.active.name
            }
        except:
            return {
                'has_excel': False,
                'has_workbook': False,
                'workbook_name': None,
                'sheet_names': [],
                'active_sheet': None
            }
    
    def call_openrouter_api(self, context):
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
                return None
                
        except Exception as e:
            return None
    
    def perform_trial_balance_update(self, update_data):
        """Perform the actual trial balance update"""
        try:
            self.status_updated.emit("Performing trial balance update...")
            self.progress_updated.emit(10)
            
            # Get Excel app and workbook
            app = xw.App.active
            wb = app.books.active
            
            # Extract update parameters
            sheet_name = update_data.get('sheet_name')
            column_mapping = update_data.get('column_mapping', {})
            updates = update_data.get('updates', [])
            
            if not updates:
                self.error_occurred.emit("No updates to perform")
                return
                
            # Get the target sheet
            if sheet_name and sheet_name in [s.name for s in wb.sheets]:
                ws = wb.sheets[sheet_name]
            else:
                ws = wb.sheets.active
                
            self.progress_updated.emit(30)
            
            # Perform updates
            updated_accounts = []
            failed_accounts = []
            
            for update in updates:
                try:
                    account_name = update.get('account')
                    new_amount = update.get('amount')
                    row_number = update.get('row')
                    
                    if row_number and new_amount is not None:
                        # Update the amount in the specified row
                        amount_col = column_mapping.get('amount', 'C')  # Default to column C
                        cell_address = f"{amount_col}{row_number}"
                        ws.range(cell_address).value = new_amount
                        updated_accounts.append(account_name)
                    else:
                        failed_accounts.append(account_name)
                        
                except Exception as e:
                    failed_accounts.append(f"{account_name} (Error: {str(e)})")
            
            self.progress_updated.emit(80)
            
            # Save the workbook
            wb.save()
            
            # Report results
            message = f"✅ **Update Successful!**\n\n"
            message += f"**Updated {len(updated_accounts)} accounts:**\n"
            for account in updated_accounts:
                message += f"• {account}\n"
                
            if failed_accounts:
                message += f"\n**⚠️ Failed to update {len(failed_accounts)} accounts:**\n"
                for account in failed_accounts:
                    message += f"• {account}\n"
                    
            self.message_received.emit(message, "assistant")
            self.progress_updated.emit(100)
            self.status_updated.emit("Update complete")
            
        except Exception as e:
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
        
        # Chat title
        title_label = QLabel("💬 Chat with Assistant")
        title_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        title_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Chat messages area
        self.chat_scroll = QScrollArea()
        self.chat_scroll.setWidgetResizable(True)
        self.chat_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.chat_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # Chat messages container
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.addStretch()
        self.chat_scroll.setWidget(self.chat_container)
        
        layout.addWidget(self.chat_scroll)
        
        # Input area
        input_layout = QHBoxLayout()
        
        self.message_input = QLineEdit()
        self.message_input.setPlaceholderText("Type your message here... (or try 'help', 'analyze', 'update')")
        self.message_input.setFont(QFont("Arial", 10))
        
        self.send_button = QPushButton("Send")
        self.send_button.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.send_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        
        input_layout.addWidget(self.message_input)
        input_layout.addWidget(self.send_button)
        
        layout.addLayout(input_layout)
        
        return chat_widget
    
    def create_control_panel(self):
        """Create the control panel"""
        control_widget = QWidget()
        layout = QVBoxLayout(control_widget)
        
        # Control title
        title_label = QLabel("🎛️ Controls & Info")
        title_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        title_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Quick actions
        actions_group = QGroupBox("Quick Actions")
        actions_layout = QVBoxLayout(actions_group)
        
        self.analyze_button = QPushButton("📊 Analyze Excel")
        self.analyze_button.setToolTip("Analyze the structure of your current Excel workbook")
        
        self.update_button = QPushButton("🔄 Start Update Process")
        self.update_button.setToolTip("Begin the trial balance update process")
        
        self.clear_button = QPushButton("🗑️ Clear Chat")
        self.clear_button.setToolTip("Clear all chat messages")
        
        actions_layout.addWidget(self.analyze_button)
        actions_layout.addWidget(self.update_button)
        actions_layout.addWidget(self.clear_button)
        
        layout.addWidget(actions_group)
        
        # Excel status
        status_group = QGroupBox("Excel Status")
        status_layout = QVBoxLayout(status_group)
        
        self.excel_status_label = QLabel("Checking Excel status...")
        self.excel_status_label.setWordWrap(True)
        self.excel_status_label.setStyleSheet("padding: 5px; background-color: #f9f9f9; border-radius: 3px;")
        
        self.refresh_status_button = QPushButton("🔄 Refresh Status")
        
        status_layout.addWidget(self.excel_status_label)
        status_layout.addWidget(self.refresh_status_button)
        
        layout.addWidget(status_group)
        
        # Help section
        help_group = QGroupBox("Help & Tips")
        help_layout = QVBoxLayout(help_group)
        
        help_text = QTextBrowser()
        help_text.setMaximumHeight(150)
        help_text.setMarkdown("""
**Quick Commands:**
• `help` - Show available commands
• `analyze` - Analyze Excel structure
• `update` - Start update process

**Tips:**
• Make sure Excel is open with your trial balance data
• Use clear column headers for best results
• The assistant can guide you through each step
        """)
        
        help_layout.addWidget(help_text)
        layout.addWidget(help_group)
        
        layout.addStretch()
        
        return control_widget
    
    def setup_connections(self):
        """Setup signal connections"""
        # Chat connections
        self.send_button.clicked.connect(self.send_message)
        self.message_input.returnPressed.connect(self.send_message)
        
        # Control panel connections
        self.analyze_button.clicked.connect(self.analyze_excel)
        self.update_button.clicked.connect(self.start_update_process)
        self.clear_button.clicked.connect(self.clear_chat)
        self.refresh_status_button.clicked.connect(self.refresh_excel_status)
        
        # ChatBot connections
        self.chatbot.message_received.connect(self.add_message)
        self.chatbot.error_occurred.connect(self.show_error)
        self.chatbot.progress_updated.connect(self.update_progress)
        self.chatbot.status_updated.connect(self.update_status)
        
        # Timer for periodic status updates
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.refresh_excel_status)
        self.status_timer.start(5000)  # Update every 5 seconds
        
        # Initial status check
        QTimer.singleShot(1000, self.refresh_excel_status)
    
    def send_message(self):
        """Send a message to the chatbot"""
        message = self.message_input.text().strip()
        if not message:
            return
            
        # Add user message to chat
        self.add_message(message, "user")
        self.message_input.clear()
        
        # Handle special commands
        message_lower = message.lower()
        
        if message_lower in ['clear', 'clear chat']:
            self.clear_chat()
            return
        elif message_lower in ['status', 'excel status']:
            self.refresh_excel_status()
            return
        elif message_lower in ['help', 'commands']:
            self.chatbot.handle_excel_request('chat', {'message': 'help'})
            return
        elif message_lower in ['analyze', 'analyze excel']:
            self.analyze_excel()
            return
        elif message_lower in ['update', 'start update']:
            self.start_update_process()
            return
        
        # Send to chatbot for processing
        self.chatbot.handle_excel_request('chat', {'message': message})
    
    def add_message(self, message, sender):
        """Add a message to the chat"""
        chat_message = ChatMessage(message, sender)
        
        # Insert before the stretch
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, chat_message)
        
        # Scroll to bottom
        QTimer.singleShot(100, self.scroll_to_bottom)
    
    def scroll_to_bottom(self):
        """Scroll chat to bottom"""
        scrollbar = self.chat_scroll.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def clear_chat(self):
        """Clear all chat messages"""
        # Remove all message widgets except the stretch
        for i in reversed(range(self.chat_layout.count() - 1)):
            child = self.chat_layout.itemAt(i).widget()
            if child:
                child.setParent(None)
        
        # Add welcome message back
        self.add_message(
            "👋 Chat cleared! I'm ready to help with your Excel trial balance operations.",
            "assistant"
        )
    
    def analyze_excel(self):
        """Analyze Excel structure"""
        self.chatbot.handle_excel_request('analyze_structure')
    
    def start_update_process(self):
        """Start the trial balance update process"""
        self.add_message("Starting trial balance update process...", "user")
        self.chatbot.handle_excel_request('chat', {
            'message': 'I want to update my trial balance. Please guide me through the process.'
        })
    
    def refresh_excel_status(self):
        """Refresh Excel status display"""
        try:
            status = self.chatbot.get_excel_status()
            
            if not status['has_excel']:
                status_text = "❌ Excel not detected"
                color = "#ffebee"
            elif not status['has_workbook']:
                status_text = "⚠️ Excel running, no workbook open"
                color = "#fff3e0"
            else:
                status_text = f"✅ Excel ready\nWorkbook: {status['workbook_name']}\nActive Sheet: {status['active_sheet']}"
                color = "#e8f5e8"
                
            self.excel_status_label.setText(status_text)
            self.excel_status_label.setStyleSheet(f"padding: 5px; background-color: {color}; border-radius: 3px;")
            
        except Exception as e:
            self.excel_status_label.setText(f"❌ Error checking Excel: {str(e)}")
            self.excel_status_label.setStyleSheet("padding: 5px; background-color: #ffebee; border-radius: 3px;")
    
    def show_error(self, error_message):
        """Show error message"""
        self.add_message(f"❌ **Error:** {error_message}", "assistant")
        QMessageBox.warning(self, "Error", error_message)
    
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
    
    def show_table_data(self, data, title="Table Data"):
        """Show table data in a dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setGeometry(200, 200, 800, 600)
        
        layout = QVBoxLayout(dialog)
        
        # Create table widget
        table_text = QTextBrowser()
        
        # Format data as markdown table
        if data and len(data) > 0:
            # Get headers
            headers = list(data[0].keys()) if isinstance(data[0], dict) else [f"Column {i+1}" for i in range(len(data[0]))]
            
            # Create markdown table
            markdown = "| " + " | ".join(headers) + " |\n"
            markdown += "| " + " | ".join(["---"] * len(headers)) + " |\n"
            
            for row in data[:50]:  # Limit to first 50 rows
                if isinstance(row, dict):
                    values = [str(row.get(header, "")) for header in headers]
                else:
                    values = [str(cell) for cell in row]
                markdown += "| " + " | ".join(values) + " |\n"
                
            if len(data) > 50:
                markdown += f"\n*... and {len(data) - 50} more rows*"
                
            table_text.setMarkdown(markdown)
        else:
            table_text.setText("No data to display")
            
        layout.addWidget(table_text)
        
        # Close button
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.exec()
    
    def show_column_preview(self, columns, title="Column Preview"):
        """Show column preview dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setGeometry(200, 200, 600, 400)
        
        layout = QVBoxLayout(dialog)
        
        # Info label
        info_label = QLabel(f"Found {len(columns)} columns:")
        info_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        layout.addWidget(info_label)
        
        # Column list
        column_list = QListWidget()
        for i, col in enumerate(columns, 1):
            column_list.addItem(f"{i}. {col}")
        layout.addWidget(column_list)
        
        # Close button
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.exec()
    
    def perform_trial_balance_update(self, update_data):
        """Perform trial balance update with user confirmation"""
        # Show confirmation dialog
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Question)
        msg.setWindowTitle("Confirm Update")
        msg.setText("Are you sure you want to perform the trial balance update?")
        
        details = f"Sheet: {update_data.get('sheet_name', 'Active sheet')}\n"
        details += f"Updates: {len(update_data.get('updates', []))} accounts\n"
        details += "\nThis will modify your Excel workbook. Make sure you have a backup!"
        msg.setDetailedText(details)
        
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.No)
        
        if msg.exec() == QMessageBox.StandardButton.Yes:
            self.chatbot.handle_excel_request('perform_update', update_data)
        else:
            self.add_message("Update cancelled by user.", "assistant")
    
    def show_interactive_dialog(self, dialog_type, data=None):
        """Show interactive dialog for user input"""
        if dialog_type == 'sheet_selection':
            return self.show_sheet_selection_dialog(data)
        elif dialog_type == 'column_mapping':
            return self.show_column_mapping_dialog(data)
        elif dialog_type == 'preview_changes':
            return self.show_preview_changes_dialog(data)
        
        return None
    
    def show_sheet_selection_dialog(self, sheets):
        """Show sheet selection dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Sheet")
        dialog.setGeometry(300, 300, 400, 300)
        
        layout = QVBoxLayout(dialog)
        
        # Info label
        info_label = QLabel("Select the sheet containing your trial balance data:")
        layout.addWidget(info_label)
        
        # Sheet list
        sheet_combo = QComboBox()
        sheet_combo.addItems(sheets)
        layout.addWidget(sheet_combo)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            return sheet_combo.currentText()
        return None
    
    def show_column_mapping_dialog(self, columns):
        """Show column mapping dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Map Columns")
        dialog.setGeometry(250, 250, 500, 400)
        
        layout = QVBoxLayout(dialog)
        
        # Info label
        info_label = QLabel("Map your columns to trial balance fields:")
        layout.addWidget(info_label)
        
        # Mapping grid
        grid = QGridLayout()
        
        # Account column
        grid.addWidget(QLabel("Account Name:"), 0, 0)
        account_combo = QComboBox()
        account_combo.addItems([''] + columns)
        grid.addWidget(account_combo, 0, 1)
        
        # Debit column
        grid.addWidget(QLabel("Debit Amount:"), 1, 0)
        debit_combo = QComboBox()
        debit_combo.addItems([''] + columns)
        grid.addWidget(debit_combo, 1, 1)
        
        # Credit column
        grid.addWidget(QLabel("Credit Amount:"), 2, 0)
        credit_combo = QComboBox()
        credit_combo.addItems([''] + columns)
        grid.addWidget(credit_combo, 2, 1)
        
        # Balance column (optional)
        grid.addWidget(QLabel("Balance (optional):"), 3, 0)
        balance_combo = QComboBox()
        balance_combo.addItems([''] + columns)
        grid.addWidget(balance_combo, 3, 1)
        
        layout.addLayout(grid)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            return {
                'account': account_combo.currentText(),
                'debit': debit_combo.currentText(),
                'credit': credit_combo.currentText(),
                'balance': balance_combo.currentText()
            }
        return None
    
    def show_preview_changes_dialog(self, changes):
        """Show preview of changes dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Preview Changes")
        dialog.setGeometry(200, 200, 700, 500)
        
        layout = QVBoxLayout(dialog)
        
        # Info label
        info_label = QLabel(f"Preview of {len(changes)} proposed changes:")
        info_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        layout.addWidget(info_label)
        
        # Changes table
        changes_text = QTextBrowser()
        
        # Format changes as markdown table
        markdown = "| Account | Current | Proposed | Change |\n"
        markdown += "| --- | --- | --- | --- |\n"
        
        for change in changes:
            account = change.get('account', 'Unknown')
            current = change.get('current_value', 'N/A')
            proposed = change.get('proposed_value', 'N/A')
            diff = change.get('difference', 'N/A')
            markdown += f"| {account} | {current} | {proposed} | {diff} |\n"
            
        changes_text.setMarkdown(markdown)
        layout.addWidget(changes_text)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        return dialog.exec() == QDialog.DialogCode.Accepted
    
    def autonomous_mode(self):
        """Run autonomous mode to automatically detect and update trial balance"""
        try:
            self.add_message("🤖 Starting autonomous mode...", "assistant")
            self.update_status("Running autonomous mode...")
            
            # Step 1: Detect Excel and sheets
            excel_status = self.chatbot.get_excel_status()
            if not excel_status['has_excel'] or not excel_status['has_workbook']:
                self.add_message("❌ Please ensure Excel is running with a workbook open.", "assistant")
                return
            
            # Step 2: Auto-detect trial balance sheets
            trial_balance_sheets = self.auto_detect_sheets(excel_status['sheet_names'])
            
            if not trial_balance_sheets:
                self.add_message("❌ No trial balance sheets detected. Please ensure your workbook contains trial balance data.", "assistant")
                return
            
            self.add_message(f"📊 Detected {len(trial_balance_sheets)} potential trial balance sheet(s): {', '.join(trial_balance_sheets)}", "assistant")
            
            # Step 3: Process each sheet
            for sheet_name in trial_balance_sheets:
                self.add_message(f"🔍 Analyzing sheet: {sheet_name}", "assistant")
                
                # Auto-detect columns
                column_mapping = self.auto_detect_columns(sheet_name)
                
                if not column_mapping:
                    self.add_message(f"⚠️ Could not detect trial balance columns in sheet '{sheet_name}'. Skipping...", "assistant")
                    continue
                
                self.add_message(f"✅ Column mapping detected for '{sheet_name}': {column_mapping}", "assistant")
                
                # Preview changes (simplified for autonomous mode)
                self.add_message(f"📋 Sheet '{sheet_name}' is ready for updates. Column mapping: {column_mapping}", "assistant")
            
            self.add_message("🎉 Autonomous analysis complete! Use the update commands to proceed with modifications.", "assistant")
            self.update_status("Autonomous mode complete")
            
        except Exception as e:
            self.add_message(f"❌ Error in autonomous mode: {str(e)}", "assistant")
            self.update_status("Ready")
    
    def auto_detect_sheets(self, sheet_names):
        """Auto-detect sheets that likely contain trial balance data"""
        trial_balance_keywords = [
            'trial', 'balance', 'tb', 'trial balance', 'trialbalance',
            'accounts', 'ledger', 'general ledger', 'gl', 'chart of accounts'
        ]
        
        detected_sheets = []
        
        for sheet_name in sheet_names:
            sheet_lower = sheet_name.lower()
            
            # Check for keywords
            for keyword in trial_balance_keywords:
                if keyword in sheet_lower:
                    detected_sheets.append(sheet_name)
                    break
            
            # Also check sheet structure (if it has typical trial balance columns)
            try:
                app = xw.App.active
                wb = app.books.active
                ws = wb.sheets[sheet_name]
                
                # Get first few rows to check for typical headers
                if ws.used_range and ws.used_range.shape[0] > 0:
                    first_row = ws.range(f"A1:{chr(64 + min(10, ws.used_range.shape[1]))}1").value
                    if isinstance(first_row, list):
                        headers = [str(cell).lower() if cell else '' for cell in first_row]
                    else:
                        headers = [str(first_row).lower() if first_row else '']
                    
                    # Check for typical trial balance headers
                    account_found = any('account' in h or 'name' in h for h in headers)
                    amount_found = any(word in h for h in headers for word in ['debit', 'credit', 'balance', 'amount'])
                    
                    if account_found and amount_found and sheet_name not in detected_sheets:
                        detected_sheets.append(sheet_name)
                        
            except Exception:
                continue  # Skip sheets that can't be analyzed
        
        return detected_sheets
    
    def auto_detect_columns(self, sheet_name):
        """Auto-detect column mapping for a trial balance sheet"""
        try:
            app = xw.App.active
            wb = app.books.active
            ws = wb.sheets[sheet_name]
            
            if not ws.used_range or ws.used_range.shape[0] == 0:
                return None
            
            # Get headers (try first few rows)
            headers = []
            for row in range(1, min(4, ws.used_range.shape[0] + 1)):
                row_data = ws.range(f"A{row}:{chr(64 + min(20, ws.used_range.shape[1]))}{row}").value
                if isinstance(row_data, list):
                    potential_headers = [str(cell) if cell else '' for cell in row_data]
                else:
                    potential_headers = [str(row_data) if row_data else '']
                
                # Check if this looks like a header row
                if any(word in h.lower() for h in potential_headers for word in ['account', 'debit', 'credit', 'balance']):
                    headers = potential_headers
                    break
            
            if not headers:
                return None
            
            # Map columns based on keywords
            column_mapping = {}
            
            for i, header in enumerate(headers):
                header_lower = header.lower()
                col_letter = chr(65 + i)  # A, B, C, etc.
                
                # Account column
                if any(word in header_lower for word in ['account', 'name', 'description']):
                    if 'account' not in column_mapping:
                        column_mapping['account'] = col_letter
                
                # Debit column
                elif 'debit' in header_lower:
                    column_mapping['debit'] = col_letter
                
                # Credit column
                elif 'credit' in header_lower:
                    column_mapping['credit'] = col_letter
                
                # Balance column
                elif 'balance' in header_lower:
                    column_mapping['balance'] = col_letter
                
                # Amount column (generic)
                elif 'amount' in header_lower and 'debit' not in column_mapping and 'credit' not in column_mapping:
                    column_mapping['amount'] = col_letter
            
            # Fallback: if we found account but no debit/credit, look for numeric columns
            if 'account' in column_mapping and 'debit' not in column_mapping and 'credit' not in column_mapping:
                # Find numeric columns after the account column
                account_col_index = ord(column_mapping['account']) - 65
                
                for i in range(account_col_index + 1, len(headers)):
                    if i < len(headers):
                        col_letter = chr(65 + i)
                        # Check if this column contains numeric data
                        try:
                            sample_value = ws.range(f"{col_letter}2").value
                            if isinstance(sample_value, (int, float)):
                                if 'debit' not in column_mapping:
                                    column_mapping['debit'] = col_letter
                                elif 'credit' not in column_mapping:
                                    column_mapping['credit'] = col_letter
                                    break
                        except:
                            continue
            
            # Return mapping if we found at least account column
            if 'account' in column_mapping:
                return column_mapping
            
            return None
            
        except Exception as e:
            return None

def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Excel Trial Balance ChatBot")
    app.setApplicationVersion("1.0")
    
    # Load API key from environment
    from dotenv import load_dotenv
    load_dotenv()
    
    # Create and show the main window
    window = ExcelChatBotGUI()
    window.show()
    
    # Run the application
    sys.exit(app.exec())

if __name__ == "__main__":
    main()