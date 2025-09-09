# Excel Trial Balance Autonomous Agent 📊

An intelligent autonomous agent that automatically updates Excel working papers leadsheet accounts using AI-powered fuzzy matching. The agent detects open Excel files, performs interactive setup, and updates trial balance data while ignoring empty and bold cells.

## 🚀 Features

- **🤖 AI-Powered**: Uses OpenRouter API with Qwen 3 Coder model for intelligent processing
- **📂 Auto-Detection**: Automatically detects and connects to open Excel files using xlwings
- **🎯 Fuzzy Matching**: Performs intelligent account matching with 80% similarity threshold
- **🛡️ Smart Filtering**: Ignores empty cells and bold cells as per business rules
- **📊 Interactive Setup**: Guides you through sheet and column mapping with previews
- **🔄 Complete Updates**: Updates both current year and prior year amounts
- **➕ New Account Detection**: Automatically identifies and adds new accounts from reference sheet
- **📝 Comprehensive Logging**: Detailed logging with colored console output

## 📋 Prerequisites

1. **Python 3.8+** installed on your system
2. **Microsoft Excel** with xlwings add-in (automatically installed with xlwings package)
3. **OpenRouter API Key** - Get one from [OpenRouter](https://openrouter.ai/keys)
4. **Excel file** with trial balance sheets open before running the agent

## 🛠️ Installation

1. **Clone or download** this project to your local machine

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Setup environment variables**:
   ```bash
   # Copy the example environment file
   copy .env.example .env
   
   # Edit .env file and add your OpenRouter API key
   # OPENROUTER_API_KEY=your_actual_api_key_here
   ```

4. **Install xlwings Excel add-in** (if not already installed):
   ```bash
   xlwings addin install
   ```

## 🎯 Usage

### Step 1: Prepare Your Excel File

1. Open your Excel file containing the trial balance sheets
2. Ensure you have:
   - **Sheet with data to be updated** (contains incorrect amounts)
   - **Sheet with correct reference data** (contains correct amounts)

### Step 2: Run the Agent

```bash
python main.py
```

### Step 3: Interactive Setup

The agent will guide you through an interactive setup process:

#### 🟩 Sheet Selection
- **To-Update Sheet**: Select the sheet with data that needs to be updated
- **Reference Sheet**: Select the sheet with correct amounts

#### 🟦 Column Mapping

For each sheet, you'll specify:
- **Account Names Column**: Column containing account names (e.g., "A")
- **Current Year Column**: Column with current year amounts (e.g., "B")
- **Prior Year Column**: Column with prior year amounts (e.g., "C")

The agent will show you previews of the first 3 identified accounts and amounts to confirm your selections.

### Step 4: Automated Processing

The agent will:
1. **Extract Data**: Get all non-empty, non-bold account data from both sheets
2. **Fuzzy Match**: Match accounts between sheets using 80% similarity threshold
3. **Update Amounts**: Update current year and prior year amounts for matched accounts
4. **Add New Accounts**: Identify and add accounts that exist in reference sheet but not in to-update sheet

## 📊 Example Workflow

```
=== Excel Trial Balance Agent - Interactive Setup ===

Connected to Excel file: Trial_Balance_2024.xlsx

Available sheets: TB_Draft, TB_Final, Adjustments

✅ Which sheet contains the data to be updated (the one with incorrect amounts)?
Enter sheet name: TB_Draft

✅ Which sheet contains the correct amounts (reference data)?
Enter sheet name: TB_Final

🟩 Setting up 'TB_Draft' sheet mappings:
Available columns: {'A': 'Account Name', 'B': 'Current Year', 'C': 'Prior Year'...}

Which column contains the account names?
Enter column letter (e.g., A, B, C): A

First 3 identified accounts:
  1. Cash and Cash Equivalents
  2. Accounts Receivable
  3. Inventory

...

=== Starting Update Process ===

📊 Extracting account data...
Found 45 accounts in 'TB_Draft' sheet
Found 48 accounts in 'TB_Final' sheet

🔍 Performing fuzzy matching (80% threshold)...
Match found: 'Cash and Cash Equivalents' -> 'Cash & Cash Equivalents' (Score: 95%)
...
Found 42 matches out of 45 accounts

💰 Updating amounts...
Updated row 2: Current=125000, Prior=98000
...
Successfully updated 42 accounts

🆕 Finding new accounts...
Found 3 new accounts to add
Added new account 'Deferred Tax Assets' at row 46
...
Successfully added 3 new accounts

✅ Update process completed successfully!
Summary:
  - Updated: 42 existing accounts
  - Added: 3 new accounts
  - Total processed: 45 accounts
```

## ⚙️ Configuration

You can customize the agent behavior by modifying the `.env` file:

```env
# Fuzzy matching threshold (0-100)
FUZZY_MATCH_THRESHOLD=80

# Maximum columns to check for headers
MAX_COLUMNS_TO_CHECK=20

# Starting row for data (usually 2, as row 1 contains headers)
START_ROW=2

# Logging level (DEBUG, INFO, WARNING, ERROR)
LOG_LEVEL=INFO
```

## 🔧 Troubleshooting

### Common Issues

1. **"No active Excel workbook found"**
   - Ensure Excel is open with your file loaded
   - Try opening Excel as Administrator

2. **"OpenRouter API key not set"**
   - Check your `.env` file exists and contains the correct API key
   - Ensure the key starts with `sk-or-`

3. **"Error connecting to Excel"**
   - Install xlwings add-in: `xlwings addin install`
   - Restart Excel after installing xlwings

4. **"Invalid sheet names"**
   - Check sheet names for exact spelling and case sensitivity
   - Ensure sheets exist in the active workbook

### Debug Mode

For detailed debugging, set `LOG_LEVEL=DEBUG` in your `.env` file and check the `excel_agent.log` file.

## 📁 Project Structure

```
Autonomous agents/
├── main.py                 # Main execution script
├── excel_agent.py          # Core agent implementation
├── requirements.txt        # Python dependencies
├── .env.example           # Environment variables template
├── .env                   # Your environment variables (create this)
├── README.md              # This file
└── excel_agent.log        # Log file (created when running)
```

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve this agent.

## 📄 License

This project is open source and available under the MIT License.

## 🆘 Support

If you encounter any issues:
1. Check the troubleshooting section above
2. Review the log file `excel_agent.log`
3. Ensure all prerequisites are met
4. Verify your Excel file structure matches the expected format

---

**Happy Excel Automation! 🎉**