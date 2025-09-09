#!/usr/bin/env python3
"""
Excel Trial Balance Autonomous Agent

This script provides an autonomous agent that can update Excel working papers
leadsheet accounts using fuzzy matching and AI-powered analysis.

Features:
- Automatic detection of open Excel files
- Interactive setup for sheet and column mapping
- Fuzzy matching with 80% threshold
- Ignores empty and bold cells
- Updates current year and prior year amounts
- Adds new accounts from reference sheet
- Powered by OpenRouter API with Qwen 3 Coder model

Usage:
    python main.py

Requirements:
    - Excel file must be open before running
    - OpenRouter API key must be set in .env file
    - All required packages must be installed (see requirements.txt)
"""

import sys
import os
from pathlib import Path
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# Add current directory to path
sys.path.append(str(Path(__file__).parent))

try:
    from excel_agent import ExcelTrialBalanceAgent
except ImportError as e:
    print(f"{Fore.RED}Error importing ExcelTrialBalanceAgent: {e}{Style.RESET_ALL}")
    print(f"{Fore.YELLOW}Please ensure all dependencies are installed: pip install -r requirements.txt{Style.RESET_ALL}")
    sys.exit(1)


def check_prerequisites():
    """
    Check if all prerequisites are met before running the agent.
    
    Returns:
        bool: True if all prerequisites are met, False otherwise.
    """
    print(f"{Fore.CYAN}🔍 Checking prerequisites...{Style.RESET_ALL}")
    
    # Check if .env file exists
    env_file = Path(".env")
    if not env_file.exists():
        print(f"{Fore.YELLOW}⚠️  .env file not found. Please create one based on .env.example{Style.RESET_ALL}")
        print(f"{Fore.CYAN}   Copy .env.example to .env and add your OpenRouter API key{Style.RESET_ALL}")
        return False
    
    # Check if API key is set
    from dotenv import load_dotenv
    load_dotenv()
    
    api_key = os.getenv('OPENROUTER_API_KEY')
    if not api_key or api_key == 'your_openrouter_api_key_here':
        print(f"{Fore.RED}❌ OpenRouter API key not set in .env file{Style.RESET_ALL}")
        print(f"{Fore.CYAN}   Please add your API key to the .env file{Style.RESET_ALL}")
        print(f"{Fore.CYAN}   Get your API key from: https://openrouter.ai/keys{Style.RESET_ALL}")
        return False
    
    print(f"{Fore.GREEN}✅ All prerequisites met{Style.RESET_ALL}")
    return True


def print_banner():
    """
    Print the application banner.
    """
    banner = f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║           📊 Excel Trial Balance Autonomous Agent 📊          ║
║                                                              ║
║  🤖 Powered by OpenRouter API & Qwen 3 Coder Model          ║
║  📈 Intelligent Fuzzy Matching for Account Updates          ║
║  🎯 Automatic Detection of Excel Files                      ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝{Style.RESET_ALL}
"""
    print(banner)


def print_instructions():
    """
    Print usage instructions.
    """
    instructions = f"""
{Fore.YELLOW}📋 Before running this agent, please ensure:

1. 📂 Open your Excel file with the trial balance sheets
2. 🔑 Set your OpenRouter API key in the .env file
3. 📊 Have both sheets ready:
   - Sheet with data to be updated (incorrect amounts)
   - Sheet with correct reference amounts

🎯 The agent will:
   ✅ Ask you to identify the sheets and columns
   ✅ Show previews of identified data
   ✅ Perform fuzzy matching (80% threshold)
   ✅ Update amounts while ignoring empty and bold cells
   ✅ Add new accounts from the reference sheet

Press Enter to continue or Ctrl+C to exit...{Style.RESET_ALL}
"""
    print(instructions)
    
    try:
        input()
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}Exiting...{Style.RESET_ALL}")
        sys.exit(0)


def main():
    """
    Main execution function.
    """
    try:
        # Print banner
        print_banner()
        
        # Check prerequisites
        if not check_prerequisites():
            print(f"\n{Fore.RED}❌ Prerequisites not met. Please fix the issues above and try again.{Style.RESET_ALL}")
            return 1
        
        # Print instructions
        print_instructions()
        
        # Initialize and run the agent
        print(f"{Fore.CYAN}🚀 Initializing Excel Trial Balance Agent...{Style.RESET_ALL}")
        
        agent = ExcelTrialBalanceAgent()
        agent.run()
        
        print(f"\n{Fore.GREEN}🎉 Agent execution completed successfully!{Style.RESET_ALL}")
        return 0
        
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}⏹️  Process interrupted by user. Exiting gracefully...{Style.RESET_ALL}")
        return 0
        
    except Exception as e:
        print(f"\n{Fore.RED}💥 Unexpected error occurred: {str(e)}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}💡 Please check the log file 'excel_agent.log' for detailed error information.{Style.RESET_ALL}")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)