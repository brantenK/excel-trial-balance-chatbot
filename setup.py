#!/usr/bin/env python3
"""
Setup script for Excel Trial Balance Autonomous Agent

This script helps users set up the environment and install all necessary dependencies.
"""

import subprocess
import sys
import os
from pathlib import Path
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

def print_banner():
    """Print setup banner."""
    banner = f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║              🛠️  Excel Agent Setup Wizard 🛠️                ║
║                                                              ║
║         Setting up your autonomous Excel agent...           ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝{Style.RESET_ALL}
"""
    print(banner)

def check_python_version():
    """Check if Python version is compatible."""
    print(f"{Fore.YELLOW}🐍 Checking Python version...{Style.RESET_ALL}")
    
    if sys.version_info < (3, 8):
        print(f"{Fore.RED}❌ Python 3.8+ is required. Current version: {sys.version}{Style.RESET_ALL}")
        return False
    
    print(f"{Fore.GREEN}✅ Python {sys.version.split()[0]} is compatible{Style.RESET_ALL}")
    return True

def install_dependencies():
    """Install required Python packages."""
    print(f"\n{Fore.YELLOW}📦 Installing dependencies...{Style.RESET_ALL}")
    
    try:
        # Upgrade pip first
        print(f"{Fore.CYAN}Upgrading pip...{Style.RESET_ALL}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        
        # Install requirements
        print(f"{Fore.CYAN}Installing packages from requirements.txt...{Style.RESET_ALL}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        
        print(f"{Fore.GREEN}✅ All dependencies installed successfully{Style.RESET_ALL}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"{Fore.RED}❌ Error installing dependencies: {e}{Style.RESET_ALL}")
        return False
    except FileNotFoundError:
        print(f"{Fore.RED}❌ requirements.txt file not found{Style.RESET_ALL}")
        return False

def setup_xlwings():
    """Setup xlwings Excel add-in."""
    print(f"\n{Fore.YELLOW}🔧 Setting up xlwings Excel add-in...{Style.RESET_ALL}")
    
    try:
        # Install xlwings add-in
        subprocess.check_call([sys.executable, "-m", "xlwings", "addin", "install"])
        print(f"{Fore.GREEN}✅ xlwings add-in installed successfully{Style.RESET_ALL}")
        print(f"{Fore.CYAN}💡 Please restart Excel to activate the add-in{Style.RESET_ALL}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"{Fore.YELLOW}⚠️  xlwings add-in installation failed: {e}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}💡 You may need to install it manually or run Excel as Administrator{Style.RESET_ALL}")
        return False

def setup_environment():
    """Setup environment variables."""
    print(f"\n{Fore.YELLOW}🔐 Setting up environment configuration...{Style.RESET_ALL}")
    
    env_file = Path(".env")
    env_example = Path(".env.example")
    
    if env_file.exists():
        print(f"{Fore.GREEN}✅ .env file already exists{Style.RESET_ALL}")
        return True
    
    if not env_example.exists():
        print(f"{Fore.RED}❌ .env.example file not found{Style.RESET_ALL}")
        return False
    
    # Copy .env.example to .env
    try:
        with open(env_example, 'r') as src, open(env_file, 'w') as dst:
            dst.write(src.read())
        
        print(f"{Fore.GREEN}✅ Created .env file from template{Style.RESET_ALL}")
        print(f"{Fore.CYAN}💡 Please edit .env file and add your OpenRouter API key{Style.RESET_ALL}")
        return True
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error creating .env file: {e}{Style.RESET_ALL}")
        return False

def get_api_key():
    """Prompt user for OpenRouter API key."""
    print(f"\n{Fore.YELLOW}🔑 OpenRouter API Key Setup{Style.RESET_ALL}")
    print(f"{Fore.CYAN}To use this agent, you need an OpenRouter API key.{Style.RESET_ALL}")
    print(f"{Fore.CYAN}Get one for free at: https://openrouter.ai/keys{Style.RESET_ALL}")
    
    while True:
        choice = input(f"\n{Fore.YELLOW}Do you want to enter your API key now? (y/n): {Style.RESET_ALL}").lower().strip()
        
        if choice == 'y':
            api_key = input(f"{Fore.CYAN}Enter your OpenRouter API key: {Style.RESET_ALL}").strip()
            
            if api_key and api_key.startswith('sk-or-'):
                # Update .env file
                try:
                    env_file = Path(".env")
                    if env_file.exists():
                        content = env_file.read_text()
                        content = content.replace('your_openrouter_api_key_here', api_key)
                        env_file.write_text(content)
                        
                        print(f"{Fore.GREEN}✅ API key saved to .env file{Style.RESET_ALL}")
                        return True
                    else:
                        print(f"{Fore.RED}❌ .env file not found{Style.RESET_ALL}")
                        return False
                        
                except Exception as e:
                    print(f"{Fore.RED}❌ Error saving API key: {e}{Style.RESET_ALL}")
                    return False
            else:
                print(f"{Fore.RED}❌ Invalid API key format. Should start with 'sk-or-'{Style.RESET_ALL}")
                continue
                
        elif choice == 'n':
            print(f"{Fore.YELLOW}⚠️  You can add your API key later by editing the .env file{Style.RESET_ALL}")
            return True
        else:
            print(f"{Fore.RED}Please enter 'y' or 'n'{Style.RESET_ALL}")

def run_test():
    """Run a basic test to verify installation."""
    print(f"\n{Fore.YELLOW}🧪 Running basic tests...{Style.RESET_ALL}")
    
    try:
        # Test imports
        import xlwings
        import pandas
        import fuzzywuzzy
        from dotenv import load_dotenv
        
        print(f"{Fore.GREEN}✅ All required packages can be imported{Style.RESET_ALL}")
        
        # Test xlwings
        try:
            apps = xlwings.apps
            print(f"{Fore.GREEN}✅ xlwings is working correctly{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️  xlwings test failed: {e}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}💡 This is normal if Excel is not currently running{Style.RESET_ALL}")
        
        return True
        
    except ImportError as e:
        print(f"{Fore.RED}❌ Import test failed: {e}{Style.RESET_ALL}")
        return False

def print_next_steps():
    """Print next steps for the user."""
    next_steps = f"""
{Fore.GREEN}🎉 Setup completed successfully!

{Fore.CYAN}📋 Next Steps:

1. 📂 Open your Excel file with trial balance sheets
2. 🚀 Run the agent: {Fore.WHITE}python main.py{Fore.CYAN}
3. 📊 Follow the interactive setup process
4. ✅ Let the agent update your trial balance automatically

{Fore.YELLOW}💡 Tips:
- Make sure Excel is open before running the agent
- Have both your 'to-update' and 'reference' sheets ready
- The agent will guide you through column mapping
- Check the log file 'excel_agent.log' for detailed information

{Fore.GREEN}Happy Excel automation! 🎯{Style.RESET_ALL}
"""
    print(next_steps)

def main():
    """Main setup function."""
    print_banner()
    
    success = True
    
    # Check Python version
    if not check_python_version():
        return 1
    
    # Install dependencies
    if not install_dependencies():
        success = False
    
    # Setup xlwings
    if not setup_xlwings():
        success = False
    
    # Setup environment
    if not setup_environment():
        success = False
    
    # Get API key
    if not get_api_key():
        success = False
    
    # Run tests
    if not run_test():
        success = False
    
    if success:
        print_next_steps()
        return 0
    else:
        print(f"\n{Fore.RED}❌ Setup completed with some issues. Please check the messages above.{Style.RESET_ALL}")
        return 1

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}Setup interrupted by user. Exiting...{Style.RESET_ALL}")
        sys.exit(0)
    except Exception as e:
        print(f"\n{Fore.RED}Unexpected error during setup: {e}{Style.RESET_ALL}")
        sys.exit(1)