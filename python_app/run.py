#!/usr/bin/env python3
"""
Quick launcher script for Microsoft Graph API Explorer
"""

import sys
import subprocess
import os

def check_dependencies():
    """Check if required dependencies are installed"""
    try:
        import tkinter
        print("✅ tkinter found")
    except ImportError:
        print("❌ tkinter not found. Please install python3-tk")
        print("   Ubuntu/Debian: sudo apt-get install python3-tk")
        print("   CentOS/RHEL: sudo yum install tkinter")
        return False
    
    try:
        import pyperclip
        print("✅ pyperclip found")
    except ImportError:
        print("⚠️  pyperclip not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])
            print("✅ pyperclip installed successfully")
        except subprocess.CalledProcessError:
            print("❌ Failed to install pyperclip. Please run: pip install pyperclip")
            return False
    
    return True

def main():
    """Main launcher function"""
    print("🔷 Microsoft Graph API Explorer - Python")
    print("=" * 50)
    
    if not check_dependencies():
        print("\n❌ Dependencies check failed. Please resolve the issues above.")
        sys.exit(1)
    
    print("\n🚀 Launching application...")
    
    # Change to script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Import and run the main application
    try:
        from main import main as run_app
        run_app()
    except Exception as e:
        print(f"❌ Error launching application: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
