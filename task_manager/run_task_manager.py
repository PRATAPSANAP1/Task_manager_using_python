#!/usr/bin/env python3
"""
Task Manager Application Runner
Run this script to start the Task Manager GUI application.
"""

import sys
import subprocess
import importlib.util

def check_and_install_requirements():
    """Check if required packages are installed, install if missing."""
    required_packages = ['openpyxl', 'tkcalendar']
    missing_packages = []
    
    for package in required_packages:
        if importlib.util.find_spec(package) is None:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"Installing missing packages: {', '.join(missing_packages)}")
        for package in missing_packages:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print("All required packages installed successfully!")

if __name__ == "__main__":
    try:
        # Check and install requirements
        check_and_install_requirements()
        
        # Import and run the Task Manager
        from Task import TaskManagerApp
        import tkinter as tk
        
        print("Starting Task Manager Application...")
        root = tk.Tk()
        app = TaskManagerApp(root)
        root.mainloop()
        
    except ImportError as e:
        print(f"Import Error: {e}")
        print("Please ensure all required packages are installed.")
        print("Run: pip install -r requirements.txt")
    except Exception as e:
        print(f"Error running Task Manager: {e}")
        input("Press Enter to exit...")