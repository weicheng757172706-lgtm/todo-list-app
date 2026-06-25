#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script for TodoList application
Tests if the GUI can be created without errors
"""

import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from todo_list_v4 import TodoApp

# Create root window
root = tk.Tk()
root.withdraw()  # Hide the window

try:
    # Try to create the app instance
    app = TodoApp(root)
    print("[OK] TodoApp instance created successfully!")
    
    # Test data loading
    print(f"[OK] Data loaded: {len(app.data['tasks'])} tasks, {len(app.data['completed'])} completed")
    
    # Test GUI elements
    print(f"[OK] Pending tree created: {app.pending_tree is not None}")
    print(f"[OK] Completed tree created: {app.completed_tree is not None}")
    print(f"[OK] Priority options: {app.priority_options}")
    
    print("\n[SUCCESS] All tests passed! The application should work correctly.")
    
except Exception as e:
    print(f"\n[ERROR] Error: {e}")
    import traceback
    traceback.print_exc()

finally:
    root.destroy()
