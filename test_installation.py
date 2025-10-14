#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script để kiểm tra cài đặt Auto Tay Luyen Tool
"""

import sys
import importlib

def test_import(module_name, package_name=None):
    """Test import một module"""
    try:
        importlib.import_module(module_name)
        print(f"[OK] {package_name or module_name}: OK")
        return True
    except ImportError as e:
        print(f"[FAILED] {package_name or module_name}: FAILED - {e}")
        return False

def main():
    print("=" * 50)
    print("Auto Tay Luyen Tool - Installation Test")
    print("=" * 50)
    print()
    
    # Test Python version
    print(f"Python version: {sys.version}")
    print()
    
    # Test required modules
    modules = [
        ("tkinter", "Tkinter (GUI)"),
        ("PIL", "Pillow (Image processing)"),
        ("pyautogui", "PyAutoGUI (Mouse/Keyboard control)"),
        ("pytesseract", "Pytesseract (OCR)"),
        ("pygetwindow", "PyGetWindow (Window management)"),
        ("keyboard", "Keyboard (Hotkey support)"),
    ]
    
    all_ok = True
    for module, name in modules:
        if not test_import(module, name):
            all_ok = False
    
    print()
    
    # Test Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        print("[OK] Tesseract-OCR: OK")
    except pytesseract.TesseractNotFoundError:
        print("[FAILED] Tesseract-OCR: NOT FOUND")
        print("   Please install Tesseract-OCR and add it to PATH")
        all_ok = False
    except Exception as e:
        print(f"[FAILED] Tesseract-OCR: ERROR - {e}")
        all_ok = False
    
    print()
    print("=" * 50)
    
    if all_ok:
        print("[SUCCESS] All tests passed! You can run: python auto_tay_luyen.py")
    else:
        print("[WARNING] Some tests failed. Please check the installation.")
        print("   Run install.bat to install missing packages.")
    
    print("=" * 50)
    # input("Press Enter to exit...")

if __name__ == "__main__":
    main()
