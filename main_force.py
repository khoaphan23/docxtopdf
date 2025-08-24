#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FORCE RELOAD Word to PDF Converter
"""

import sys
import os
import importlib
from pathlib import Path

# Fix encoding for Windows console
if sys.platform == "win32":
    os.system("chcp 65001 >nul")
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass

print("=== FORCE CLEARING ALL PROJECT MODULES ===")

# Get all modules to clear BEFORE any imports
all_modules = list(sys.modules.keys())
project_modules = [m for m in all_modules if any(x in m for x in [
    'converters', 'interface', 'io.file_handler', 'logging.logger_setup',
    'src.converters', 'src.interface', 'src.io', 'src.logging'
])]

# Remove all project modules
for module in project_modules:
    if module in sys.modules:
        del sys.modules[module]
        print(f"CLEARED: {module}")

# Also clear __pycache__ related modules
pycache_modules = [m for m in all_modules if 'docxtopdf' in m or 'src' in m]
for module in pycache_modules:
    if module in sys.modules and module not in ['src']:  # Don't remove 'src' itself
        del sys.modules[module] 
        print(f"CLEARED CACHE: {module}")

sys.path.insert(0, str(Path(__file__).parent))

print("=== IMPORTING FRESH MODULES ===")

# Force import fresh modules
import src.converters.doc_converter
import src.interface.gui
import src.logging.logger_setup

# Reload them to be extra sure
importlib.reload(src.converters.doc_converter)
importlib.reload(src.interface.gui)
importlib.reload(src.logging.logger_setup)

from src.interface.gui import DocToPdfGUI
from src.logging.logger_setup import setup_logger
# Import config để sử dụng
from src import APP_NAME, VERSION, get_app_info

def main():
    try:
        # Sử dụng config cho logger setup
        logger = setup_logger()  # Sử dụng mặc định từ config
        
        # Log thông tin ứng dụng từ config
        app_info = get_app_info()
        logger.info(f"Starting {app_info['name']} v{app_info['version']} by {app_info['author']}")
        print(f"Starting {APP_NAME} v{VERSION}...")
        
        app = DocToPdfGUI()
        print("✓ GUI created successfully")
        
        # Test converter methods
        methods = app.converter.check_conversion_methods()
        print(f"✓ Available methods: {methods}")
        logger.info(f"Converter methods: {methods}")
        
        # Print converter source to verify it's the right one
        import inspect
        converter_file = inspect.getfile(app.converter.__class__)
        print(f"✓ Using converter from: {converter_file}")
        
        if not methods['supported']:
            print("❌ WARNING: No conversion methods available!")
            logger.error("No conversion methods available!")
        
        print("🚀 Starting GUI with FRESH converter...")
        app.run()
        
        logger.info("Application closed")
        
    except Exception as e:
        print(f"Error starting application: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()