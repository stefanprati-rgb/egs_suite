import sys
import os
import importlib.util
from pathlib import Path
import traceback

# Setup path
ROOT_DIR = Path(__file__).parent
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

# Fake globals for some tkinter imports that might run code on import
os.environ['TK_SILENCE_DEPRECATION'] = '1'

def check_syntax(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            source = f.read()
        compile(source, file_path, 'exec')
        return True, None
    except Exception as e:
        return False, str(e)

def check_import(module_name):
    try:
        importlib.import_module(module_name)
        return True, None
    except Exception as e:
        return False, traceback.format_exc()

def main():
    print("=== Verification Started ===")
    
    # 1. Syntax Check
    print("\n--- Syntax Check ---")
    files = list(ROOT_DIR.rglob("*.py"))
    errors = 0
    for f in files:
        if "venv" in str(f) or "__pycache__" in str(f):
            continue
        
        rel_path = f.relative_to(ROOT_DIR)
        ok, err = check_syntax(f)
        if not ok:
            print(f"❌ Syntax Error in {rel_path}: {err}")
            errors += 1
        else:
            # print(f"✅ {rel_path}") # Noise reduction
            pass
            
    if errors == 0:
        print("✅ All files passed syntax check.")
    else:
        print(f"❌ Found {errors} syntax errors.")

    # 2. Import Check (Modules)
    print("\n--- Import Check ---")
    modules_to_check = [
        "EGS_Suite.common.logging",
        "EGS_Suite.common.config",
        # We can't easily check apps that launch UI on import without guarding
    ]
    
    # Add root to sys.path to simulate package structure
    sys.path.insert(0, str(ROOT_DIR.parent))
    
    for mod in modules_to_check:
        ok, err = check_import(mod)
        if ok:
            print(f"✅ Imported {mod}")
        else:
            print(f"❌ Failed to import {mod}\n{err}")

    print("\n=== Verification Finished ===")

if __name__ == "__main__":
    main()
