#!/usr/bin/env python3
"""
Test script to verify the macOS packaging setup is working correctly
"""

import sys
import os

print("=" * 70)
print("PowerPoint Generator - macOS Packaging Verification")
print("=" * 70)
print()

all_passed = True

# Test 1: Import all modules
print("Test 1: Importing modules...")
try:
    from version import __version__, __app_name__, __github_repo__
    print(f"  ✅ version.py imported successfully")
    print(f"     App: {__app_name__} v{__version__}")
    print(f"     Repo: {__github_repo__}")
except ImportError as e:
    print(f"  ❌ Failed to import version.py: {e}")
    all_passed = False

try:
    from update_manager import UpdateManager
    print(f"  ✅ update_manager.py imported successfully")
except ImportError as e:
    print(f"  ❌ Failed to import update_manager.py: {e}")
    all_passed = False

try:
    from gui_app import ImageUploaderGUI
    print(f"  ✅ gui_app.py imported successfully")
except ImportError as e:
    print(f"  ❌ Failed to import gui_app.py: {e}")
    all_passed = False

print()

# Test 2: Check required files exist
print("Test 2: Checking required files...")
required_files = [
    'version.py',
    'update_manager.py',
    'gui_app.py',
    'build_app.spec',
    'build_macos.sh',
    'requirements.txt',
    'MACOS_DISTRIBUTION.md',
    'QUICKSTART_MACOS.md',
    'check_release.py',
]

for file in required_files:
    if os.path.exists(file):
        print(f"  ✅ {file}")
    else:
        print(f"  ❌ {file} not found")
        all_passed = False

print()

# Test 3: Check dependencies
print("Test 3: Checking dependencies...")
dependencies = ['PIL', 'pptx', 'packaging']

for dep in dependencies:
    try:
        __import__(dep)
        print(f"  ✅ {dep}")
    except ImportError:
        print(f"  ❌ {dep} not installed (run: pip install -r requirements.txt)")
        all_passed = False

print()

# Test 4: Test UpdateManager basic functionality
print("Test 4: Testing UpdateManager...")
try:
    manager = UpdateManager()
    current_version = manager.get_current_version()
    print(f"  ✅ UpdateManager initialized")
    print(f"     Current version: {current_version}")
    print(f"     GitHub repo: {manager.github_repo}")
    
    if manager.github_repo == "username/PowerPointGenerator":
        print(f"  ⚠️  Warning: GitHub repo not configured")
        print(f"     Update __github_repo__ in version.py before building")
    
except Exception as e:
    print(f"  ❌ UpdateManager test failed: {e}")
    all_passed = False

print()

# Test 5: Check build script permissions
print("Test 5: Checking build script...")
if os.path.exists('build_macos.sh'):
    if os.access('build_macos.sh', os.X_OK):
        print(f"  ✅ build_macos.sh is executable")
    else:
        print(f"  ⚠️  build_macos.sh is not executable")
        print(f"     Run: chmod +x build_macos.sh")
else:
    print(f"  ❌ build_macos.sh not found")
    all_passed = False

print()

# Summary
print("=" * 70)
if all_passed:
    print("✅ ALL TESTS PASSED!")
    print()
    print("Your setup is ready for macOS distribution!")
    print()
    print("Next steps:")
    print("  1. Update __github_repo__ in version.py")
    print("  2. Run: ./build_macos.sh")
    print("  3. Test: open dist/PowerPointGenerator.app")
    print("  4. Create a GitHub release with the DMG file")
else:
    print("❌ SOME TESTS FAILED")
    print()
    print("Please address the issues above before building.")
print("=" * 70)
print()

sys.exit(0 if all_passed else 1)
