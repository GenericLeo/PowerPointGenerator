"""
Release preparation script
Run this before creating a new release to ensure version consistency
"""

import sys
from version import __version__, __app_name__, __github_repo__

def check_release_readiness():
    """Check if the application is ready for release"""
    
    print(f"\n{'='*60}")
    print(f"  {__app_name__} - Release Readiness Check")
    print(f"{'='*60}\n")
    
    issues = []
    warnings = []
    
    # Check version format
    try:
        parts = __version__.split('.')
        if len(parts) != 3:
            issues.append(f"Version format should be X.Y.Z, got: {__version__}")
        else:
            major, minor, patch = parts
            if not (major.isdigit() and minor.isdigit() and patch.isdigit()):
                issues.append(f"Version should only contain numbers: {__version__}")
    except Exception as e:
        issues.append(f"Invalid version format: {e}")
    
    # Check GitHub repo
    if __github_repo__ == "username/PowerPointGenerator":
        issues.append("GitHub repository not configured in version.py")
        print("❌ GitHub repo: NOT CONFIGURED")
        print("   Update __github_repo__ in version.py")
    elif '/' not in __github_repo__ or __github_repo__.count('/') != 1:
        issues.append(f"Invalid GitHub repo format: {__github_repo__}")
        print(f"❌ GitHub repo: INVALID FORMAT")
    else:
        print(f"✅ GitHub repo: {__github_repo__}")
    
    # Check if running in venv
    if sys.prefix == sys.base_prefix:
        warnings.append("Not running in virtual environment")
        print("⚠️  Virtual environment: NOT ACTIVATED")
    else:
        print(f"✅ Virtual environment: ACTIVE")
    
    # Display version info
    print(f"✅ Version: {__version__}")
    print(f"✅ App name: {__app_name__}")
    
    # Check for required files
    import os
    required_files = [
        'gui_app.py',
        'update_manager.py', 
        'version.py',
        'build_app.spec',
        'build_macos.sh',
        'requirements.txt'
    ]
    
    print("\nRequired files:")
    for file in required_files:
        if os.path.exists(file):
            print(f"  ✅ {file}")
        else:
            issues.append(f"Missing required file: {file}")
            print(f"  ❌ {file}")
    
    # Check for dependencies
    print("\nChecking dependencies...")
    try:
        import PIL
        print("  ✅ Pillow")
    except ImportError:
        issues.append("Pillow not installed")
        print("  ❌ Pillow")
    
    try:
        import pptx
        print("  ✅ python-pptx")
    except ImportError:
        issues.append("python-pptx not installed")
        print("  ❌ python-pptx")
    
    try:
        import packaging
        print("  ✅ packaging")
    except ImportError:
        issues.append("packaging not installed")
        print("  ❌ packaging")
    
    # Summary
    print(f"\n{'='*60}")
    if issues:
        print("❌ RELEASE NOT READY")
        print(f"\n{len(issues)} issue(s) found:")
        for issue in issues:
            print(f"  • {issue}")
    elif warnings:
        print("⚠️  RELEASE READY WITH WARNINGS")
        print(f"\n{len(warnings)} warning(s):")
        for warning in warnings:
            print(f"  • {warning}")
    else:
        print("✅ RELEASE READY!")
        print(f"\nNext steps:")
        print(f"  1. Run: ./build_macos.sh")
        print(f"  2. Test: open dist/PowerPointGenerator.app")
        print(f"  3. Create GitHub release with tag: v{__version__}")
        print(f"  4. Upload the DMG file from dist/")
    
    print(f"{'='*60}\n")
    
    return len(issues) == 0


if __name__ == "__main__":
    ready = check_release_readiness()
    sys.exit(0 if ready else 1)
