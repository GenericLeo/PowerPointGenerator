"""
Version information for PowerPoint Generator
Update this file when releasing new versions
"""

__version__ = "1.0.9"
__app_name__ = "PowerPoint Generator"
__bundle_id__ = "com.powerpoint.generator"
__author__ = "Leo Soler"
__github_repo__ = "GenericLeo/PowerPointGenerator"

# Version history
VERSION_HISTORY = {
    "1.0.9": "Hotfix: Restored Shift+Up/Down keyboard range selection in file list.",
    "1.0.8": "Enhanced Ctrl+Click multi-select: additive first click, removal requires second click on same file. Improved modifier detection for macOS.",
    "1.0.7": "Added Shift+Up/Down keyboard range selection in the file list viewer.",
    "1.0.6": "Applied NCSU visual theme updates across the GUI and added implementation summary documentation.",
    "1.0.5": "Fixed update checker to download the correct platform asset (Windows exe vs macOS dmg).",
    "1.0.4": "Added automatic GitHub Actions CI/CD for multi-platform builds.",
    "1.0.3": "App automatically checks for updates on startup.",
    "1.0.2": "Fixed app launch from DMG by storing data in a writable user directory.",
    "1.0.1": "Fixed updater SSL reliability and robust GitHub release version parsing.",
    "1.0.0": "Initial release with macOS packaging and auto-update support"
}
