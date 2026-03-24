"""
Version information for PowerPoint Generator
Update this file when releasing new versions
"""

__version__ = "1.0.4"
__app_name__ = "PowerPoint Generator"
__bundle_id__ = "com.powerpoint.generator"
__author__ = "Leo Soler"
__github_repo__ = "GenericLeo/PowerPointGenerator"

# Version history
VERSION_HISTORY = {
    "1.0.4": "Fixed update checker to download the correct platform asset (Windows exe vs macOS dmg).",
    "1.0.3": "App automatically checks for updates on startup.",
    "1.0.2": "Fixed app launch from DMG by storing data in a writable user directory.",
    "1.0.1": "Fixed updater SSL reliability and robust GitHub release version parsing.",
    "1.0.0": "Initial release with macOS packaging and auto-update support"
}
