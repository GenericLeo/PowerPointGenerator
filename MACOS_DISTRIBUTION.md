# macOS Distribution & Auto-Update Guide

## Overview

This guide explains how to build, distribute, and automatically update your PowerPoint Generator application for macOS users.

## Table of Contents

1. [Quick Start](#quick-start)
2. [Building the Application](#building-the-application)
3. [Creating GitHub Releases](#creating-github-releases)
4. [Auto-Update System](#auto-update-system)
5. [Distribution Workflow](#distribution-workflow)
6. [Troubleshooting](#troubleshooting)

---

## Quick Start

### First Time Setup

1. **Update version.py with your GitHub repository:**
   ```python
   __github_repo__ = "yourusername/PowerPointGenerator"  # Update this!
   ```

2. **Install dependencies:**
   ```bash
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. **Build the application:**
   ```bash
   ./build_macos.sh
   ```

4. **Test the application:**
   ```bash
   open dist/PowerPointGenerator.app
   ```

---

## Building the Application

### Using the Build Script (Recommended)

The `build_macos.sh` script automates the entire build process:

```bash
./build_macos.sh
```

This script will:
- ✅ Verify dependencies
- ✅ Clean previous builds
- ✅ Build the .app bundle with PyInstaller
- ✅ Optionally create a distributable DMG file
- ✅ Display next steps

### Manual Build

If you prefer to build manually:

```bash
# Activate virtual environment
source .venv/bin/activate

# Install PyInstaller
pip install pyinstaller

# Clean previous builds
rm -rf dist build

# Build the application
pyinstaller build_app.spec
```

---

## Creating GitHub Releases

### Step 1: Prepare for Release

1. **Update the version in [version.py](version.py):**
   ```python
   __version__ = "1.1.0"  # Increment version number
   
   VERSION_HISTORY = {
       "1.1.0": "Added new feature X and fixed bug Y",
       "1.0.0": "Initial release"
   }
   ```

2. **Build the DMG:**
   ```bash
   ./build_macos.sh
   # Answer "y" when prompted to create DMG
   ```

3. **Test the DMG:**
   - Mount the DMG file
   - Drag the app to Applications
   - Test that it works correctly

### Step 2: Create the GitHub Release

1. **Commit and push your changes:**
   ```bash
   git add .
   git commit -m "Release v1.1.0"
   git push
   ```

2. **Create a new release on GitHub:**
   - Go to your repository on GitHub
   - Click "Releases" → "Create a new release"
   - Tag version: `v1.1.0` (must match version.py with 'v' prefix)
   - Release title: `PowerPoint Generator v1.1.0`
   - Description: Copy the release notes from VERSION_HISTORY

3. **Upload the DMG:**
   - Drag and drop `dist/PowerPointGenerator-1.1.0-macOS.dmg`
   - Important: The filename should contain "macOS" or "mac" or end with ".dmg"

4. **Publish the release**

### Step 3: Verify Auto-Update

1. Open an older version of your app
2. Go to Help → Check for Updates
3. Verify that the update notification appears correctly

---

## Auto-Update System

### How It Works

1. **Update Check:**
   - The app checks GitHub's API for the latest release
   - Uses the repository specified in `version.py`
   - Compares semantic versions (1.0.0 < 1.1.0 < 2.0.0)

2. **User Notification:**
   - If update available: Shows dialog with release notes
   - User can download directly or skip
   - Update check happens on demand (Help → Check for Updates)

3. **Installation:**
   - Downloads open in the user's browser
   - User manually installs the new version
   - Standard macOS .app replacement workflow

### Update Manager API

The [update_manager.py](update_manager.py) provides:

```python
from update_manager import UpdateManager

manager = UpdateManager()

# Check for updates
result = manager.check_for_updates()

if result['update_available']:
    print(f"New version: {result['latest_version']}")
    print(f"Download: {result['download_url']}")
    print(f"Notes: {result['release_notes']}")

# Open download page
manager.download_and_install_update(result['download_url'])
```

### Customizing Update Behavior

You can trigger automatic checks at startup:

```python
# In gui_app.py __init__ method:
self.root.after(5000, lambda: self.check_for_updates(silent=True))
```

This checks for updates 5 seconds after launch (silent = no popup if no update).

---

## Distribution Workflow

### For Team Distribution

1. **Create a shared location:**
   - Upload DMG to GitHub Releases (recommended)
   - Or use company file server/Dropbox

2. **Share with coworkers:**
   - Send them the DMG download link
   - Example: `https://github.com/yourusername/PowerPointGenerator/releases/latest`

3. **Installation instructions for users:**
   ```
   1. Download PowerPointGenerator-X.X.X-macOS.dmg
   2. Open the DMG file
   3. Drag PowerPoint Generator.app to Applications folder
   4. Launch from Applications
   ```

### For End Users

1. **First-time setup:**
   - Download and install from GitHub Releases
   - macOS may show security warning (right-click → Open to bypass)

2. **Getting updates:**
   - Updates are announced through Help → Check for Updates
   - Or check GitHub Releases page manually

### Gatekeeper & Code Signing (Optional)

For wider distribution, consider signing your app:

```bash
# Sign the app (requires Apple Developer account)
codesign --deep --force --sign "Developer ID Application: Your Name" \
         dist/PowerPointGenerator.app

# Verify signature
codesign --verify --verbose dist/PowerPointGenerator.app
```

To enable code signing:
1. Get an Apple Developer account ($99/year)
2. Create a Developer ID Application certificate
3. Update `build_app.spec`: `codesign_identity="Developer ID Application: Your Name"`

---

## Troubleshooting

### Build Issues

**Problem:** PyInstaller not found
```bash
pip install pyinstaller
```

**Problem:** Module not found during build
- Add to `hiddenimports` in `build_app.spec`
- Example: `hiddenimports=['missing_module']`

**Problem:** App crashes on launch
- Run from terminal to see error: `./dist/PowerPointGenerator.app/Contents/MacOS/PowerPointGenerator`
- Check Console.app for crash logs

### Update Check Issues

**Problem:** "Repository not found" error
- Verify `__github_repo__` in `version.py` is correct
- Format: `"username/repository-name"`
- Make sure repository is public (or use GitHub token for private repos)

**Problem:** No updates detected
- Verify release tag matches format: `v1.0.0` (must have 'v' prefix)
- Verify DMG filename contains "mac" or "macos" or ends with ".dmg"
- Check that release is published (not draft)

**Problem:** Can't download update
- Verify DMG was uploaded to GitHub release
- Check that asset is publicly accessible

### Distribution Issues

**Problem:** macOS blocks app ("cannot be opened because it is from an unidentified developer")
**Solution:** 
- Right-click app → Open (first time only)
- Or: System Preferences → Security & Privacy → "Open Anyway"

**Problem:** App permission errors
- The app may need permissions for certain folders
- macOS will prompt users automatically

---

## Version Numbering

Follow Semantic Versioning (semver):

- **Major** (1.0.0 → 2.0.0): Breaking changes
- **Minor** (1.0.0 → 1.1.0): New features, backward compatible
- **Patch** (1.0.0 → 1.0.1): Bug fixes

Update `__version__` in `version.py` before each release.

---

## GitHub Repository Setup

### Repository Settings

1. **Make repository public** (for free hosting)
   - Or add GitHub token support for private repos

2. **Enable Releases:**
   - Releases are enabled by default
   - Users can subscribe to release notifications

3. **Set up Release Notes:**
   - Use the `VERSION_HISTORY` in `version.py` as a template
   - Include clear descriptions of changes

---

## Support & Maintenance

### Regular Updates

1. Fix bugs or add features
2. Update version in `version.py`
3. Build and test locally
4. Create GitHub release with new DMG
5. Users get update notification automatically

### Monitoring

- Check GitHub release download counts
- Monitor issues/feedback from users
- Keep dependencies updated: `pip list --outdated`

---

## Additional Resources

- [PyInstaller Documentation](https://pyinstaller.org/)
- [GitHub Releases Guide](https://docs.github.com/en/repositories/releasing-projects-on-github)
- [Apple Code Signing Guide](https://developer.apple.com/documentation/security/notarizing_macos_software_before_distribution)
- [Semantic Versioning](https://semver.org/)

---

## Quick Reference

### Build Commands
```bash
./build_macos.sh           # Build with DMG creation option
pyinstaller build_app.spec # Build without DMG
```

### Test Commands
```bash
open dist/PowerPointGenerator.app                          # Launch app
./dist/PowerPointGenerator.app/Contents/MacOS/PowerPointGenerator  # Debug mode
python update_manager.py                                    # Test update check
```

### Release Checklist
- [ ] Update `__version__` in version.py
- [ ] Update `VERSION_HISTORY` in version.py
- [ ] Update `__github_repo__` (first time only)
- [ ] Run `./build_macos.sh`
- [ ] Test the built application
- [ ] Create DMG
- [ ] Commit and push changes
- [ ] Create GitHub release with tag `vX.Y.Z`
- [ ] Upload DMG to release
- [ ] Publish release
- [ ] Test update notification in old version

---

**Need help?** Create an issue on GitHub or check the troubleshooting section above.
