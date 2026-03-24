# Quick Start Guide - macOS Distribution

## 🚀 One-Time Setup

1. **Edit [version.py](version.py) - Set your GitHub repo:**
   ```python
   __github_repo__ = "YOUR_USERNAME/PowerPointGenerator"
   ```

2. **Install dependencies:**
   ```bash
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

## 📦 Building Your App

```bash
./build_macos.sh
```

When prompted, say **"y"** to create a DMG file for distribution.

## 🌟 Creating a Release

1. **Update version in version.py:**
   ```python
   __version__ = "1.1.0"
   ```

2. **Build the app** (see above)

3. **Upload to GitHub:**
   - Go to your repo → Releases → "New release"
   - Tag: `v1.1.0` (must have 'v' prefix)
   - Upload the DMG file from `dist/`
   - Publish!

## ✨ That's It!

Your coworkers can now:
- Download the DMG from GitHub Releases
- Drag it to Applications
- Get automatic update notifications

## 📖 Need More Details?

See [MACOS_DISTRIBUTION.md](MACOS_DISTRIBUTION.md) for complete documentation.

## 💡 Update Your App Version

Every time you make changes:
1. Change version number in `version.py`
2. Run `./build_macos.sh`
3. Create new GitHub release
4. Users automatically see "Update Available" dialog
