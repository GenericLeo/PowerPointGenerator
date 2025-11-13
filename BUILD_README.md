# PowerPoint Generator - Standalone Application

## ✓ Build Complete!

Your standalone application has been successfully created!

### Location
```
dist/PowerPointGenerator.app
```

### How to Use

#### Option 1: Double-click to Run
1. Open Finder
2. Navigate to the `dist` folder
3. Double-click `PowerPointGenerator.app`

#### Option 2: Move to Applications
1. Drag `PowerPointGenerator.app` from the `dist` folder to your `/Applications` folder
2. Launch from Launchpad or Applications folder

#### Option 3: Run from Terminal
```bash
open dist/PowerPointGenerator.app
```

### First Launch on macOS

When you first open the app, macOS may show a security warning because the app is not from the App Store. To open it:

1. **If you see "cannot be opened":**
   - Right-click (or Control-click) on `PowerPointGenerator.app`
   - Select "Open" from the menu
   - Click "Open" in the dialog that appears

2. **Alternative method:**
   - Go to System Settings → Privacy & Security
   - Scroll down to find the blocked app
   - Click "Open Anyway"

### Features

The standalone application includes all features:
- ✓ Image upload and indexing (TIFF and other formats)
- ✓ Intelligent grouping by filename identifiers
- ✓ Visual layout designer
- ✓ PowerPoint generation with multiple layouts
- ✓ Search and filter capabilities
- ✓ Image preview
- ✓ Group editing

### Data Storage

The application creates an `image_index.json` file in the same directory where it's launched to store your indexed images.

### Rebuilding the Application

If you make changes to the source code and want to rebuild:

```bash
cd "/Users/leosoler/Layer Measurements/PowerPointGenerator"
source .venv/bin/activate
./build.sh
```

### Technical Details

- **Built with:** PyInstaller 6.16.0
- **Python Version:** 3.13.3
- **Architecture:** ARM64 (Apple Silicon)
- **Bundle Type:** macOS Application Bundle (.app)
- **Size:** ~50-60 MB (includes Python runtime and all dependencies)

### Dependencies Included

- Python 3.13 runtime
- Pillow (image processing)
- python-pptx (PowerPoint generation)
- tkinter (GUI framework)
- All other required libraries

### Troubleshooting

**App won't open:**
- Make sure you've allowed it in System Settings → Privacy & Security
- Try running from Terminal to see any error messages

**"Damaged" app warning:**
- Remove the quarantine attribute: `xattr -d com.apple.quarantine dist/PowerPointGenerator.app`

**Missing image_index.json:**
- The app creates this file automatically on first use
- Make sure the app has write permissions in its directory

---

**Note:** This is a standalone application and does not require Python or any dependencies to be installed on the target machine.
