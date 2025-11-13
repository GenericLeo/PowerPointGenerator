# PowerPoint Generator - Windows Build Instructions

## Building on Windows

### Prerequisites

1. **Python 3.8 or later** installed from [python.org](https://www.python.org/downloads/)
   - During installation, check "Add Python to PATH"

2. **Git** (optional) to clone the repository

### Step-by-Step Instructions

#### 1. Transfer Project to Windows Machine

**Option A: Using Git**
```cmd
git clone <your-repository-url>
cd PowerPointGenerator
```

**Option B: Manual Transfer**
- Copy the entire project folder to your Windows machine
- Make sure to include all `.py` files, `requirements.txt`, and the build scripts

#### 2. Run the Build Script

Open Command Prompt in the project directory and run:

```cmd
build_windows.bat
```

This script will:
- Create a virtual environment (if needed)
- Install all dependencies
- Install PyInstaller
- Build the executable
- Output to `dist\PowerPointGenerator.exe`

#### 3. Manual Build (Alternative)

If the batch script doesn't work, you can build manually:

```cmd
# Create and activate virtual environment
python -m venv .venv
.venv\Scripts\activate.bat

# Install dependencies
pip install -r requirements.txt
pip install pyinstaller

# Build the application
pyinstaller build_windows.spec
```

### Output

After successful build, you'll find:
- **Executable:** `dist\PowerPointGenerator.exe`
- **Size:** ~40-50 MB
- **Dependencies:** All included (standalone)

### Distribution

To share the application with others:

1. **Copy the entire `dist` folder** to the target machine
2. The `.exe` file needs the supporting files in the `dist` folder
3. No Python installation required on target machines

### Alternative: Single-File Executable

To create a single `.exe` file (easier to distribute but slower startup):

Edit `build_windows.spec` and change:
```python
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,      # Include these
    a.zipfiles,      # Include these
    a.datas,         # Include these
    [],
    name='PowerPointGenerator',
    # ... rest of settings
    onefile=True,    # Add this line
)
```

Then run:
```cmd
pyinstaller build_windows.spec
```

## Option 2: Cross-Compilation (Advanced)

### Using Wine on macOS/Linux

You can attempt to build Windows executables from macOS using Wine, but this is **not recommended** as it's complex and unreliable:

```bash
# Install Wine (on macOS with Homebrew)
brew install wine-stable

# Install Python for Windows via Wine
# This is complex and often has compatibility issues
```

**This approach is NOT recommended** - building on native Windows is much more reliable.

## Option 3: Using CI/CD (GitHub Actions)

Create a GitHub Actions workflow to automatically build for both platforms:

Create `.github/workflows/build.yml`:

```yaml
name: Build Applications

on:
  push:
    branches: [ main ]
  workflow_dispatch:

jobs:
  build-macos:
    runs-on: macos-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          pip install pyinstaller
      - name: Build macOS app
        run: pyinstaller build_app.spec
      - name: Upload artifact
        uses: actions/upload-artifact@v3
        with:
          name: macos-app
          path: dist/PowerPointGenerator.app

  build-windows:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          pip install pyinstaller
      - name: Build Windows exe
        run: pyinstaller build_windows.spec
      - name: Upload artifact
        uses: actions/upload-artifact@v3
        with:
          name: windows-exe
          path: dist/PowerPointGenerator.exe
```

This will automatically build both versions when you push to GitHub.

## Recommended Approach

**For best results:**

1. ✅ **Build on native Windows** using `build_windows.bat`
2. ✅ **Build on macOS** using `build.sh` (already done)
3. ✅ Distribute both versions to users

**OR**

Use GitHub Actions to build both automatically.

## Testing

After building on Windows, test that:
- [ ] Application launches without errors
- [ ] GUI displays correctly
- [ ] Image upload works
- [ ] PowerPoint generation works
- [ ] All features function as expected

## Troubleshooting

**Missing DLLs:**
- Ensure all dependencies are in `requirements.txt`
- Some Windows systems may need Microsoft Visual C++ Redistributable

**Antivirus blocks exe:**
- This is common with PyInstaller executables
- Users may need to add an exception in their antivirus

**Large file size:**
- This is normal - includes Python runtime and all libraries
- Use `onefile=True` for single-file distribution

## Support

For Windows-specific issues:
- Check PyInstaller documentation: https://pyinstaller.org/
- Ensure Python version compatibility (3.8-3.12 recommended)
- Test on multiple Windows versions (10, 11)
