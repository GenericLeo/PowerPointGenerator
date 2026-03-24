#!/bin/bash

# PowerPoint Generator - Enhanced Build Script for macOS
# Creates a distributable .app bundle and optionally a .dmg file

set -e  # Exit on error

# Color codes for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${BLUE}======================================"
echo "PowerPoint Generator - Build Script"
echo "======================================${NC}"
echo ""

# Get version from version.py
VERSION=$(python3 -c "from version import __version__; print(__version__)")
echo -e "${BLUE}Building version: ${GREEN}${VERSION}${NC}"
echo ""

# Check if virtual environment is activated
if [ -z "$VIRTUAL_ENV" ]; then
    echo "Activating virtual environment..."
    if [ ! -d ".venv" ]; then
        echo -e "${RED}Error: Virtual environment not found. Please create it first:${NC}"
        echo "  python3 -m venv .venv"
        echo "  source .venv/bin/activate"
        echo "  pip install -r requirements.txt"
        exit 1
    fi
    source .venv/bin/activate
fi

# Install/upgrade required packages
echo "Ensuring dependencies are up to date..."
pip install -q --upgrade pip
pip install -q -r requirements.txt

# Install PyInstaller if not already installed
echo "Checking for PyInstaller..."
pip show pyinstaller > /dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Installing PyInstaller..."
    pip install pyinstaller
else
    echo "PyInstaller is installed"
fi

echo ""
echo -e "${BLUE}Building application...${NC}"
echo ""

# Clean previous builds
if [ -d "dist" ]; then
    echo "Cleaning previous builds..."
    rm -rf dist
fi

if [ -d "build" ]; then
    rm -rf build
fi

# Create dist directory
mkdir -p dist

# Build the application
echo "Running PyInstaller..."
pyinstaller --clean build_app.spec

# Check if build was successful
if [ -d "dist/PowerPointGenerator.app" ]; then
    echo ""
    echo -e "${GREEN}======================================"
    echo "✓ Build successful!"
    echo "======================================${NC}"
    echo ""
    echo "Application: dist/PowerPointGenerator.app"
    
    # Ask if user wants to create DMG
    read -p "Create a distributable DMG file? (y/n) " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo ""
        echo -e "${BLUE}Creating DMG file...${NC}"
        
        DMG_NAME="PowerPointGenerator-${VERSION}-macOS.dmg"
        
        # Remove existing DMG if present
        [ -f "dist/${DMG_NAME}" ] && rm "dist/${DMG_NAME}"
        
        # Create DMG using hdiutil
        hdiutil create -volname "PowerPoint Generator" \
                       -srcfolder "dist/PowerPointGenerator.app" \
                       -ov -format UDZO \
                       "dist/${DMG_NAME}"
        
        if [ -f "dist/${DMG_NAME}" ]; then
            echo ""
            echo -e "${GREEN}✓ DMG created successfully!${NC}"
            echo "DMG file: dist/${DMG_NAME}"
            echo ""
            echo -e "${BLUE}To distribute your app:${NC}"
            echo "1. Upload dist/${DMG_NAME} to GitHub Releases"
            echo "2. Tag the release with v${VERSION}"
            echo "3. Users will be notified of the update automatically"
        else
            echo -e "${RED}✗ DMG creation failed${NC}"
        fi
    fi
    
    echo ""
    echo -e "${BLUE}Next steps:${NC}"
    echo "• Test the app: open dist/PowerPointGenerator.app"
    echo "• Create a GitHub release with tag v${VERSION}"
    echo "• Upload the DMG file to the GitHub release"
    echo "• Users will receive automatic update notifications"
    echo ""
    
else
    echo ""
    echo -e "${RED}======================================"
    echo "✗ Build failed!"
    echo "======================================${NC}"
    echo ""
    echo "Check the output above for errors."
    exit 1
fi
