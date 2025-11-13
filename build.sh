#!/bin/bash

# PowerPoint Generator - Build Script
# This script builds a standalone macOS application

echo "======================================"
echo "PowerPoint Generator - Build Script"
echo "======================================"
echo ""

# Check if virtual environment is activated
if [ -z "$VIRTUAL_ENV" ]; then
    echo "Activating virtual environment..."
    source .venv/bin/activate
fi

# Install PyInstaller if not already installed
echo "Checking for PyInstaller..."
pip show pyinstaller > /dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Installing PyInstaller..."
    pip install pyinstaller
else
    echo "PyInstaller already installed"
fi

echo ""
echo "Building application..."
echo ""

# Clean previous builds
if [ -d "dist" ]; then
    echo "Cleaning previous builds..."
    rm -rf dist
fi

if [ -d "build" ]; then
    rm -rf build
fi

# Build the application
pyinstaller build_app.spec

# Check if build was successful
if [ -d "dist/PowerPointGenerator.app" ]; then
    echo ""
    echo "======================================"
    echo "✓ Build successful!"
    echo "======================================"
    echo ""
    echo "Your application is located at:"
    echo "  dist/PowerPointGenerator.app"
    echo ""
    echo "You can:"
    echo "  1. Double-click to run it from Finder"
    echo "  2. Move it to /Applications folder"
    echo "  3. Run from terminal: open dist/PowerPointGenerator.app"
    echo ""
else
    echo ""
    echo "======================================"
    echo "✗ Build failed"
    echo "======================================"
    echo ""
    echo "Please check the error messages above."
    exit 1
fi
