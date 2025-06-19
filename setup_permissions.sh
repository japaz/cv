#!/bin/bash
set -euo pipefail

# Script to grant Microsoft Word full access to the output directory
# This helps avoid individual file permission requests

echo "Setting up Microsoft Word permissions for batch PDF conversion..."

# Get the current directory
CURRENT_DIR="$(pwd)"
OUTPUT_DIR="$CURRENT_DIR/output"

echo "Current working directory: $CURRENT_DIR"
echo "Output directory: $OUTPUT_DIR"

# Method 1: Try to add the output directory to Word's accessible locations
echo ""
echo "Method 1: Adding output directory to Microsoft Word's Full Disk Access..."
echo "Please follow these steps manually:"
echo ""
echo "1. Open System Preferences (or System Settings on newer macOS)"
echo "2. Go to Security & Privacy (or Privacy & Security)"
echo "3. Click on 'Full Disk Access' in the sidebar"
echo "4. Click the lock icon and enter your password"
echo "5. Click the '+' button"
echo "6. Navigate to Applications and select 'Microsoft Word'"
echo "7. Make sure Microsoft Word is checked/enabled"
echo ""
echo "This will give Word access to all files and eliminate permission prompts."
echo ""

# Method 2: Alternative using Terminal access
echo "Method 2: Alternative - Grant Terminal Full Disk Access..."
echo ""
echo "If Method 1 doesn't work, you can:"
echo "1. Follow the same steps above, but add 'Terminal' instead of Microsoft Word"
echo "2. This allows our scripts (running through Terminal) to access all files"
echo ""

# Method 3: Create a user-approved directory
echo "Method 3: Create approved directory structure..."

# Check if we can create a symlink approach
if [ ! -d "$HOME/Documents/CV_Processing" ]; then
    mkdir -p "$HOME/Documents/CV_Processing"
    echo "Created ~/Documents/CV_Processing directory"
fi

# Create symlinks to make files accessible from Documents folder
if [ ! -L "$HOME/Documents/CV_Processing/output" ]; then
    ln -sf "$OUTPUT_DIR" "$HOME/Documents/CV_Processing/output"
    echo "Created symlink from ~/Documents/CV_Processing/output to $OUTPUT_DIR"
fi

echo ""
echo "Method 3 complete: Files are now accessible via ~/Documents/CV_Processing/output"
echo "You can try running the conversion from there if other methods don't work."
echo ""

# Method 4: Interactive permission granting
echo "Method 4: Interactive permission granting..."
echo ""
echo "Run this command to grant permissions interactively:"
echo "./process.sh -p --interactive"
echo ""

echo "Choose the method that works best for your system:"
echo "- Method 1 (Full Disk Access) is the most comprehensive"
echo "- Method 2 works if you trust Terminal with full access"  
echo "- Method 3 uses the Documents folder which typically has fewer restrictions"
echo "- Method 4 allows you to grant permissions one by one with better control"
