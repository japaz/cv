# PDF Conversion Scripts for Microsoft Word on macOS

This directory contains scripts to convert DOCX files to PDF using Microsoft Word on macOS. The main script `process.sh` has been enhanced to optionally generate PDF files alongside DOCX files.

## Files

- `docx_to_pdf.applescript` - AppleScript for converting DOCX to PDF
- `convert_to_pdf.sh` - Standalone shell script wrapper for converting DOCX to PDF
- `batch_convert_pdf.sh` - Batch processing script with fewer permission prompts
- `batch_docx_to_pdf.applescript` - Batch AppleScript for processing multiple files
- `setup_permissions.sh` - Helper script for setting up permissions
- `process.sh` - Main processing script with integrated PDF generation

## Prerequisites

- macOS with Microsoft Word installed (for PDF generation)
- pandoc (for DOCX generation)
- The scripts will automatically detect if Microsoft Word is available and skip PDF generation if not installed

## Usage

### Method 1: Direct AppleScript Usage

```bash
# Using AppleScript
osascript docx_to_pdf.applescript "/path/to/input.docx" "/path/to/output.pdf"
```

### Method 2: Shell Script Wrapper

```bash
# Convert single file
./convert_to_pdf.sh document.docx
./convert_to_pdf.sh document.docx output.pdf

# Convert all DOCX files in a directory
./convert_to_pdf.sh output/

# Force overwrite existing PDFs
./convert_to_pdf.sh -f document.docx

# Verbose output
./convert_to_pdf.sh -v document.docx
```

### Method 3: Main Process Script (Recommended)

```bash
# Generate only DOCX files (original behavior)
./process.sh

# Generate both DOCX and PDF files
./process.sh -p

# Force regenerate everything including PDFs
./process.sh -f -p

# Debug mode with PDF generation
./process.sh -d -p
```

## Options for convert_to_pdf.sh

- `-f, --force` - Overwrite existing PDF files
- `-v, --verbose` - Enable verbose output
- `-h, --help` - Show help message

## Options for process.sh

- `-f, --force` - Force reprocessing of all files
- `-d, --debug` - Enable debug mode with verbose output  
- `-p, --pdf` - Generate PDF files using Microsoft Word
- `-h, --help` - Show help message

## How It Works

1. **AppleScript**: These scripts use macOS automation to control Microsoft Word directly
   - Open the DOCX file in Word
   - Save it as PDF format
   - Close the document

2. **Shell Wrapper**: Provides a user-friendly interface with batch processing capabilities

3. **Enhanced Process Script**: Integrates PDF generation into your existing workflow

## Troubleshooting

### Microsoft Word Permission Issues

If you're getting permission prompts for each file, here are several solutions:

#### Method 1: Grant Full Disk Access (Recommended)

1. Open System Preferences → Security & Privacy → Privacy → Full Disk Access
2. Click the lock and enter your password
3. Click "+" and add Microsoft Word
4. Ensure Microsoft Word is checked/enabled

#### Method 2: Use Batch Processing

```bash
# Process all files in one batch (fewer permission prompts)
./batch_convert_pdf.sh -b

# Or process specific directories
./batch_convert_pdf.sh -b output/cv output/cover-letters
```

#### Method 3: Set up Approved Directory

```bash
# Run the permission setup script
./setup_permissions.sh
```

#### Method 4: Interactive Mode

```bash
# Interactive mode allows you to control each conversion
./process.sh -p --interactive
```

### Microsoft Word Not Found

- Ensure Microsoft Word is installed from the Mac App Store or Office 365
- The script checks for Word using the bundle identifier `com.microsoft.Word`

### Permission Issues

- macOS may ask for permission to control Microsoft Word the first time you run the script
- Grant the necessary accessibility permissions in System Preferences > Security & Privacy > Privacy > Accessibility

### Script Execution Issues

- Make sure the scripts are executable: `chmod +x script_name.sh`
- Run from the correct directory where the scripts are located

## Performance Notes

- Microsoft Word needs to launch for each conversion, so batch processing multiple files will be more efficient than converting one by one
- The scripts will reuse an already-running instance of Word if available

## Examples

```bash
# Quick single file conversion
./convert_to_pdf.sh output/cv/Alberto_Paz_Jimenez.docx

# Batch convert all DOCX files in output directory to PDF
./convert_to_pdf.sh output/

# Generate everything (DOCX + PDF) with your existing workflow
./process.sh -p

# Force regenerate everything when you've updated templates
./process.sh -f -p
```
