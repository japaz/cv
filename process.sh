#!/bin/bash 
set -euo pipefail

# Default values
FORCE=false
DEBUG=false
GENERATE_PDF=false
INTERACTIVE_MODE=false

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -f|--force)
            FORCE=true
            shift
            ;;
        -d|--debug)
            DEBUG=true
            shift
            ;;
        -p|--pdf)
            GENERATE_PDF=true
            shift
            ;;
        --interactive)
            INTERACTIVE_MODE=true
            shift
            ;;
        -h|--help)
            echo "Usage: $0 [-f|--force] [-d|--debug] [-p|--pdf] [--interactive] [-h|--help]"
            echo "  -f, --force       Force reprocessing of all files"
            echo "  -d, --debug       Enable debug mode with verbose output"
            echo "  -p, --pdf         Generate PDF files using Microsoft Word"
            echo "  --interactive     Interactive mode with permission prompts"
            echo "  -h, --help        Show this help message"
            echo ""
            echo "Examples:"
            echo "  $0                    # Generate only DOCX files"
            echo "  $0 -p                 # Generate DOCX and PDF files"
            echo "  $0 -f -p              # Force regenerate everything including PDFs"
            echo "  $0 -p --interactive   # Interactive mode with permission control"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            echo "Use -h or --help for usage information"
            exit 1
            ;;
    esac
done

# Enable debug mode if requested
if [ "$DEBUG" = true ]; then
    set -x
    echo "DEBUG: Debug mode enabled"
    echo "DEBUG: Current working directory: $(pwd)"
    echo "DEBUG: Current user: $(whoami)"
    echo "DEBUG: PDF generation: $GENERATE_PDF"
    echo "DEBUG: PDF method: $PDF_METHOD"
    echo "DEBUG: Available commands:"
    command -v pandoc || echo "DEBUG: pandoc not found"
    echo "DEBUG: Files in current directory:"
    ls -la *.md 2>/dev/null || echo "DEBUG: No .md files found"
    echo "DEBUG: Checking for custom-reference.docx:"
    ls -la custom-reference.docx 2>/dev/null || echo "DEBUG: custom-reference.docx not found"
fi

# Check if Microsoft Word is available for PDF generation
if [ "$GENERATE_PDF" = true ]; then
    if ! osascript -e 'tell application "Finder" to return exists (application file id "com.microsoft.Word")' >/dev/null 2>&1; then
        echo "Warning: Microsoft Word is not installed. PDF generation will be skipped."
        GENERATE_PDF=false
    else
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Microsoft Word found, PDF generation enabled"
        fi
    fi
fi

# Create output directory if it doesn't exist
if [ "$DEBUG" = true ]; then
    echo "DEBUG: Creating output directory..."
fi

if ! mkdir -p output; then
    echo "Error: Failed to create output directory"
    if [ "$DEBUG" = true ]; then
        echo "DEBUG: mkdir -p output failed with exit code: $?"
        echo "DEBUG: Current directory permissions:"
        ls -la .
    fi
    exit 1
fi

# Verify output directory exists and is writable
if [ ! -d output ]; then
    echo "Error: Output directory was not created"
    exit 1
fi

if [ ! -w output ]; then
    echo "Error: Output directory is not writable"
    if [ "$DEBUG" = true ]; then
        echo "DEBUG: Output directory permissions:"
        ls -la output
        echo "DEBUG: Current user and groups:"
        id
    fi
    exit 1
fi

if [ "$DEBUG" = true ]; then
    echo "DEBUG: Output directory created successfully"
    echo "DEBUG: Output directory info:"
    ls -la output
fi

# Counters for reporting
processed_count=0
skipped_count=0
pdf_converted=0
pdf_skipped=0

# Function to convert DOCX to PDF using Microsoft Word
convert_docx_to_pdf() {
    local docx_file="$1"
    local pdf_file="$2"
    
    if [ "$DEBUG" = true ]; then
        echo "DEBUG: Converting DOCX to PDF: $docx_file -> $pdf_file"
    fi
    
    # Check if PDF already exists and is newer than DOCX (unless force is enabled)
    if [ "$FORCE" = false ] && [ -f "$pdf_file" ] && [ "$pdf_file" -nt "$docx_file" ]; then
        echo "Skipped PDF: $pdf_file (already up to date)"
        return 0
    fi
    
    # Interactive mode - ask user for permission
    if [ "$INTERACTIVE_MODE" = true ]; then
        echo "About to convert: $docx_file -> $pdf_file"
        echo "This will open Microsoft Word and may require file access permission."
        read -p "Continue? [y/N]: " -n 1 -r
        echo
        if [[ ! $REPLY =~ ^[Yy]$ ]]; then
            echo "Skipped: $pdf_file (user choice)"
            return 0
        fi
    fi
    
    # Create output directory for PDF if it doesn't exist
    local pdf_dir
    pdf_dir="$(dirname "$pdf_file")"
    mkdir -p "$pdf_dir"
    
    # Get absolute paths (macOS compatible)
    local abs_docx_file
    local abs_pdf_file
    
    # Convert to absolute path for DOCX (must exist)
    abs_docx_file="$(cd "$(dirname "$docx_file")" && pwd)/$(basename "$docx_file")"
    
    # Convert to absolute path for PDF (may not exist yet)
    abs_pdf_file="$(cd "$pdf_dir" && pwd)/$(basename "$pdf_file")"
    
    if [ "$DEBUG" = true ]; then
        echo "DEBUG: Absolute DOCX path: $abs_docx_file"
        echo "DEBUG: Absolute PDF path: $abs_pdf_file"
    fi
    
    # Convert using AppleScript
    local result
    result=$(osascript "$SCRIPT_DIR/docx_to_pdf.applescript" "$abs_docx_file" "$abs_pdf_file" 2>&1)
    
    # Check if conversion was successful
    if [[ "$result" == *"SUCCESS"* && -f "$abs_pdf_file" ]]; then
        echo "Generated PDF: $pdf_file"
        return 0
    else
        echo "Error: Failed to convert $docx_file to PDF"
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Conversion result: $result"
        fi
        return 1
    fi
}

# Process all Markdown files in cv and cover-letters directories
for dir in cv cover-letters; do
    if [ ! -d "$dir" ]; then
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Directory $dir does not exist, skipping"
        fi
        continue
    fi
    
    if [ "$DEBUG" = true ]; then
        echo "DEBUG: Processing directory: $dir"
    fi
    
    for file in "$dir"/*.md; do
        # Skip if no .md files exist (glob doesn't match)
        [ ! -f "$file" ] && continue        
        # Get base filename without extension and directory
        base_name=$(basename "$file" .md)
        dir_name=$(dirname "$file")
        output_file="output/${dir_name}/${base_name}.docx"
        
        # Create output subdirectory if it doesn't exist
        output_dir="output/${dir_name}"
        if ! mkdir -p "$output_dir"; then
            echo "Error: Failed to create output directory: $output_dir"
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: mkdir -p $output_dir failed with exit code: $?"
            fi
            exit 1
        fi
    
    # Check if we need to process this file
    should_process=false
    
    if [ "$FORCE" = true ]; then
        should_process=true
        reason="forced reprocessing"
    elif [ ! -f "$output_file" ]; then
        should_process=true
        reason="output file doesn't exist"
    elif [ "$file" -nt "$output_file" ]; then
        should_process=true
        reason="source file is newer than output"
    fi
    
    if [ "$should_process" = true ]; then
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Processing file: $file"
            echo "DEBUG: Output file: $output_file"
            echo "DEBUG: Reason: $reason"
        fi
        
        # Check if pandoc is available
        if ! command -v pandoc &> /dev/null; then
            echo "Error: pandoc is not installed or not in PATH"
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: PATH=$PATH"
                echo "DEBUG: Available commands in PATH:"
                echo "$PATH" | tr ':' '\n' | while read dir; do
                    if [ -d "$dir" ]; then
                        ls "$dir" 2>/dev/null | grep -E '^pandoc' || true
                    fi
                done
            fi
            exit 1
        fi
        
        # Check if reference document exists
        if [ ! -f "custom-reference.docx" ]; then
            echo "Error: custom-reference.docx not found"
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: Current directory contents:"
                ls -la
            fi
            exit 1
        fi
        
        # Run pandoc command with specified settings
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Running pandoc command:"
            echo "DEBUG: pandoc \"$file\" -o \"$output_file\" --reference-doc=custom-reference.docx"
        fi
        
        # Capture pandoc output and exit code
        pandoc_output=""
        pandoc_exit_code=0
        
        if [ "$DEBUG" = true ]; then
            # Run pandoc with verbose output in debug mode
            pandoc_output=$(pandoc "$file" -o "$output_file" --reference-doc=custom-reference.docx --verbose 2>&1) || pandoc_exit_code=$?
        else
            # Run pandoc normally
            pandoc_output=$(pandoc "$file" -o "$output_file" --reference-doc=custom-reference.docx 2>&1) || pandoc_exit_code=$?
        fi
        
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Pandoc exit code: $pandoc_exit_code"
            if [ -n "$pandoc_output" ]; then
                echo "DEBUG: Pandoc output:"
                echo "$pandoc_output"
            fi
        fi
        
        if [ $pandoc_exit_code -eq 0 ]; then
            # Verify the output file was actually created
            if [ ! -f "$output_file" ]; then
                echo "Error: Output file was not created: $output_file"
                if [ "$DEBUG" = true ]; then
                    echo "DEBUG: Pandoc claimed success but output file doesn't exist"
                    echo "DEBUG: Output directory contents:"
                    ls -la output/ 2>/dev/null || echo "DEBUG: Output directory doesn't exist"
                fi
                exit 1
            fi
            
            echo "Generated: $output_file ($reason)"
            
            # Use a more reliable way to increment the counter
            processed_count=$((processed_count + 1))
            
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: Successfully generated $output_file"
                echo "DEBUG: Processed count: $processed_count"
                echo "DEBUG: File size: $(ls -lh "$output_file" 2>/dev/null | awk '{print $5}' || echo 'unknown')"
                echo "DEBUG: File permissions: $(ls -la "$output_file" 2>/dev/null || echo 'unknown')"
            fi
            
            # Generate PDF if requested and DOCX exists
            if [ "$GENERATE_PDF" = true ] && [ -f "$output_file" ]; then
                pdf_output_file="${output_file%.docx}.pdf"
                if convert_docx_to_pdf "$output_file" "$pdf_output_file"; then
                    pdf_converted=$((pdf_converted + 1))
                else
                    pdf_skipped=$((pdf_skipped + 1))
                fi
            fi
        else
            echo "Error: Failed to convert $file to $output_file (exit code: $pandoc_exit_code)"
            if [ "$DEBUG" = true ]; then
                if [ -n "$pandoc_output" ]; then
                    echo "DEBUG: Pandoc error output:"
                    echo "$pandoc_output"
                fi
                echo "DEBUG: Output directory contents:"
                ls -la output/ 2>/dev/null || echo "DEBUG: Output directory doesn't exist"
            fi
            exit 1
        fi
    else
        echo "Skipped: $file (output is up to date)"
        
        # Use a more reliable way to increment the counter
        skipped_count=$((skipped_count + 1))
        
        if [ "$DEBUG" = true ]; then
            echo "DEBUG: Skipped $file - output file: $output_file"
            if [ -f "$output_file" ]; then
                # Use compatible stat commands for both macOS and Linux
                if stat -c %Y "$output_file" >/dev/null 2>&1; then
                    # Linux
                    echo "DEBUG: Output file timestamp: $(stat -c %y "$output_file")"
                    echo "DEBUG: Source file timestamp: $(stat -c %y "$file")"
                else
                    # macOS
                    echo "DEBUG: Output file timestamp: $(stat -f %Sm "$output_file")"
                    echo "DEBUG: Source file timestamp: $(stat -f %Sm "$file")"
                fi
            fi
        fi
        
        # Generate PDF if requested and DOCX exists
        if [ "$GENERATE_PDF" = true ] && [ -f "$output_file" ]; then
            pdf_output_file="${output_file%.docx}.pdf"
            if convert_docx_to_pdf "$output_file" "$pdf_output_file"; then
                pdf_converted=$((pdf_converted + 1))
            else
                pdf_skipped=$((pdf_skipped + 1))
            fi
        fi
        fi
    done
done

echo "Conversion complete. Processed: $processed_count, Skipped: $skipped_count"
if [ "$GENERATE_PDF" = true ]; then
    echo "PDF generation complete. Converted: $pdf_converted, Skipped/Failed: $pdf_skipped"
fi
echo "All files are in the output/ directory."
