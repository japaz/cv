#!/bin/bash
set -euo pipefail

# Batch PDF conversion script that minimizes permission prompts
# This processes all files in a directory at once

BATCH_MODE=false
DIRECTORIES=()

# Parse arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -b|--batch)
            BATCH_MODE=true
            shift
            ;;
        -h|--help)
            echo "Usage: $0 [-b|--batch] [directory1] [directory2] ..."
            echo ""
            echo "  -b, --batch    Use batch processing mode (fewer permission prompts)"
            echo "  -h, --help     Show this help message"
            echo ""
            echo "Examples:"
            echo "  $0 output/cv                    # Convert all DOCX files in output/cv"
            echo "  $0 -b output/cv output/cover-letters  # Batch convert both directories"
            echo "  $0 output/                      # Convert all DOCX files in all subdirectories"
            exit 0
            ;;
        *)
            DIRECTORIES+=("$1")
            shift
            ;;
    esac
done

# Default to output directory if none specified
if [ ${#DIRECTORIES[@]} -eq 0 ]; then
    DIRECTORIES=("output/cv" "output/cover-letters")
fi

echo "Batch PDF Conversion Tool"
echo "========================="
echo ""

# Check if Microsoft Word is available
if ! osascript -e 'tell application "Finder" to return exists (application file id "com.microsoft.Word")' >/dev/null 2>&1; then
    echo "Error: Microsoft Word is not installed or not found"
    exit 1
fi

if [ "$BATCH_MODE" = true ]; then
    echo "Using batch processing mode (recommended for fewer permission prompts)"
    echo ""
    
    for dir in "${DIRECTORIES[@]}"; do
        if [ ! -d "$dir" ]; then
            echo "Warning: Directory '$dir' does not exist, skipping"
            continue
        fi
        
        echo "Processing directory: $dir"
        
        # Count DOCX files
        docx_count=$(find "$dir" -name "*.docx" -type f | wc -l | tr -d ' ')
        
        if [ "$docx_count" -eq 0 ]; then
            echo "  No DOCX files found in $dir"
            continue
        fi
        
        echo "  Found $docx_count DOCX file(s)"
        
        # Use batch AppleScript
        result=$(osascript batch_docx_to_pdf.applescript "$(realpath "$dir")" 2>&1)
        
        if [[ "$result" == *"SUCCESS"* ]]; then
            echo "  ✓ $result"
        else
            echo "  ✗ $result"
        fi
        
        echo ""
    done
else
    echo "Using individual file processing mode"
    echo ""
    
    total_converted=0
    total_failed=0
    
    for dir in "${DIRECTORIES[@]}"; do
        if [ ! -d "$dir" ]; then
            echo "Warning: Directory '$dir' does not exist, skipping"
            continue
        fi
        
        echo "Processing directory: $dir"
        
        # Find all DOCX files and convert them
        while IFS= read -r -d '' docx_file; do
            # Generate PDF filename
            pdf_file="${docx_file%.docx}.pdf"
            
            echo "  Converting: $(basename "$docx_file")"
            
            # Get absolute paths (required for AppleScript)
            abs_docx_file="$(cd "$(dirname "$docx_file")" && pwd)/$(basename "$docx_file")"
            abs_pdf_file="$(cd "$(dirname "$pdf_file")" && pwd)/$(basename "$pdf_file")"
            
            # Convert using single-file AppleScript
            if result=$(osascript docx_to_pdf.applescript "$abs_docx_file" "$abs_pdf_file" 2>&1); then
                if [[ "$result" == *"SUCCESS"* ]]; then
                    echo "    ✓ Success"
                    ((total_converted++))
                else
                    echo "    ✗ Failed: $result"
                    ((total_failed++))
                fi
            else
                echo "    ✗ Failed: $result"
                ((total_failed++))
            fi
        done < <(find "$dir" -name "*.docx" -type f -print0)
        
        echo ""
    done
    
    echo "Summary: $total_converted converted, $total_failed failed"
fi

echo "Batch conversion complete!"
