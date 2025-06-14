#!/bin/bash 
set -euo pipefail

# Default values
FORCE=false

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -f|--force)
            FORCE=true
            shift
            ;;
        -h|--help)
            echo "Usage: $0 [-f|--force] [-h|--help]"
            echo "  -f, --force    Force reprocessing of all files"
            echo "  -h, --help     Show this help message"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            echo "Use -h or --help for usage information"
            exit 1
            ;;
    esac
done

# Create output directory if it doesn't exist
if ! mkdir -p output; then
    echo "Error: Failed to create output directory"
    exit 1
fi

# Verify output directory is writable
if [ ! -w output ]; then
    echo "Error: Output directory is not writable"
    exit 1
fi

# Counters for reporting
processed_count=0
skipped_count=0

# Process all Markdown files in current directory
for file in *.md; do
    # Skip if no .md files exist (glob doesn't match)
    [ ! -f "$file" ] && continue
    
    # Get base filename without extension
    base_name=$(basename "$file" .md)
    output_file="output/${base_name}.docx"
    
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
        # Check if pandoc is available
        if ! command -v pandoc &> /dev/null; then
            echo "Error: pandoc is not installed or not in PATH"
            exit 1
        fi
        
        # Check if reference document exists
        if [ ! -f "custom-reference.docx" ]; then
            echo "Error: custom-reference.docx not found"
            exit 1
        fi
        
        # Run pandoc command with specified settings
        if pandoc "$file" -o "$output_file" --reference-doc=custom-reference.docx; then
            echo "Generated: $output_file ($reason)"
            ((processed_count++))
        else
            echo "Error: Failed to convert $file to $output_file"
            exit 1
        fi
    else
        echo "Skipped: $file (output is up to date)"
        ((skipped_count++))
    fi
done

echo "Conversion complete. Processed: $processed_count, Skipped: $skipped_count"
echo "All docx files are in the output/ directory."
