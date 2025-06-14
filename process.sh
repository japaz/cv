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
mkdir -p output

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
        # Run pandoc command with specified settings
        pandoc "$file" -o "$output_file" --reference-doc=custom-reference.docx
        echo "Generated: $output_file ($reason)"
        ((processed_count++))
    else
        echo "Skipped: $file (output is up to date)"
        ((skipped_count++))
    fi
done

echo "Conversion complete. Processed: $processed_count, Skipped: $skipped_count"
echo "All docx files are in the output/ directory."
