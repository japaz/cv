#!/bin/bash 
set -euo pipefail

# Default values
FORCE=false
DEBUG=false

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
        -h|--help)
            echo "Usage: $0 [-f|--force] [-d|--debug] [-h|--help]"
            echo "  -f, --force    Force reprocessing of all files"
            echo "  -d, --debug    Enable debug mode with verbose output"
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

# Enable debug mode if requested
if [ "$DEBUG" = true ]; then
    set -x
    echo "DEBUG: Debug mode enabled"
    echo "DEBUG: Current working directory: $(pwd)"
    echo "DEBUG: Current user: $(whoami)"
    echo "DEBUG: Available commands:"
    command -v pandoc || echo "DEBUG: pandoc not found"
    echo "DEBUG: Files in current directory:"
    ls -la *.md 2>/dev/null || echo "DEBUG: No .md files found"
    echo "DEBUG: Checking for custom-reference.docx:"
    ls -la custom-reference.docx 2>/dev/null || echo "DEBUG: custom-reference.docx not found"
fi

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
        
        if pandoc "$file" -o "$output_file" --reference-doc=custom-reference.docx; then
            echo "Generated: $output_file ($reason)"
            ((processed_count++))
            
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: Successfully generated $output_file"
                echo "DEBUG: File size: $(ls -lh "$output_file" 2>/dev/null | awk '{print $5}' || echo 'unknown')"
            fi
        else
            echo "Error: Failed to convert $file to $output_file"
            if [ "$DEBUG" = true ]; then
                echo "DEBUG: Pandoc exit code: $?"
                echo "DEBUG: Output directory contents:"
                ls -la output/ 2>/dev/null || echo "DEBUG: Output directory doesn't exist"
            fi
            exit 1
        fi
    else
        echo "Skipped: $file (output is up to date)"
        ((skipped_count++))
        
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
    fi
done

echo "Conversion complete. Processed: $processed_count, Skipped: $skipped_count"
echo "All docx files are in the output/ directory."
