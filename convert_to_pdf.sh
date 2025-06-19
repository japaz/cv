#!/bin/bash
set -euo pipefail

# Shell script wrapper for converting DOCX to PDF using Microsoft Word
# This script can convert a single file or batch process multiple files

# Default values
VERBOSE=false
FORCE=false

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Usage function
usage() {
    echo "Usage: $0 [OPTIONS] INPUT [OUTPUT]"
    echo ""
    echo "Convert DOCX files to PDF using Microsoft Word"
    echo ""
    echo "Arguments:"
    echo "  INPUT     Input DOCX file or directory containing DOCX files"
    echo "  OUTPUT    Output PDF file or directory (optional for single files)"
    echo ""
    echo "Options:"
    echo "  -f, --force      Overwrite existing PDF files"
    echo "  -v, --verbose    Enable verbose output"
    echo "  -h, --help       Show this help message"
    echo ""
    echo "Examples:"
    echo "  $0 document.docx                    # Convert to document.pdf"
    echo "  $0 document.docx output.pdf         # Convert to specific output file"
    echo "  $0 output/                          # Convert all DOCX files in directory"
}

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -f|--force)
            FORCE=true
            shift
            ;;
        -v|--verbose)
            VERBOSE=true
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        -*)
            echo "Unknown option: $1"
            usage
            exit 1
            ;;
        *)
            break
            ;;
    esac
done

# Check if we have at least one argument
if [[ $# -eq 0 ]]; then
    echo "Error: No input file or directory specified"
    usage
    exit 1
fi

INPUT="$1"
OUTPUT="${2:-}"

# Check if Microsoft Word is installed
if ! osascript -e 'tell application "System Events" to return exists (application process "Microsoft Word")' >/dev/null 2>&1; then
    if ! osascript -e 'tell application "Finder" to return exists (application file id "com.microsoft.Word")' >/dev/null 2>&1; then
        echo "Error: Microsoft Word is not installed or not found"
        exit 1
    fi
fi

# Function to convert a single DOCX file to PDF
convert_file() {
    local input_file="$1"
    local output_file="$2"
    
    # Check if input file exists
    if [[ ! -f "$input_file" ]]; then
        echo "Error: Input file '$input_file' does not exist"
        return 1
    fi
    
    # Check if input file is a DOCX file
    if [[ "${input_file}" != *.docx && "${input_file}" != *.DOCX ]]; then
        echo "Error: Input file '$input_file' is not a DOCX file"
        return 1
    fi
    
    # Check if output file already exists and force is not enabled
    if [[ -f "$output_file" && "$FORCE" != true ]]; then
        echo "Skipped: $output_file already exists (use -f to overwrite)"
        return 0
    fi
    
    # Create output directory if it doesn't exist
    local output_dir
    output_dir="$(dirname "$output_file")"
    mkdir -p "$output_dir"
    
    # Convert absolute paths (macOS compatible)
    local abs_input
    local abs_output
    
    # Get absolute path for input file (must exist)
    abs_input="$(cd "$(dirname "$input_file")" && pwd)/$(basename "$input_file")"
    
    # Get absolute path for output file (may not exist yet)
    local output_dir
    output_dir="$(dirname "$output_file")"
    abs_output="$(cd "$output_dir" && pwd)/$(basename "$output_file")"
    
    if [[ "$VERBOSE" == true ]]; then
        echo "Converting: $abs_input -> $abs_output"
    fi
    
    # Run the AppleScript
    local result
    result=$(osascript "$SCRIPT_DIR/docx_to_pdf.applescript" "$abs_input" "$abs_output" 2>&1)
    
    # Check if conversion was successful
    if [[ "$result" == *"SUCCESS"* && -f "$abs_output" ]]; then
        echo "Generated: $output_file"
        return 0
    else
        echo "Error: Failed to convert $input_file"
        if [[ "$VERBOSE" == true ]]; then
            echo "Result: $result"
        fi
        return 1
    fi
}

# Main conversion logic
if [[ -f "$INPUT" ]]; then
    # Single file conversion
    if [[ -z "$OUTPUT" ]]; then
        # Generate output filename by replacing .docx with .pdf
        OUTPUT="${INPUT%.docx}.pdf"
    fi
    
    convert_file "$INPUT" "$OUTPUT"
    
elif [[ -d "$INPUT" ]]; then
    # Directory conversion
    if [[ -n "$OUTPUT" && ! -d "$OUTPUT" ]]; then
        echo "Error: When input is a directory, output must also be a directory or empty"
        exit 1
    fi
    
    # Set default output directory if not specified
    if [[ -z "$OUTPUT" ]]; then
        OUTPUT="$INPUT"
    fi
    
    # Find all DOCX files in the input directory
    local docx_files
    mapfile -t docx_files < <(find "$INPUT" -name "*.docx" -type f)
    
    if [[ ${#docx_files[@]} -eq 0 ]]; then
        echo "No DOCX files found in $INPUT"
        exit 0
    fi
    
    echo "Found ${#docx_files[@]} DOCX file(s) to convert"
    
    local converted=0
    local failed=0
    
    for docx_file in "${docx_files[@]}"; do
        # Generate relative path for output
        local rel_path
        rel_path="$(realpath --relative-to="$INPUT" "$docx_file")"
        local pdf_file="$OUTPUT/${rel_path%.docx}.pdf"
        
        if convert_file "$docx_file" "$pdf_file"; then
            ((converted++))
        else
            ((failed++))
        fi
    done
    
    echo "Conversion complete. Success: $converted, Failed: $failed"
    
else
    echo "Error: Input '$INPUT' is neither a file nor a directory"
    exit 1
fi
