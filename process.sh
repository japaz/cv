#!/bin/bash 
set -euo pipefail

# Create output directory if it doesn't exist
mkdir -p output

# Process all Markdown files in current directory
for file in *.md; do
    # Get base filename without extension
    base_name=$(basename "$file" .md)
    
    # Run pandoc command with specified settings
    pandoc "$file" -o "output/${base_name}.docx" --reference-doc=custom-reference.docx

    echo "Generated: output/${base_name}.docx"
done

echo "Conversion complete. docxs are in the output/ directory."
