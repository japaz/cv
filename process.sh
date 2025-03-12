#!/bin/bash

# Create output directory if it doesn't exist
mkdir -p output

# Process all Markdown files in current directory
for file in *.md; do
    # Get base filename without extension
    base_name=$(basename "$file" .md)
    
    # Run pandoc command with specified settings
    pandoc "$file" -o "output/${base_name}.pdf" \
        -V geometry:margin=0.5in \
        -V mainfont="Arial" \
        -V fontsize=11pt \
        -V papersize=a4 \
        --pdf-engine=xelatex

    echo "Generated: output/${base_name}.pdf"
done

echo "Conversion complete. PDFs are in the output/ directory."
