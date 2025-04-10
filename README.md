# CV Processing with Docker

This project converts Markdown (.md) files to Word (.docx) documents using Pandoc with a custom reference style.

## Prerequisites

- Docker installed on your machine

## Building the Docker Image

```bash
docker build -t cv-processor .
```

## Running the Container

```bash
# Run and process the markdown files
docker run --rm -v "$(pwd):/app" cv-processor

# The output files will be available in the output/ directory
```

## How it Works

1. The Docker container uses Ubuntu with Pandoc and necessary dependencies installed
2. The `process.sh` script converts all `.md` files in the root directory to `.docx` files
3. Converted files are saved to the `output/` directory
4. The custom styling comes from `custom-reference.docx`

## Files

- `process.sh`: Main script that processes markdown files
- `custom-reference.docx`: Template for styling the output documents
- `*.md`: Markdown files to be converted
- `output/`: Directory containing generated .docx files
