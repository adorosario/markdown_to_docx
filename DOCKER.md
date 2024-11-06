# Markdown to DOCX Converter with LaTeX Support

This Docker container provides an environment for converting Markdown files to DOCX format while properly handling LaTeX mathematical formulas.

## Features
- Converts Markdown to DOCX format
- Supports LaTeX mathematical formulas
- Handles both inline and block math notation
- Properly formats tables with borders
- Full TeXLive support for mathematical rendering

## Quick Start

### Build the Container
```bash
docker build -t md-to-docx .
```

### Interactive Usage (Recommended for Development)
1. Start container with bash:
```bash
docker run -it --rm -v $(pwd):/app md-to-docx bash
```

2. Once inside the container, convert your files:
```bash
python index.py -i formulas.md -o output.docx
```

### Direct Usage
Convert files directly (without entering the container):
```bash
docker run --rm -v $(pwd):/app md-to-docx python index.py -i formulas.md -o output.docx
```

## Volume Mounting
The `-v $(pwd):/app` flag mounts your current directory to `/app` in the container. Ensure your Markdown files are in your current directory.

For different operating systems:
- Windows (CMD): Use `%cd%` instead of `$(pwd)`
- Windows (PowerShell): Use `${PWD}` instead of `$(pwd)`

## Included Packages
- Python 3 (latest)
- pandoc
- texlive-latex-base
- texlive-fonts-recommended
- pypandoc
- python-docx

## Examples

### Basic Usage
```bash
# Start interactive shell
docker run -it --rm -v $(pwd):/app md-to-docx bash

# Convert a file
python index.py -i formulas.md -o output.docx
```

### Testing LaTeX Support
```markdown
# Input example.md
This is an inline formula: \(E = mc^2\)

This is a block formula:
\[
\int_{0}^{\infty} e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
\]
```

## Troubleshooting

1. If you get permission errors:
```bash
docker run -it --rm -v $(pwd):/app --user $(id -u):$(id -g) md-to-docx bash
```

2. To verify the environment inside container:
```bash
docker run -it --rm md-to-docx bash
python -c "import pypandoc; print(pypandoc.__version__)"
pandoc --version
```

3. For conversion errors, check:
   - File permissions
   - Input file encoding (should be UTF-8)
   - LaTeX syntax in your Markdown

## Notes
- The container includes full LaTeX support for mathematical formulas
- All files are processed in the mounted `/app` directory
- Temporary files are handled securely and cleaned up automatically
- Output files will have the same owner as your host user when using `--user`