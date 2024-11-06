# Markdown to DOCX Converter

A robust Python-based tool that converts Markdown files to DOCX format with special support for LaTeX mathematical formulas. Perfect for academic papers, technical documentation, and mathematical content.

## Features

- ✨ Converts Markdown to DOCX while preserving formatting
- 🔢 Full LaTeX mathematical formula support (both inline and block)
- 📊 Table formatting with proper borders and styling
- 🔄 Clean conversion of special characters and symbols
- 📝 Proper handling of headings, paragraphs, and text formatting
- 🐳 Docker support for consistent environment

## Prerequisites

- Python 3.x
- Docker (optional, for containerized usage)

## Installation

### Local Installation

1. Clone the repository:
```bash
git clone https://github.com/adorosario/markdown_to_docx
cd md-to-docx-converter
```

2. Install required Python packages:
```bash
pip install python-docx markdown beautifulsoup4
```

### Docker Installation

Build the Docker container:
```bash
docker build -t md-to-docx .
```

## Usage

### Local Usage

Convert a Markdown file to DOCX using the command line:

```bash
python index.py -i input.md -o output.docx
```

Arguments:
- `-i, --input`: Input Markdown file path
- `-o, --output`: Output DOCX file path

### Docker Usage

#### Interactive Mode (Recommended for Development)
```bash
# Start container with bash
docker run -it --rm -v $(pwd):/app md-to-docx bash

# Inside container
python index.py -i input.md -o output.docx
```

#### Direct Usage
```bash
docker run --rm -v $(pwd):/app md-to-docx python index.py -i input.md -o output.docx
```

### Volume Mounting Notes

- Linux/Mac: Use `$(pwd)`
- Windows CMD: Use `%cd%`
- Windows PowerShell: Use `${PWD}`

## LaTeX Support

The converter supports both inline and block LaTeX formulas:

### Inline Formulas
```markdown
This is an inline formula: \(E = mc^2\)
```

### Block Formulas
```markdown
This is a block formula:
\[
\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
\]
```

## Features in Detail

### Mathematical Formula Handling
- Converts LaTeX formulas to Word equations
- Proper handling of fractions, symbols, and spacing
- Support for complex mathematical expressions

### Text Processing
- Clean conversion of special characters
- Proper spacing and formatting
- UTF-8 encoding support

### Table Support
- Converts Markdown tables to Word tables
- Applies proper borders and styling
- Maintains cell alignment and formatting

## Troubleshooting

### Common Issues

1. Permission Errors in Docker:
```bash
docker run -it --rm -v $(pwd):/app --user $(id -u):$(id -g) md-to-docx bash
```

2. Encoding Issues:
- Ensure your Markdown files are saved in UTF-8 encoding
- Check for special characters in file names

3. LaTeX Conversion Issues:
- Verify LaTeX syntax in your Markdown
- Ensure proper spacing around formulas
- Check for matching delimiters

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Python-docx library for DOCX file handling
- Markdown library for initial conversion
- BeautifulSoup4 for HTML parsing
- Docker community for containerization support
