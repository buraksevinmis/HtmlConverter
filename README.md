# HTML Converter

A tool for converting DOCX files to raw HTML with automatic FAQ schema detection and structured data markup.

## Features

- **Bulk Conversion**: Process multiple DOCX files at once
- **Excel Export**: Export HTML content in Excel format with preserved structure
- **Automatic FAQ Detection**: Intelligently detects FAQ sections in documents
- **JSON-LD Schema Support**: Adds structured data for improved SEO
- **Clean Code View**: View raw HTML alongside rendered previews

## How It Works

1. Upload one or more DOCX files
2. Process them with a single click
3. Get clean HTML output with proper structure
4. Export to Excel for easy integration and storage

## FAQ Schema Detection

The application automatically detects FAQ sections when:
- There are 3+ consecutive question headings in the last third of the document
- Headings end with a question mark (?)
- When enabled, generates proper JSON-LD schema markup for SEO benefits

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/html-converter.git
cd html-converter

# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

## Requirements

- Python 3.9+
- Streamlit
- python-docx
- pandas
- beautifulsoup4
- xlsxwriter

## About

Browser-based HTML editors often focus on individual editing but don't support bulk translation and conversion. This tool fills that gap by providing a fast, practical, and functional HTML translation/editor tool.

Developed by [Opsis Digital](https://opsisdigital.com) to speed up and automate the HTML translation process.