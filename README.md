# PDF to Excel Converter

A powerful Python-based tool that converts **unstructured data** from PDF files into **structured, organized Excel spreadsheets**. This tool intelligently parses raw text, tables, and mixed content from PDFs and transforms them into clean, tabular data ready for analysis.

## Features

- 📄 **Unstructured Data Processing** - Handles PDFs with raw text, mixed layouts, and non-tabular content
- 🔄 **Smart Data Structuring** - Automatically identifies patterns and converts unstructured data into structured format
- 📊 **Excel Export** - Generates clean, well-organized Excel spreadsheets (.xlsx)
- 🧹 **Intelligent Data Parsing** - Extracts key information and organizes it into rows and columns
- 📑 **Multi-page Support** - Processes complex, multi-page PDF documents
- 🎯 **Easy to Use** - Simple command-line interface or Python function calls
- ⚡ **Fast Processing** - Efficiently handles large PDF files with complex layouts

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/UzmaKhatun/PDF-to-Excel-Converter.git
cd PDF-to-Excel-Converter
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Dependencies

- **pandas** - Data manipulation and structuring
- **PyPDF2** or **pdfplumber** - PDF text extraction
- **openpyxl** - Excel file creation and manipulation
- **re** (built-in) - Pattern recognition for data parsing
- **numpy** - Data processing and transformation

## Project Structure

```
PDF-to-Excel-Converter/
├── pdf_extractor.py                 # Main conversion script
├── structured_data_evaluation.py    # Evaluation file
├── evaluation_report.txt            # Evaluation report 
├── app.py                           # full code + ui
├── Sample_Data_Input.pdf            # Example PDF files
├── output.xlsx                      # Generated Excel files
├── requirements.txt                 # Python dependencies
└── README.md                        # Project documentation
    
```

## How It Works

1. **PDF Text Extraction** - Reads and extracts all text content from the PDF file
2. **Pattern Recognition** - Identifies patterns, key-value pairs, and data structures in unstructured text
3. **Data Parsing** - Intelligently parses the extracted text to identify relevant information
4. **Structure Mapping** - Maps unstructured data into structured rows and columns
5. **Data Cleaning** - Removes duplicates, empty entries, and formats data consistently
6. **Excel Generation** - Creates a well-organized Excel file with headers and formatted cells

## Use Cases

This tool is perfect for converting:
- 📋 **Unstructured Reports** - Transform narrative reports into structured data
- 📄 **Documents with Mixed Content** - PDFs containing paragraphs, lists, and scattered information
- 📝 **Forms and Applications** - Extract field-value pairs from filled forms
- 📊 **Legacy Documents** - Modernize old documents into spreadsheet format
- 🗂️ **Data Extraction** - Pull specific information from large documents
- 📑 **Text-heavy PDFs** - Convert descriptive content into analytical format

## Supported PDF Types

- ✅ PDFs with unstructured text content
- ✅ Documents with mixed layouts (text + tables)
- ✅ Reports and narratives
- ✅ Forms with key-value pairs
- ✅ Multi-page documents with varying formats
- ✅ Text-based PDFs (non-image)
- ⚠️ Image-based/Scanned PDFs (requires OCR preprocessing)
- ⚠️ PDFs with complex graphics (focuses on text extraction)

## Troubleshooting

### Java Not Found Error
If you encounter dependency errors, ensure all packages are properly installed:
```bash
# Reinstall dependencies
pip install --upgrade pip
pip install -r requirements.txt
```

### Empty Excel Output
- Ensure your PDF contains extractable text (not just images)
- Check if the PDF is password-protected or encrypted
- Verify the PDF structure is readable
- Try adjusting parsing parameters for different document layouts

## Author

**Uzma Khatun**
- GitHub: [@UzmaKhatun](https://github.com/UzmaKhatun)

## Acknowledgments

- [PyPDF2](https://pypdf2.readthedocs.io/) / [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF text extraction
- [pandas](https://pandas.pydata.org/) - Data structuring and manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling

## Support

If you find this project helpful, please give it a ⭐ on GitHub!

For issues and feature requests, please use the [Issues](https://github.com/UzmaKhatun/PDF-to-Excel-Converter/issues) page.

---

**Note:** This tool is designed to convert **unstructured data** from PDFs into **structured Excel format**. It works best with text-based PDFs. For scanned documents or image-based PDFs, consider using OCR preprocessing tools like Tesseract before conversion.
