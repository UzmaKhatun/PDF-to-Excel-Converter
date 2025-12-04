# AI-Powered PDF to Excel Data Extractor

An intelligent document processing system that transforms unstructured PDF content into structured Excel spreadsheets using Claude AI.

## 🎯 Project Overview

This tool automatically extracts information from PDF documents and organizes it into a clean Key-Value-Comments format in Excel, with zero hardcoded keys. The AI determines what information to extract and how to structure it.

## ✨ Features

- **100% Data Capture**: No information is lost or summarized
- **AI-Powered Extraction**: Uses Groq AI to intelligently identify key-value pairs
- **Dynamic Key Detection**: No predefined keys - AI determines appropriate field names
- **Contextual Comments**: Extracts and preserves contextual information
- **Original Language Preservation**: Maintains exact wording from source documents

## 📋 Requirements

- Python 3.8 or higher
- Groq API key (from Anthropic)

## 🚀 Installation

1. **Clone the repository**
```bash
git clone <your-repo-url>
cd pdf-to-excel-extractor
```

2. **Install dependencies**
```bash
pip install -r requirements.txt
```

3. **Set up API key**

Option A - Environment Variable (Recommended):
```bash
export GROQ_API_KEY="your-api-key-here"
```

Option B - Hardcode in script:
Edit `pdf_extractor.py` and replace:
```python
API_KEY = "your-api-key-here"
```

## 📖 Usage

### Basic Usage

Place your PDF file in the project directory and run:

```bash
python pdf_extractor.py
```

By default, it will:
- Read from: `Data Input.pdf`
- Output to: `Output.xlsx`

### Custom File Paths

Edit the `main()` function in `pdf_extractor.py`:

```python
INPUT_PDF = "your_file.pdf"
OUTPUT_EXCEL = "your_output.xlsx"
```

### Using as a Module

```python
from pdf_extractor import PDFToExcelExtractor

# Initialize
extractor = PDFToExcelExtractor(api_key="your-api-key")

# Process
extractor.process("input.pdf", "output.xlsx")
```

## 📊 Output Format

The generated Excel file contains:

| # | Key | Value | Comments |
|---|-----|-------|----------|
| 1 | First Name | Vijay | |
| 2 | Date of Birth | 15-Mar-89 | |
| 3 | Age | 35 years | As on year 2024... |
| 4 | Birth City | Jaipur | Born and raised in the Pink City... |

## 🔧 How It Works

1. **PDF Extraction**: Reads all text from PDF using PyPDF2
2. **AI Processing**: Sends text to Groq AI with specific extraction instructions
3. **Structuring**: AI identifies key-value pairs and contextual comments
4. **Excel Generation**: Creates formatted Excel file with all extracted data

## 🎓 Using Different LLM Providers

### Switch to OpenAI (GPT-4)

1. Install OpenAI package:
```bash
pip install openai
```

2. Replace the `extract_structured_data` method:
```python
from openai import OpenAI

def extract_structured_data(self, pdf_text):
    client = OpenAI(api_key=self.api_key)
    
    response = client.chat.completions.create(
        model="gpt-4-turbo-preview",
        messages=[
            {"role": "system", "content": "You are an expert data extraction system."},
            {"role": "user", "content": f"{prompt}\n\nTEXT:\n{pdf_text}"}
        ]
    )
    
    return json.loads(response.choices[0].message.content)
```

### Switch to Google Gemini

1. Install Google package:
```bash
pip install google-generativeai
```

2. Replace the extraction method with Gemini API calls

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

## ⚠️ Important Notes

- **API Costs**: Groq API calls are free - with limited access
- **Rate Limits**: Anthropic has rate limits on API calls
- **PDF Quality**: Works best with text-based PDFs (not scanned images)
- **Data Privacy**: Ensure compliance when processing sensitive documents

## 🐛 Troubleshooting

### "GROQ_API_KEY not found"
- Set the environment variable or hardcode in script

### "Error parsing JSON"
- Check API response format
- Verify Groq API is working correctly

### Missing data in output
- Increase `max_tokens` in the API call
- Check PDF text extraction quality

### PDF reading errors
- Ensure PDF is not password-protected
- Try alternative PDF libraries (pdfplumber, pypdf)


## 🔐 Security Best Practices

- Never commit API keys to Git
- Use environment variables for sensitive data
- Add `.env` files to `.gitignore`
- Rotate API keys regularly

## Author

**Uzma Khatun**
- GitHub: [@UzmaKhatun](https://github.com/UzmaKhatun)

## Acknowledgments

- [PyPDF2](https://pypdf2.readthedocs.io/) / [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF text extraction
- [pandas](https://pandas.pydata.org/) - Data structuring and manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling

## Support

If you find this project helpful, please give it a ⭐ on GitHub!

---

**Note:** This tool is designed to convert **unstructured data** from PDFs into **structured Excel format**. It works best with text-based PDFs. For scanned documents or image-based PDFs, consider using OCR preprocessing tools like Tesseract before conversion.
