import PyPDF2
import json
import pandas as pd
from groq import Groq
import os
from dotenv import load_dotenv
load_dotenv()

class PDFToExcelExtractor:
    def __init__(self, api_key):
        """Initialize with API key for Groq AI service"""
        self.client = Groq(api_key=api_key)
        
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF file"""
        print(f"📄 Reading PDF: {pdf_path}")
        text = ""
        
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num, page in enumerate(pdf_reader.pages):
                text += page.extract_text()
                print(f"   ✓ Extracted page {page_num + 1}")
        
        return text
    
    def extract_structured_data(self, pdf_text):
        """Use Groq AI to extract structured key-value pairs"""
        print("\n🤖 Sending to Groq AI for extraction...")
        
        prompt = f"""You are an expert data extraction system. Your task is to extract ALL information from the following text and structure it into key-value pairs with optional comments.

CRITICAL REQUIREMENTS:
1. Extract 100% of the content - nothing should be missed
2. Identify logical key names (e.g., "First Name", "Date of Birth", "Current Salary")
3. Extract corresponding values
4. Add contextual information as comments where relevant
5. Preserve original wording from the text
6. Do NOT summarize or omit any information

Return the data as a JSON array with this structure:
[
  {{"key": "First Name", "value": "Vijay", "comments": ""}},
  {{"key": "Last Name", "value": "Kumar", "comments": ""}},
  {{"key": "Date of Birth", "value": "15-Mar-89", "comments": ""}},
  {{"key": "Age", "value": "35 years", "comments": "As on year 2024. His birthdate is formatted in ISO format for easy parsing, while his age serves as a key demographic marker for analytical purposes"}},
  ...
]

TEXT TO EXTRACT:
{pdf_text}

Return ONLY the JSON array, no additional text."""

        response = self.client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=8000,
            temperature=0.1
        )

        
        # Extract JSON from response
        response_text = response.choices[0].message.content
        
        # Clean up response (remove markdown code blocks if present)
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        print("   ✓ Received structured data from AI")
        
        try:
            data = json.loads(response_text)
            print(f"   ✓ Extracted {len(data)} key-value pairs")
            return data
        except json.JSONDecodeError as e:
            print(f"   ❌ Error parsing JSON: {e}")
            print(f"   Response: {response_text[:500]}")
            raise
    
    def create_excel(self, structured_data, output_path):
        """Create Excel file from structured data"""
        print(f"\n📊 Creating Excel file: {output_path}")
        
        # Create DataFrame
        df = pd.DataFrame(structured_data)
        
        # Add row numbers (starting from 1)
        df.insert(0, '#', range(1, len(df) + 1))
        
        # Rename columns to match expected format
        df.columns = ['#', 'Key', 'Value', 'Comments']
        
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Output', index=False)
            
            # Get the worksheet
            worksheet = writer.sheets['Output']
            
            # Adjust column widths
            worksheet.column_dimensions['A'].width = 5
            worksheet.column_dimensions['B'].width = 40
            worksheet.column_dimensions['C'].width = 35
            worksheet.column_dimensions['D'].width = 80
            
        print(f"   ✓ Excel file created with {len(df)} rows")
        print(f"   ✓ Saved to: {output_path}")
        
    def process(self, pdf_path, output_path):
        """Main processing pipeline"""
        print("=" * 60)
        print("🚀 PDF TO EXCEL EXTRACTION STARTED")
        print("=" * 60)
        
        # Step 1: Extract text from PDF
        pdf_text = self.extract_text_from_pdf(pdf_path)
        
        # Step 2: Extract structured data using AI
        structured_data = self.extract_structured_data(pdf_text)
        
        # Step 3: Create Excel file
        self.create_excel(structured_data, output_path)
        
        print("\n" + "=" * 60)
        print("✅ EXTRACTION COMPLETE!")
        print("=" * 60)
        
        return structured_data


def main():
    """Main execution function"""
    
    # Configuration
    API_KEY = os.getenv("GROQ_API_KEY")  # Set this as environment variable
    # OR hardcode for testing: API_KEY = "your-api-key-here"
    
    if not API_KEY:
        print("❌ ERROR: GROQ_API_KEY not found!")
        print("Set it as environment variable or hardcode in the script")
        return
    
    # INPUT_PDF = "Data Input.pdf"   
    INPUT_PDF = "Sample_Data_Input.pdf"   
    OUTPUT_EXCEL = "Output.xlsx"
    
    # Create extractor instance
    extractor = PDFToExcelExtractor(api_key=API_KEY)
    
    # Process the PDF
    try:
        extractor.process(INPUT_PDF, OUTPUT_EXCEL)
    except FileNotFoundError:
        print(f"❌ Error: {INPUT_PDF} not found!")
    except Exception as e:
        print(f"❌ Error during processing: {e}")


if __name__ == "__main__":
    main()