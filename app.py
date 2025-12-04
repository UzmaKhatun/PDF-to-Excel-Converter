import streamlit as st
import PyPDF2
import pandas as pd
from groq import Groq
import json
import io
import re
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for beautiful styling
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .upload-box {
        border: 2px dashed #4CAF50;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        background-color: rgba(255, 255, 255, 0.9);
        margin: 20px 0;
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 10px 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    </style>
""", unsafe_allow_html=True)


class PDFToExcelConverter:
    def __init__(self, api_key):
        self.client = Groq(api_key=api_key)
        
    def extract_text_from_pdf(self, pdf_file):
        """Extract text from uploaded PDF"""
        text = ""
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    
    def extract_structured_data(self, pdf_text):
        """Use Groq AI to extract structured data"""
        prompt = f"""You are an expert data extraction system. Extract ALL information from the following text and structure it into key-value pairs with optional comments.

CRITICAL REQUIREMENTS:
1. Extract 100% of the content - nothing should be missed
2. Identify logical key names (e.g., "First Name", "Date of Birth", "Current Salary")
3. Extract corresponding values
4. Add contextual information as comments where relevant
5. Preserve original wording from the text
6. Do NOT summarize or omit any information

Return ONLY a JSON array with this structure:
[
  {{"key": "First Name", "value": "Vijay", "comments": ""}},
  {{"key": "Last Name", "value": "Kumar", "comments": ""}},
  {{"key": "Age", "value": "35 years", "comments": "As on year 2024"}},
  ...
]

TEXT TO EXTRACT:
{pdf_text}

Return ONLY the JSON array, no additional text."""

        response = self.client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=8000
        )
        
        response_text = response.choices[0].message.content
        
        # Clean response
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        return json.loads(response_text)

#     def extract_structured_data(self, pdf_text):
#         """Use Groq AI to extract structured data"""
#         prompt = f"""You are an expert data extraction system. Extract ALL information from the following text and structure it into key-value pairs with optional comments.

# CRITICAL REQUIREMENTS:
# 1. Extract 100% of the content - nothing should be missed
# 2. Identify logical key names (e.g., "First Name", "Date of Birth", "Current Salary")
# 3. Extract corresponding values
# 4. Add contextual information as comments where relevant
# 5. Preserve original wording from the text
# 6. Do NOT summarize or omit any information

# Return ONLY a JSON array with this structure:
# [
#   {{"key": "First Name", "value": "Vijay", "comments": ""}},
#   {{"key": "Last Name", "value": "Kumar", "comments": ""}},
#   {{"key": "Age", "value": "35 years", "comments": "As on year 2024"}},
#   ...
# ]

# TEXT TO EXTRACT:
# {pdf_text}

# Return ONLY the JSON array, no additional text."""

#         try:
#             response = self.client.chat.completions.create(
#                 model="llama-3.3-70b-versatile",
#                 messages=[{"role": "user", "content": prompt}],
#                 temperature=0.1,
#                 max_tokens=8000
#             )
        
#             response_text = response.choices[0].message.content
            
#             # Clean response
#             if "```json" in response_text:
#                 response_text = response_text.split("```json")[1].split("```")[0].strip()
#             elif "```" in response_text:
#                 response_text = response_text.split("```")[1].split("```")[0].strip()
            
#             return json.loads(response_text)
        
#         except Exception as e:
#             error_msg = str(e)
            
#             # Convert API errors to user-friendly messages
#             if "invalid_api_key" in error_msg or "401" in error_msg:
#                 raise Exception("❌ Invalid API Key! Please check your Groq API key and try again. Get a valid key from: https://console.groq.com")
#             elif "rate_limit" in error_msg or "429" in error_msg:
#                 raise Exception("⏳ Rate limit reached! Please wait a moment and try again.")
#             elif "timeout" in error_msg:
#                 raise Exception("⏱️ Request timed out! The document might be too large. Try with a smaller PDF.")
#             else:
#                 raise Exception(f"🔴 AI Processing Error: {error_msg}")
    
    def create_excel(self, structured_data):
        """Create Excel file from structured data"""
        df = pd.DataFrame(structured_data)
        df.insert(0, '#', range(1, len(df) + 1))
        df.columns = ['#', 'Key', 'Value', 'Comments']
        return df


class DataEvaluator:
    def __init__(self, df, pdf_text):
        self.df = df
        self.pdf_text = pdf_text
        
    def evaluate_completeness(self):
        """Calculate completeness percentage"""
        output_text = ""
        for _, row in self.df.iterrows():
            output_text += str(row['Value']) + " " + str(row['Comments']) + " "
        
        # Check numbers
        pdf_numbers = re.findall(r'\b\d+[\d\.,]*\b', self.pdf_text)
        numbers_found = sum(1 for n in pdf_numbers if n in output_text)
        number_score = (numbers_found / len(pdf_numbers) * 100) if pdf_numbers else 100
        
        # Check important words
        pdf_words = re.findall(r'\b[A-Z][a-z]+\b', self.pdf_text)
        words_found = sum(1 for w in pdf_words if w in output_text)
        word_score = (words_found / len(pdf_words) * 100) if pdf_words else 100
        
        return (number_score + word_score) / 2
    
    def evaluate_structure(self):
        """Evaluate structure quality"""
        score = 100
        
        # Check for nulls
        if self.df['Key'].isna().sum() > 0:
            score -= 20
        if self.df['Value'].isna().sum() > 0:
            score -= 20
            
        # Check duplicates
        if self.df.duplicated(subset=['Key']).sum() > 0:
            score -= 10
            
        return max(0, score)
    
    def evaluate_keys(self):
        """Evaluate key quality"""
        keys = self.df['Key'].tolist()
        good_patterns = ['name', 'date', 'birth', 'age', 'salary', 'education', 
                        'certification', 'skill', 'organization', 'designation']
        
        meaningful = sum(1 for k in keys if any(p in str(k).lower() for p in good_patterns))
        return (meaningful / len(keys) * 100) if keys else 0
    
    def get_overall_score(self):
        """Calculate overall quality score"""
        completeness = self.evaluate_completeness()
        structure = self.evaluate_structure()
        key_quality = self.evaluate_keys()
        
        overall = (completeness * 0.5) + (structure * 0.3) + (key_quality * 0.2)
        return round(overall, 1)


def main():
    # Header
    st.markdown("""
        <h1 style='text-align: center; color: white; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);'>
            📄 PDF to Excel Converter
        </h1>
        <h3 style='text-align: center; color: white; margin-bottom: 30px;'>
            AI-Powered Document Structuring System
        </h3>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/pdf-2.png", width=80)
        st.title("⚙️ Settings")
        
        api_key = st.text_input(
            "🔑 Groq API Key",
            type="password",
            help="Get your free API key from console.groq.com"
        )
        
        st.markdown("---")
        st.markdown("""
        ### 📋 How it works:
        1. Upload your PDF file
        2. AI extracts all data
        3. Get structured Excel output
        4. Download the result
        
        ### 🎯 Features:
        - ✅ 100% Data Capture
        - ✅ Smart Key Detection
        - ✅ Quality Evaluation
        - ✅ Instant Download
        """)
        
        st.markdown("---")
        st.info("💡 **Tip:** Make sure your PDF contains readable text (not scanned images)")
    
    # Main content
    if not api_key:
        st.warning("⚠️ Please enter your Groq API key in the sidebar to continue")
        st.info("🆓 Get a free API key at: https://console.groq.com")
        return
    
    # File upload section
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        uploaded_file = st.file_uploader(
            "📤 Upload PDF Document",
            type=['pdf'],
            help="Upload a PDF file containing text data"
        )
    
    if uploaded_file:
        # Create two columns for layout
        left_col, right_col = st.columns([1, 1])
        
        with left_col:
            st.markdown("### 📄 Input Document")
            st.info(f"**Filename:** {uploaded_file.name}")
            st.info(f"**Size:** {uploaded_file.size / 1024:.2f} KB")
            
            if st.button("🚀 Start Extraction", use_container_width=True, type="primary"):
                with st.spinner("🔄 Processing document..."):
                    try:
                        # Initialize converter
                        converter = PDFToExcelConverter(api_key)
                        
                        # Progress bar
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Step 1: Extract text
                        status_text.text("📖 Reading PDF...")
                        progress_bar.progress(25)
                        pdf_text = converter.extract_text_from_pdf(uploaded_file)
                        
                        # Step 2: AI processing
                        status_text.text("🤖 AI is analyzing document...")
                        progress_bar.progress(50)
                        structured_data = converter.extract_structured_data(pdf_text)
                        
                        # Step 3: Create Excel
                        status_text.text("📊 Creating Excel file...")
                        progress_bar.progress(75)
                        df = converter.create_excel(structured_data)
                        
                        # Step 4: Evaluate
                        status_text.text("✅ Evaluating quality...")
                        progress_bar.progress(100)
                        evaluator = DataEvaluator(df, pdf_text)
                        overall_score = evaluator.get_overall_score()
                        
                        # Store in session state
                        st.session_state['df'] = df
                        st.session_state['score'] = overall_score
                        st.session_state['completeness'] = evaluator.evaluate_completeness()
                        st.session_state['structure'] = evaluator.evaluate_structure()
                        st.session_state['key_quality'] = evaluator.evaluate_keys()
                        
                        status_text.empty()
                        progress_bar.empty()
                        
                        st.success("✅ Extraction completed successfully!")
                        
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
        
        # Display results if available
        if 'df' in st.session_state:
            with right_col:
                st.markdown("### 📊 Quality Metrics")
                
                # Overall score with color
                score = st.session_state['score']
                if score >= 90:
                    color = "green"
                    grade = "A+"
                elif score >= 80:
                    color = "lightgreen"
                    grade = "A"
                elif score >= 70:
                    color = "orange"
                    grade = "B"
                else:
                    color = "red"
                    grade = "C"
                
                st.markdown(f"""
                    <div style='background-color: {color}; padding: 20px; border-radius: 10px; text-align: center;'>
                        <h1 style='color: white; margin: 0;'>{score}%</h1>
                        <h3 style='color: white; margin: 0;'>Grade: {grade}</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Detailed metrics
                col_a, col_b, col_c = st.columns(3)
                
                with col_a:
                    st.metric(
                        "📋 Completeness",
                        f"{st.session_state['completeness']:.1f}%"
                    )
                
                with col_b:
                    st.metric(
                        "🏗️ Structure",
                        f"{st.session_state['structure']:.1f}%"
                    )
                
                with col_c:
                    st.metric(
                        "🔑 Key Quality",
                        f"{st.session_state['key_quality']:.1f}%"
                    )
            
            # Full width sections
            st.markdown("---")
            
            # Preview section
            st.markdown("### 👀 Data Preview")
            st.dataframe(
                st.session_state['df'],
                use_container_width=True,
                height=400
            )
            
            # Statistics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📊 Total Rows", len(st.session_state['df']))
            
            with col2:
                st.metric("🔑 Unique Keys", st.session_state['df']['Key'].nunique())
            
            with col3:
                comments_count = st.session_state['df'][
                    st.session_state['df']['Comments'].notna() & 
                    (st.session_state['df']['Comments'] != '')
                ].shape[0]
                st.metric("💬 Comments", comments_count)
            
            with col4:
                st.metric("📈 Data Points", len(st.session_state['df']) * 3)
            
            # Download section
            st.markdown("---")
            st.markdown("### 📥 Download Results")
            
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col2:
                # Create Excel file in memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state['df'].to_excel(writer, sheet_name='Output', index=False)
                output.seek(0)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"structured_output_{timestamp}.xlsx"
                
                st.download_button(
                    label="⬇️ Download Excel File",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                st.success("✅ Ready to download!")
    
    else:
        # Welcome message when no file uploaded
        st.markdown("""
            <div style='text-align: center; padding: 50px; background: rgba(255,255,255,0.15); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.2); border-radius: 15px; margin: 50px; box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);'>
                <h2 style='color: white;'>👋 Welcome to PDF to Excel Converter!</h2>
                <p style='font-size: 18px; color: rgba(255,255,255,0.9);'>
                    Upload a PDF document to get started with AI-powered data extraction
                </p>
                <br>
                <p style='font-size: 16px; color: white;'>
                    📄 → 🤖 → 📊 → ⬇️
                </p>
            </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()