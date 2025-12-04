"""
Standalone Evaluation Script - Works WITHOUT Expected Output
Evaluates extraction quality based on PDF content only
"""

import pandas as pd
import PyPDF2
import re
from collections import Counter

class StandaloneEvaluator:
    def __init__(self, generated_excel, input_pdf):
        self.generated_excel = generated_excel
        self.input_pdf = input_pdf
        self.pdf_text = ""
        self.df = None
        
    def extract_pdf_text(self):
        """Extract all text from PDF"""
        print("📄 Reading input PDF...")
        text = ""
        try:
            with open(self.input_pdf, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text()
            self.pdf_text = text
            print(f"   ✓ Extracted {len(text)} characters from PDF")
            return text
        except Exception as e:
            print(f"   ❌ Error reading PDF: {e}")
            return ""
    
    def load_excel(self):
        """Load generated Excel file"""
        print(f"\n📊 Loading Excel: {self.generated_excel}")
        try:
            df = pd.read_excel(self.generated_excel, sheet_name='Output')
            self.df = df
            print(f"   ✓ Loaded {len(df)} rows")
            return df
        except Exception as e:
            print(f"   ❌ Error loading Excel: {e}")
            return None
    
    def check_structure(self):
        """Validate Excel structure"""
        print("\n" + "="*70)
        print("1️⃣  STRUCTURE VALIDATION")
        print("="*70)
        
        score = 0
        max_score = 20
        
        # Check columns
        expected_cols = ['#', 'Key', 'Value', 'Comments']
        if list(self.df.columns) == expected_cols:
            print("✅ Columns correct: ['#', 'Key', 'Value', 'Comments']")
            score += 5
        else:
            print(f"❌ Columns: {list(self.df.columns)}")
            print(f"   Expected: {expected_cols}")
        
        # Check row numbering
        if self.df['#'].tolist() == list(range(1, len(self.df) + 1)):
            print(f"✅ Row numbers sequential (1 to {len(self.df)})")
            score += 5
        else:
            print("❌ Row numbering incorrect")
        
        # Check for nulls
        null_keys = self.df['Key'].isna().sum()
        null_values = self.df['Value'].isna().sum()
        
        if null_keys == 0:
            print("✅ No null keys")
            score += 5
        else:
            print(f"❌ Found {null_keys} null keys")
        
        if null_values == 0:
            print("✅ No null values")
            score += 5
        else:
            print(f"❌ Found {null_values} null values")
        
        print(f"\n📊 Structure Score: {score}/{max_score}")
        return score
    
    def check_data_loss(self):
        """Check if data from PDF is lost"""
        print("\n" + "="*70)
        print("2️⃣  DATA COMPLETENESS CHECK")
        print("="*70)
        
        score = 0
        max_score = 30
        
        # Extract all output text
        output_text = ""
        for _, row in self.df.iterrows():
            output_text += str(row['Value']) + " "
            output_text += str(row['Comments']) + " "
        
        # Find all numbers in PDF (dates, ages, salaries, scores, etc.)
        pdf_numbers = re.findall(r'\b\d+[\d\.,]*\b', self.pdf_text)
        output_numbers = re.findall(r'\b\d+[\d\.,]*\b', output_text)
        
        # Find all important words (capitalized, likely proper nouns)
        pdf_words = re.findall(r'\b[A-Z][a-z]+\b', self.pdf_text)
        output_words = re.findall(r'\b[A-Z][a-z]+\b', output_text)
        
        # Calculate coverage
        numbers_found = len([n for n in pdf_numbers if n in output_text])
        number_coverage = (numbers_found / len(pdf_numbers) * 100) if pdf_numbers else 100
        
        words_found = len([w for w in pdf_words if w in output_text])
        word_coverage = (words_found / len(pdf_words) * 100) if pdf_words else 100
        
        print(f"📈 Number Coverage:")
        print(f"   - PDF has {len(pdf_numbers)} numbers")
        print(f"   - Output captured {numbers_found} numbers")
        print(f"   - Coverage: {number_coverage:.1f}%")
        
        if number_coverage >= 90:
            score += 15
            print("   ✅ Excellent number capture!")
        elif number_coverage >= 75:
            score += 12
            print("   ✅ Good number capture")
        else:
            score += 8
            print("   ⚠️  Some numbers missing")
        
        print(f"\n📈 Named Entity Coverage:")
        print(f"   - PDF has {len(pdf_words)} capitalized words")
        print(f"   - Output captured {words_found} words")
        print(f"   - Coverage: {word_coverage:.1f}%")
        
        if word_coverage >= 85:
            score += 15
            print("   ✅ Excellent entity capture!")
        elif word_coverage >= 70:
            score += 12
            print("   ✅ Good entity capture")
        else:
            score += 8
            print("   ⚠️  Some entities missing")
        
        # Character count ratio
        pdf_chars = len(self.pdf_text)
        output_chars = len(output_text)
        char_ratio = (output_chars / pdf_chars * 100) if pdf_chars > 0 else 0
        
        print(f"\n📊 Character Coverage: {char_ratio:.1f}%")
        print(f"   - PDF: {pdf_chars} chars")
        print(f"   - Output: {output_chars} chars")
        
        print(f"\n📊 Completeness Score: {score}/{max_score}")
        return score
    
    def check_key_quality(self):
        """Evaluate quality of extracted keys"""
        print("\n" + "="*70)
        print("3️⃣  KEY QUALITY ANALYSIS")
        print("="*70)
        
        score = 0
        max_score = 25
        
        keys = self.df['Key'].tolist()
        
        # Check for meaningful patterns
        good_patterns = [
            'name', 'date', 'birth', 'age', 'salary', 'organization',
            'designation', 'role', 'education', 'degree', 'college',
            'certification', 'skill', 'proficiency', 'school', 'year',
            'joining', 'current', 'previous', 'score', 'grade'
        ]
        
        meaningful_count = 0
        for key in keys:
            key_lower = str(key).lower()
            if any(pattern in key_lower for pattern in good_patterns):
                meaningful_count += 1
        
        meaningful_percent = (meaningful_count / len(keys) * 100) if keys else 0
        print(f"📊 Meaningful Keys: {meaningful_count}/{len(keys)} ({meaningful_percent:.1f}%)")
        
        if meaningful_percent >= 80:
            score += 10
            print("   ✅ Excellent key naming!")
        elif meaningful_percent >= 60:
            score += 7
            print("   ✅ Good key naming")
        else:
            score += 4
            print("   ⚠️  Keys could be more descriptive")
        
        # Check for duplicates
        duplicates = self.df[self.df.duplicated(subset=['Key'], keep=False)]
        if len(duplicates) == 0:
            print("✅ No duplicate keys")
            score += 5
        else:
            print(f"❌ Found {len(duplicates)} duplicate keys")
            score += 2
        
        # Check key formatting (Title Case)
        properly_formatted = sum(1 for k in keys if str(k) and str(k)[0].isupper())
        format_percent = (properly_formatted / len(keys) * 100) if keys else 0
        
        print(f"📊 Proper Formatting: {properly_formatted}/{len(keys)} ({format_percent:.1f}%)")
        
        if format_percent >= 90:
            score += 5
            print("   ✅ Well formatted keys")
        elif format_percent >= 70:
            score += 3
            print("   ✅ Acceptable formatting")
        else:
            score += 1
            print("   ⚠️  Inconsistent formatting")
        
        # Check average key length (should be descriptive but not too long)
        avg_key_length = sum(len(str(k)) for k in keys) / len(keys) if keys else 0
        print(f"\n📊 Average Key Length: {avg_key_length:.1f} characters")
        
        if 15 <= avg_key_length <= 40:
            score += 5
            print("   ✅ Good key length (descriptive)")
        elif 10 <= avg_key_length <= 50:
            score += 3
            print("   ✅ Acceptable key length")
        else:
            score += 1
            print("   ⚠️  Keys too short or too long")
        
        print(f"\n📊 Key Quality Score: {score}/{max_score}")
        return score
    
    def check_value_quality(self):
        """Check quality of extracted values"""
        print("\n" + "="*70)
        print("4️⃣  VALUE QUALITY ANALYSIS")
        print("="*70)
        
        score = 0
        max_score = 15
        
        values = self.df['Value'].tolist()
        
        # Check for empty values
        empty_values = sum(1 for v in values if pd.isna(v) or str(v).strip() == '')
        if empty_values == 0:
            print("✅ No empty values")
            score += 5
        else:
            print(f"⚠️  Found {empty_values} empty values")
            score += 2
        
        # Check value diversity (not all same)
        unique_values = len(set(str(v) for v in values))
        diversity_percent = (unique_values / len(values) * 100) if values else 0
        
        print(f"📊 Value Diversity: {unique_values}/{len(values)} unique ({diversity_percent:.1f}%)")
        
        if diversity_percent >= 85:
            score += 5
            print("   ✅ Good value diversity")
        elif diversity_percent >= 70:
            score += 3
            print("   ✅ Acceptable diversity")
        else:
            score += 1
            print("   ⚠️  Low diversity (possible duplicates)")
        
        # Check average value length
        avg_value_length = sum(len(str(v)) for v in values) / len(values) if values else 0
        print(f"\n📊 Average Value Length: {avg_value_length:.1f} characters")
        
        if avg_value_length >= 5:
            score += 5
            print("   ✅ Values have good detail")
        else:
            score += 2
            print("   ⚠️  Values seem too short")
        
        print(f"\n📊 Value Quality Score: {score}/{max_score}")
        return score
    
    def check_comments(self):
        """Evaluate comments column"""
        print("\n" + "="*70)
        print("5️⃣  COMMENTS ANALYSIS")
        print("="*70)
        
        score = 0
        max_score = 10
        
        # Count non-empty comments
        non_empty = self.df[self.df['Comments'].notna() & (self.df['Comments'] != '')].shape[0]
        comment_percent = (non_empty / len(self.df) * 100) if len(self.df) > 0 else 0
        
        print(f"📊 Comments Usage: {non_empty}/{len(self.df)} rows ({comment_percent:.1f}%)")
        
        if comment_percent >= 30:
            score += 10
            print("   ✅ Good use of contextual comments")
        elif comment_percent >= 15:
            score += 7
            print("   ✅ Moderate comment usage")
        elif comment_percent >= 5:
            score += 4
            print("   ⚠️  Limited comment usage")
        else:
            score += 2
            print("   ⚠️  Very few comments (context may be lost)")
        
        print(f"\n📊 Comments Score: {score}/{max_score}")
        return score
    
    def generate_report(self):
        """Generate complete evaluation report"""
        print("\n" + "="*70)
        print("🎯 STANDALONE EVALUATION REPORT")
        print("   (No expected output needed)")
        print("="*70)
        
        # Load data
        # self.extract_pdf_text()
        # if not self.load_excel():
        #     print("❌ Cannot load Excel file. Evaluation stopped.")
        #     return
        
        df = self.load_excel()
        if df is None:
            print("Cannot load")

        # Run all checks
        total_score = 0
        total_score += self.check_structure()       # 20 points
        total_score += self.check_data_loss()       # 30 points
        total_score += self.check_key_quality()     # 25 points
        total_score += self.check_value_quality()   # 15 points
        total_score += self.check_comments()        # 10 points
        
        # Final summary
        print("\n" + "="*70)
        print("📊 FINAL EVALUATION SUMMARY")
        print("="*70)
        
        max_possible = 100
        print(f"\n🎯 Total Score: {total_score}/{max_possible}")
        
        # Grade assignment
        if total_score >= 90:
            grade = "A+ Excellent!"
            emoji = "🏆"
            feedback = "Outstanding extraction! Production ready."
        elif total_score >= 80:
            grade = "A Very Good"
            emoji = "🌟"
            feedback = "Great work! Minor improvements possible."
        elif total_score >= 70:
            grade = "B Good"
            emoji = "✅"
            feedback = "Solid extraction. Some refinements needed."
        elif total_score >= 60:
            grade = "C Satisfactory"
            emoji = "⚠️"
            feedback = "Basic requirements met. Needs improvement."
        else:
            grade = "D Needs Work"
            emoji = "❌"
            feedback = "Significant improvements required."
        
        print(f"{emoji} Grade: {grade}")
        print(f"💬 Feedback: {feedback}")
        
        # Breakdown
        print(f"\n📋 Score Breakdown:")
        print(f"   • Structure Validation (20 pts)")
        print(f"   • Data Completeness (30 pts)")
        print(f"   • Key Quality (25 pts)")
        print(f"   • Value Quality (15 pts)")
        print(f"   • Comments Usage (10 pts)")
        
        # Recommendations
        print(f"\n💡 Recommendations:")
        if total_score < 100:
            if total_score < 70:
                print("   • Review PDF text extraction - ensure all content captured")
                print("   • Improve key naming to be more descriptive")
                print("   • Add more contextual comments where relevant")
            elif total_score < 90:
                print("   • Fine-tune key names for better clarity")
                print("   • Consider adding more contextual information to comments")
        else:
            print("   • Perfect score! Great job! 🎉")
        
        print("\n" + "="*70)
        
        # Save report
        self.save_report(total_score, grade, feedback)
        
        return total_score
    
    def save_report(self, score, grade, feedback):
        """Save report to file"""
        with open("evaluation_report.txt", "w", encoding='utf-8') as f:
            f.write("="*70 + "\n")
            f.write("STANDALONE EVALUATION REPORT\n")
            f.write("="*70 + "\n\n")
            f.write(f"Input PDF: {self.input_pdf}\n")
            f.write(f"Generated Excel: {self.generated_excel}\n\n")
            f.write(f"Total Score: {score}/100\n")
            f.write(f"Grade: {grade}\n")
            f.write(f"Feedback: {feedback}\n\n")
            f.write(f"Evaluation Date: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        print(f"\n💾 Report saved to: evaluation_report.txt")


def main():
    """Main execution"""
    
    INPUT_PDF = "Sample_Data_Input.pdf"
    GENERATED_OUTPUT = "Output.xlsx"
    
    print("🚀 Starting Standalone Evaluation...")
    print("   (No expected output file needed!)\n")
    
    evaluator = StandaloneEvaluator(GENERATED_OUTPUT, INPUT_PDF)
    score = evaluator.generate_report()
    
    print(f"\n✅ Evaluation Complete! Your score: {score}/100\n")


if __name__ == "__main__":
    main()
