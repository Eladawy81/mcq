import re
import pdfplumber
import pandas as pd
from tqdm import tqdm  # لشريط التقدم

def optimized_extractor(pdf_path, excel_path):
    try:
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            for page_num, page in tqdm(enumerate(pdf.pages, 1), total=total_pages, desc="Processing"):
                text = page.extract_text()
                if not text: continue
                
                # نمط سريع للأسئلة الأساسية فقط
                matches = re.finditer(
                    r"(\d+)\.\s+(.*?)\n"
                    r"A\.\s+(.*?)\n"
                    r"B\.\s+(.*?)\n"
                    r"Answer:\s+([A-E])",
                    text, re.DOTALL
                )
                
                for match in matches:
                    all_data.append({
                        'Question': match.group(2).replace('\n', ' '),
                        'A': match.group(3),
                        'B': match.group(4),
                        'Correct': match.group(5)
                    })
        pdf.pages = pdf.pages[:1]  # معالجة الصفحة الأولى فقط
        if all_data:
            pd.DataFrame(all_data).to_excel(excel_path, index=False)
            print(f"تم استخراج {len(all_data)} سؤالًا")
        else:
            print("لا توجد نتائج")
            
    except Exception as e:
        print(f"خطأ: {str(e)}")

# التشغيل
optimized_extractor("TTTT.pdf", "Quick_Output.xlsx")