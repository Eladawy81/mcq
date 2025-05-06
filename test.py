import re
import pdfplumber
import pandas as pd

def extract_mcqs_to_excel(pdf_path, excel_path):
    # استخراج النص من PDF
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    
    # تعبير نمطي لاستخراج الأسئلة
    pattern = re.compile(
        r"(?P<number>\d+)\..*?\n"  # رقم السؤال
        r"(?P<question>.*?)\n"     # نص السؤال
        r"A[\.\s](?P<a>.*?)\n"     # الخيار أ
        r"B[\.\s](?P<b>.*?)\n"     # الخيار ب
        r"(?:C[\.\s](?P<c>.*?)\n)?"  # الخيار ج (اختياري)
        r"(?:D[\.\s](?P<d>.*?)\n)?"  # الخيار د (اختياري)
        r"(?:E[\.\s](?P<e>.*?)\n)?"  # الخيار هـ (اختياري)
        r"Answer:\s*(?P<answer>[A-E])"  # الإجابة
        r"(?:\nExplanation:\s*(?P<explanation>.*?))?"  # الشرح (اختياري)
        r"(?=\n\d+\.|\Z)", 
        re.DOTALL
    )

    mcqs = []
    for match in pattern.finditer(text):
        data = match.groupdict()
        
        # تنظيف البيانات
        cleaned_data = {
            "Question": data["question"].strip(),
            "A": data["a"].strip(),
            "B": data["b"].strip(),
            "C": data.get("c", "").strip(),
            "D": data.get("d", "").strip(),
            "E": data.get("e", "").strip(),
            "Correct Answer": data["answer"].strip().upper(),
            "Explanation": data.get("explanation", "").strip()
        }
        mcqs.append(cleaned_data)
    
    # حفظ في Excel
    df = pd.DataFrame(mcqs)
    df = df[["Question", "A", "B", "C", "D", "E", "Correct Answer", "Explanation"]]
    df.to_excel(excel_path, index=False)
    print(f"تم استخراج {len(mcqs)} سؤالًا بنجاح!")

if __name__ == "__main__":
    try:
        extract_mcqs_to_excel(
            pdf_path="Adrenal gland MCQs.pdf",
            excel_path="Structured_MCQs.xlsx"
        )
    except Exception as e:
        print(f"حدث خطأ: {str(e)}")