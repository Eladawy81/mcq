import re
import pandas as pd
import pdfplumber

# استخراج النص مع تصحيح الأخطاء
try:
    with pdfplumber.open("Adrenal gland MCQs.pdf") as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        print("تم استخراج النص بنجاح!")  # تأكيد استخراج النص
except Exception as e:
    print(f"خطأ في فتح الملف: {e}")
    exit()

# تعبير نمطي معدّل
pattern = re.compile(
    r"(?P<number>\d+)\.\s(?P<question>.*?)\n"
    r"(?:A\.\s(?P<a>.*?)\n)"
    r"(?:B\.\s(?P<b>.*?)\n)"
    r"(?:C\.\s(?P<c>.*?)\n)?"
    r"(?:D\.\s(?P<d>.*?)\n)?"
    r"(?:E\.\s(?P<e>.*?)\n)?"
    r"Answer:\s(?P<answer>[A-Ea-e])"
    r"(?:\nExplanation:\s(?P<explanation>.*?))?"
    r"(?=\n\d+\.|\Z)",
    re.DOTALL
)

data = []
matches = pattern.finditer(text)
if not matches:
    print("⚠️ لم يتم العثور على أسئلة! تحقق من تنسيق النص أو النمط.")
else:
    for match in matches:
        g = match.groupdict()
        try:
            row = {
                "Question": g["question"].strip(),
                "A": g.get("a", "").strip(),
                "B": g.get("b", "").strip(),
                "C": g.get("c", "").strip(),
                "D": g.get("d", "").strip(),
                "E": g.get("e", "").strip(),
                "Correct Answer": g["answer"].upper().strip(),
                "Explanation": g.get("explanation", "").strip()
            }
            data.append(row)
            print(f"تم استخراج السؤال {g['number']}")  # تأكيد استخراج كل سؤال
        except KeyError as e:
            print(f"خطأ في السؤال {g.get('number', 'غير معروف')}: الحقل {e} مفقود!")

# حفظ البيانات في Excel
if data:
    df = pd.DataFrame(data)
    df = df[["Question", "A", "B", "C", "D", "E", "Correct Answer", "Explanation"]]
    df.to_excel("Adrenal_MCQs_Output_Improved.xlsx", index=False)
    print("✅ تم التصدير بنجاح إلى ملف الإكسيل!")
else:
    print("❌ لم يتم استخراج أي بيانات!")