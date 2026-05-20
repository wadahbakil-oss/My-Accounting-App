import os
import re
import pandas as pd
import pdfplumber
import streamlit as st
from arabic_reshaper import reshape
from bidi.algorithm import get_display

# إعداد واجهة الصفحة وتوسيطها
st.set_page_config(page_title="محول كشوفات الحسابات", page_icon="📊", layout="centered")

# تنسيق المظهر ودعم اللغة العربية من اليمين لليسار
st.markdown("""
    <style>
    h1, h3, p, label { text-align: right; direction: rtl; }
    .stButton>button { width: 100%; font-weight: bold; background-color: #2e7d32; color: white; }
    div[data-testid="stFileUploadDropzone"] { direction: rtl; text-align: right; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 نظام استخراج كشوفات الـ PDF")
st.subheader("ارفع ملف كشف الحساب لتحويله فوراً إلى شيت Excel منسق")

def fix_arabic_text(text):
    if not text:
        return ""
    if re.search(r'[\u0600-\u06FF]', text):
        return get_display(reshape(text))
    return text

# حقل رفع الملف التفاعلي
uploaded_file = st.file_uploader("اختر ملف كشف الحساب بصيغة PDF:", type=["pdf"])

if uploaded_file is not None:
    final_table_rows = []
    
    with st.spinner("جاري معالجة الكشف واستخراج البيانات بدقة..."):
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if any(row):
                            clean_row = [fix_arabic_text(str(cell).strip()) if cell is not None else "" for cell in row]
                            while len(clean_row) < 5:
                                clean_row.append("")
                            final_table_rows.append(clean_row[:5])

    if len(final_table_rows) > 0:
        df_final = pd.DataFrame(final_table_rows)
        if df_final.shape[1] == 5:
            df_final.columns = ["رقم العميل", "اسم العميل", "مدين", "دائن", "الرصيد"]
            
        st.success("تمت المعالجة بنجاح خيالي!")
        
        # عرض معاينة حية للجدول والأسماء المصححة أمامك في الصفحة
        st.write("### معاينة الجدول الناتج:")
        st.dataframe(df_final, use_container_width=True)
        
        # بناء ملف الإكسل وحفظه مؤقتاً للتنزيل
        output_file = 'كشف_الحسابات_المنسق.xlsx'
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='أرصدة العملاء')
            workbook = writer.book
            worksheet = workbook['أرصدة العملاء']
            worksheet.sheet_view.rightToLeft = True
            
        with open(output_file, "rb") as file:
            st.download_button(
                label="📥 اضغط هنا لتحميل ملف Excel المنسق",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("لم نتمكن من العثور على جداول مقروءة داخل كشف الـ PDF المرفوع.")
