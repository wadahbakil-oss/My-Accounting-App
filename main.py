import os
import re
import pandas as pd
import easyocr
import pdfplumber
import flet as ft
from arabic_reshaper import reshape
from bidi.algorithm import get_display

# دالة تصحيح اتجاه النصوص العربية
def fix_arabic_text(text):
    if not text:
        return ""
    if re.search(r'[\u0600-\u06FF]', text):
        return get_display(reshape(text))
    return text

def main(page: ft.Page):
    # إعدادات واجهة الجوال
    page.title = "محول كشوفات الحسابات"
    page.rtl = True # تفعيل دعم اللغة العربية من اليمين لليسار
    page.theme_mode = ft.ThemeMode.LIGHT
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    status_text = ft.Text(value="الرجاء اختيار ملف (PDF أو صورة) لبدء التحويل", size=16, text_align=ft.TextAlign.CENTER)
    progress_ring = ft.ProgressRing(visible=False)

    # دالة معالجة الملف بعد اختياره من الجوال
    def on_file_picked(e: ft.FilePickerResultEvent):
        if not e.files:
            return
        
        picked_file = e.files[0]
        file_path = picked_file.path
        file_extension = os.path.splitext(file_path)[1].lower()
        
        status_text.value = "جاري معالجة الملف واستخراج البيانات..."
        progress_ring.visible = True
        page.update()
        
        final_table_rows = []
        
        try:
            # 1. إذا كان الملف PDF
            if file_extension == '.pdf':
                with pdfplumber.open(file_path) as pdf:
                    for p in pdf.pages:
                        tables = p.extract_tables()
                        for table in tables:
                            for row in table:
                                if any(row):
                                    clean_row = [fix_arabic_text(str(cell).strip()) if cell is not None else "" for cell in row]
                                    while len(clean_row) < 5:
                                        clean_row.append("")
                                    final_table_rows.append(clean_row[:5])

            # 2. إذا كان الملف صورة
            elif file_extension in ['.jpg', '.jpeg', '.png']:
                reader = easyocr.Reader(['ar', 'en'], gpu=False)
                result = reader.readtext(file_path)
                
                raw_elements = []
                if result:
                    for (bbox, text, prob) in result:
                        text = fix_arabic_text(text.strip())
                        x_center = (bbox[0][0] + bbox[1][0]) / 2
                        y_center = (bbox[0][1] + bbox[2][1]) / 2
                        raw_elements.append({'text': text, 'x': x_center, 'y': y_center})
                
                if len(raw_elements) > 0:
                    df_elements = pd.DataFrame(raw_elements).sort_values('y')
                    df_elements['row_group'] = (df_elements['y'].diff().abs() > 25).cumsum()
                    
                    for g_id, row_data in df_elements.groupby('row_group'):
                        row_data = row_data.sort_values('x', ascending=False)
                        items = row_data['text'].tolist()
                        if len(items) >= 2:
                            while len(items) < 5:
                                items.append("")
                            final_table_rows.append(items[:5])

            # 3. حفظ النتيجة في ملف إكسل على الجوال
            if len(final_table_rows) > 0:
                df_final = pd.DataFrame(final_table_rows)
                if df_final.shape[1] == 5:
                    df_final.columns = ["رقم العميل", "اسم العميل", "مدين", "دائن", "الرصيد"]
                
                # حفظ الملف في مجلد التحميلات الافتراضي أو بجانب التطبيق
                output_file = '/storage/emulated/0/Download/كشف_مصحح.xlsx' if os.path.exists('/storage/emulated/0/Download') else 'كشف_مصحح.xlsx'
                
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='الأرصدة')
                    writer.book.sheets[0].views.sheetView[0].rightToLeft = True
                
                status_text.value = f"🎉 تم التحويل بنجاح!\nتم حفظ الملف في المجلد بجهازك."
            else:
                status_text.value = "❌ تعذر استخراج بيانات محاسبية، تأكد من جودة الملف."
                
        except Exception as ex:
            status_text.value = f"خطأ أثناء المعالجة: {str(ex)}"
            
        progress_ring.visible = False
        page.update()

    # أداة اختيار الملفات من ذاكرة الهاتف
    file_picker = ft.FilePicker(on_result=on_file_picked)
    page.overlay.append(file_picker)

    # تصميم واجهة التطبيق (أزرار ونصوص)
    page.add(
        ft.Icon(name=ft.icons.ANALYTICS_ROUNDED, size=80, color=ft.colors.BLUE_ACCENT),
        ft.Text(value="نظام أتمتة كشوفات الحسابات", size=22, weight=ft.FontWeight.BOLD),
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        status_text,
        ft.Divider(height=10, color=ft.colors.TRANSPARENT),
        progress_ring,
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        ft.ElevatedButton(
            "📁 اختر ملف (PDF / صورة)",
            icon=ft.icons.UPLOAD_FILE,
            on_click=lambda _: file_picker.pick_files(allow_multiple=False, file_type=ft.FileType.ANY),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10))
        )
    )

ft.app(target=main)
