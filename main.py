import os
import re
import pandas as pd
import pdfplumber
import flet as ft
from arabic_reshaper import reshape
from bidi.algorithm import get_display

def fix_arabic_text(text):
    if not text:
        return ""
    if re.search(r'[\u0600-\u06FF]', text):
        return get_display(reshape(text))
    return text

def main(page: ft.Page):
    page.title = "محول كشوفات PDF"
    page.rtl = True 
    page.theme_mode = ft.ThemeMode.LIGHT
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    status_text = ft.Text(value="الرجاء اختيار ملف كشف الحساب (PDF) لبدء التحويل", size=16, text_align=ft.TextAlign.CENTER)
    progress_ring = ft.ProgressRing(visible=False)

    def on_file_picked(e: ft.FilePickerResultEvent):
        if not e.files or e.files[0].path is None:
            status_text.value = "❌ لم يتم اختيار أي ملف أو أن الصلاحية مرفوضة."
            page.update()
            return
        
        picked_file = e.files[0]
        file_path = picked_file.path
        
        status_text.value = "جاري معالجة ملف الـ PDF واستخراج الحسابات..."
        progress_ring.visible = True
        page.update()
        
        final_table_rows = []
        
        try:
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

            if len(final_table_rows) > 0:
                df_final = pd.DataFrame(final_table_rows)
                if df_final.shape[1] == 5:
                    df_final.columns = ["رقم العميل", "اسم العميل", "مدين", "دائن", "الرصيد"]
                
                # مسار الحفظ في مجلد التنزيلات العام لجهاز الأندرويد
                output_dir = '/storage/emulated/0/Download'
                output_file = os.path.join(output_dir, 'كشف_مصحح_نهائي.xlsx')
                
                # إذا لم تتوفر الصلاحية للمجلد العام، يحفظ في مجلد التطبيق الخاص
                if not os.path.exists(output_dir):
                    output_file = 'كشف_مصحح_نهائي.xlsx'
                
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='الأرصدة')
                    writer.book.sheets[0].views.sheetView[0].rightToLeft = True
                
                status_text.value = f"🎉 تم التحويل بنجاح!\nالملف متوفر باسم 'كشف_مصحح_نهائي.xlsx'."
            else:
                status_text.value = "❌ لم نجد جداول متوافقة داخل ملف الـ PDF."
                
        except Exception as ex:
            status_text.value = f"خطأ أثناء المعالجة: {str(ex)}\nتأكد من إعطاء التطبيق صلاحية الملفات."
            
        progress_ring.visible = False
        page.update()

    file_picker = ft.FilePicker(on_result=on_file_picked)
    page.overlay.append(file_picker)

    # بناء عناصر الواجهة مباشرة لتفادي مشكلة الشاشة البيضاء
    page.add(
        ft.Icon(name=ft.icons.PICTURE_IN_PICTURE_ALT_ROUNDED, size=80, color=ft.colors.GREEN_700),
        ft.Text(value="نظام أتمتة كشوفات PDF المحاسبية", size=20, weight=ft.FontWeight.BOLD),
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        status_text,
        progress_ring,
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        ft.ElevatedButton(
            "📁 اختر ملف كشف الحساب (PDF)",
            icon=ft.icons.PICTURE_AS_PDF,
            on_click=lambda _: file_picker.pick_files(allow_multiple=False),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10))
        )
    )

ft.app(target=main)
