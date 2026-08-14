import os
import tempfile
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image

# --------------------- إعدادات الخط العربي (Amiri) ---------------------
# حمّل ملف Amiri-Regular.ttf من هذا الرابط:
# https://fonts.google.com/specimen/Amiri
# ثم ضعه في نفس مجلد ملف Python
FONT_PATH = "Amiri-Regular.ttf"  # <--- تم التعديل هنا

try:
    pdfmetrics.registerFont(TTFont('ArabicFont', FONT_PATH))
    ARABIC_FONT = 'ArabicFont'
    print("✅ تم تحميل خط أميري بنجاح")
except Exception as e:
    ARABIC_FONT = 'Helvetica'
    print(f"⚠️ فشل تحميل خط أميري: {e}")
    print("⚠️ سيتم استخدام الخط الافتراضي Helvetica")

# --------------------- الوظائف الأساسية ---------------------
def generate_papers(excel_path, selected_class, selected_subject, exam_name, output_dir):
    """توليد أوراق الامتحان لكل طالب"""
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb.active

    students = []
    for row in range(2, sheet.max_row + 1):
        class_cell = sheet.cell(row, 3).value  # العمود C
        if class_cell and str(class_cell).strip() == selected_class:
            student_id = str(sheet.cell(row, 1).value or '').strip()
            student_name = str(sheet.cell(row, 2).value or '').strip()
            subject = str(sheet.cell(row, 4).value or '').strip()
            if subject == selected_subject and student_id and student_name:
                students.append({'id': student_id, 'name': student_name})

    if not students:
        messagebox.showwarning("تنبيه", "لا يوجد طلاب بهذه المواصفات")
        return

    # إنشاء مجلد الإخراج
    os.makedirs(output_dir, exist_ok=True)

    # مسار مجلد QR (نفترض أنه في نفس مكان الإكسيل)
    qr_folder = os.path.join(os.path.dirname(excel_path), "QR Codes")

    for student in students:
        student_id = student['id']
        student_name = student['name']

        # إنشاء PDF
        pdf_path = os.path.join(output_dir, f"{student_id}.pdf")
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4

        # كتابة النص (باستخدام خط أميري)
        c.setFont(ARABIC_FONT, 22)
        c.drawCentredString(width/2, height-50, f"اسم الطالب: {student_name}")

        # إدراج صورة QR إن وجدت
        qr_path = os.path.join(qr_folder, f"{student_id}.png")
        if os.path.exists(qr_path):
            try:
                img = Image.open(qr_path)
                # حفظ الصورة مؤقتاً بصيغة مناسبة لـ reportlab
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    img.save(tmp.name)
                    c.drawImage(tmp.name, width/2-100, height/2-50, width=200, height=200)
                    os.unlink(tmp.name)
            except Exception as e:
                print(f"⚠️ فشل تحميل QR للطالب {student_id}: {e}")

        c.setFont(ARABIC_FONT, 18)
        c.drawCentredString(width/2, height-250, f"المادة: {selected_subject}")
        c.setFont(ARABIC_FONT, 16)
        c.drawCentredString(width/2, height-280, f"الامتحان: {exam_name}")

        c.save()
        print(f"✅ تم حفظ PDF للطالب {student_id}")

    messagebox.showinfo("نجاح", f"تم توليد {len(students)} ورقة بنجاح في:\n{output_dir}")

# --------------------- الواجهة الرسومية ---------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("مولد أوراق الامتحان - خط أميري")
        root.geometry("500x420")
        root.configure(bg='#f0f0f0')

        # متغيرات
        self.excel_path = StringVar()
        self.selected_class = StringVar()
        self.selected_subject = StringVar()
        self.exam_name = StringVar()

        # إطار رئيسي
        main_frame = Frame(root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=BOTH, expand=True)

        # 1. اختيار ملف الإكسيل
        Label(main_frame, text="ملف الإكسيل:", bg='#f0f0f0', font=('Arial', 12)).grid(row=0, column=0, sticky='w')
        Entry(main_frame, textvariable=self.excel_path, width=35).grid(row=0, column=1, padx=5)
        Button(main_frame, text="استعراض", command=self.browse_excel).grid(row=0, column=2)

        # 2. الصف
        Label(main_frame, text="الصف:", bg='#f0f0f0', font=('Arial', 12)).grid(row=1, column=0, sticky='w', pady=10)
        self.class_combo = ttk.Combobox(main_frame, textvariable=self.selected_class, values=[])
        self.class_combo.grid(row=1, column=1, sticky='we', padx=5)
        Button(main_frame, text="تحديث القائمة", command=self.load_classes).grid(row=1, column=2)

        # 3. المادة
        Label(main_frame, text="المادة:", bg='#f0f0f0', font=('Arial', 12)).grid(row=2, column=0, sticky='w', pady=10)
        self.subject_combo = ttk.Combobox(main_frame, textvariable=self.selected_subject, values=[])
        self.subject_combo.grid(row=2, column=1, sticky='we', padx=5)
        Button(main_frame, text="تحديث القائمة", command=self.load_subjects).grid(row=2, column=2)

        # 4. اسم الامتحان
        Label(main_frame, text="اسم الامتحان:", bg='#f0f0f0', font=('Arial', 12)).grid(row=3, column=0, sticky='w', pady=10)
        Entry(main_frame, textvariable=self.exam_name, width=35).grid(row=3, column=1, padx=5)

        # 5. زر التوليد
        Button(main_frame, text="توليد الأوراق", command=self.generate, bg='#4CAF50', fg='white', font=('Arial', 14), padx=20, pady=10).grid(row=4, column=1, pady=30)

        # 6. مربع رسائل
        self.log_text = Text(main_frame, height=5, state='disabled', bg='white')
        self.log_text.grid(row=5, column=0, columnspan=3, sticky='we', pady=10)

        # تحميل البيانات الأولية
        self.load_classes()
        self.load_subjects()

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)
            self.load_classes()
            self.load_subjects()

    def load_classes(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, data_only=True)
            sheet = wb.active
            classes = set()
            for row in range(2, sheet.max_row + 1):
                val = sheet.cell(row, 3).value
                if val:
                    classes.add(str(val).strip())
            self.class_combo['values'] = sorted(list(classes))
        except Exception as e:
            self.log(f"❌ خطأ في قراءة الصفوف: {e}")

    def load_subjects(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, data_only=True)
            sheet = wb.active
            subjects = set()
            for row in range(2, sheet.max_row + 1):
                val = sheet.cell(row, 4).value
                if val:
                    subjects.add(str(val).strip())
            self.subject_combo['values'] = sorted(list(subjects))
        except Exception as e:
            self.log(f"❌ خطأ في قراءة المواد: {e}")

    def generate(self):
        path = self.excel_path.get()
        cls = self.selected_class.get()
        sub = self.selected_subject.get()
        name = self.exam_name.get().strip()

        if not all([path, cls, sub, name]):
            messagebox.showwarning("تنبيه", "يرجى ملء جميع الحقول")
            return

        # اختيار مجلد الحفظ
        output_dir = filedialog.askdirectory(title="اختر مجلد حفظ الأوراق")
        if not output_dir:
            return

        self.log("⏳ جاري التوليد...")
        self.root.update()

        try:
            generate_papers(path, cls, sub, name, output_dir)
            self.log("✅ تم الانتهاء بنجاح")
        except Exception as e:
            self.log(f"❌ خطأ: {e}")
            messagebox.showerror("خطأ", str(e))

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        self.log_text.config(state='disabled')

# --------------------- تشغيل التطبيق ---------------------
if __name__ == "__main__":
    root = Tk()
    app = App(root)
    root.mainloop()
