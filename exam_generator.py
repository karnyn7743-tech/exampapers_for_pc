import os
import tempfile
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
import arabic_reshaper
from bidi.algorithm import get_display
import qrcode

# --------------------- إعدادات الخط العربي ---------------------
FONT_PATH = "Amiri-Regular.ttf"
try:
    pdfmetrics.registerFont(TTFont('ArabicFont', FONT_PATH))
    ARABIC_FONT = 'ArabicFont'
    print("✅ تم تحميل خط أميري")
except:
    ARABIC_FONT = 'Helvetica'
    print("⚠️ خط عربي غير موجود، سيستخدم الافتراضي")

# --------------------- دالة إعادة تشكيل النص العربي ---------------------
def reshape_arabic(text):
    reshaped = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped)
    return bidi_text

# --------------------- دالة توليد QR Code بجودة عالية ---------------------
def generate_qr_image(data, size=400):
    """توليد صورة QR Code بجودة عالية وقابلة للقراءة"""
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=20,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    if img.size != (size, size):
        img = img.resize((size, size), Image.Resampling.LANCZOS)
    return img

# --------------------- الوظيفة الرئيسية ---------------------
def generate_papers(excel_path, qr_folder, selected_class, selected_subject, output_dir):
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb.active

    material_columns = {}
    for col in range(5, 20):
        header = sheet.cell(1, col).value
        if header and str(header).strip():
            material_columns[col] = str(header).strip()

    selected_col = None
    for col, name in material_columns.items():
        if name == selected_subject:
            selected_col = col
            break

    if selected_col is None:
        messagebox.showerror("خطأ", f"المادة '{selected_subject}' غير موجودة")
        return

    students = []
    for row in range(2, sheet.max_row + 1):
        class_val = sheet.cell(row, 3).value
        if class_val and str(class_val).strip() == selected_class:
            student_id = str(sheet.cell(row, 1).value or '').strip()
            student_name = str(sheet.cell(row, 2).value or '').strip()
            student_secret = str(sheet.cell(row, 4).value or '').strip()
            if student_id and student_name and student_secret:
                students.append({
                    'id': student_id,
                    'name': student_name,
                    'secret': student_secret
                })

    if not students:
        messagebox.showwarning("تنبيه", f"لا يوجد طلاب في الصف '{selected_class}'")
        return

    os.makedirs(output_dir, exist_ok=True)

    top_margin = 4 * mm
    bottom_margin = 4 * mm
    left_margin = 10 * mm
    right_margin = 10 * mm

    pdf_filename = f"اختبار_{selected_class}_{selected_subject}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    # إعدادات المربعات الجديدة (حجم 20 ملم، وتجميع على اليسار)
    box_size = 20 * mm          # نصف الحجم السابق (40 ملم -> 20 ملم)
    spacing = 5 * mm            # مسافة بين المربعات (نصف المسافة السابقة)
    y_start = bottom_margin
    x_start = left_margin       # المربعات تبدأ من اليسار (بدلاً من التوسيط)

    for idx, student in enumerate(students):
        if idx > 0:
            c.showPage()

        student_id = student['id']
        student_name = student['name']
        student_secret = student['secret']

        # أعلى الصفحة
        c.setFont(ARABIC_FONT, 18)
        c.drawRightString(width - right_margin, height - top_margin - 5,
                          reshape_arabic(f"{student_name}  |  {student_id}"))
        c.setFont(ARABIC_FONT, 14)
        c.drawString(left_margin, height - top_margin - 10,
                     reshape_arabic(f"المادة: {selected_subject}"))

        # المربع الأول: درجة الطالب (أزرق فاتح)
        c.setFillColorRGB(0.9, 0.95, 1.0)
        c.rect(x_start, y_start, box_size, box_size, stroke=1, fill=1)
        c.setFillColorRGB(0, 0, 0)

        # المربع الثاني: QR (من الرقم السري)
        x2 = x_start + box_size + spacing
        c.rect(x2, y_start, box_size, box_size, stroke=1, fill=0)
        try:
            qr_img = generate_qr_image(student_secret, size=400)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                qr_img.save(tmp.name, 'PNG', quality=100, dpi=(300, 300))
                margin = 1 * mm   # هامش صغير جداً
                c.drawImage(tmp.name, x2 + margin, y_start + margin,
                            width=box_size - 2*margin,
                            height=box_size - 2*margin,
                            preserveAspectRatio=True)
                os.unlink(tmp.name)
        except Exception as e:
            print(f"⚠️ فشل توليد QR للطالب {student_id}: {e}")
            c.setFont('Helvetica', 8)
            c.drawCentredString(x2 + box_size/2, y_start + box_size/2 - 3, "خطأ")

        # المربع الثالث: كود المادة
        x3 = x2 + box_size + spacing
        c.rect(x3, y_start, box_size, box_size, stroke=1, fill=0)
        c.setFont(ARABIC_FONT, 14)  # حجم خط مناسب للمربع الصغير
        material_code = selected_col - 4
        c.drawCentredString(x3 + box_size/2, y_start + box_size/2 - 3, str(material_code))

    c.save()
    messagebox.showinfo("نجاح", f"تم توليد {len(students)} ورقة بنجاح في:\n{pdf_path}")

# --------------------- الواجهة الرسومية ---------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("مولد أوراق الامتحانات - مربعات صغيرة على اليسار")
        root.geometry("850x520")
        root.resizable(True, True)
        root.configure(bg='#f0f0f0')

        self.excel_path = StringVar()
        self.qr_folder = StringVar()
        self.selected_class = StringVar()
        self.selected_subject = StringVar()
        self.classes = []
        self.subjects = []

        main_frame = Frame(root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=BOTH, expand=True)

        Label(main_frame, text="ملف الإكسيل:", bg='#f0f0f0', font=('Arial', 12)).grid(row=0, column=0, sticky='w')
        Entry(main_frame, textvariable=self.excel_path, width=40, state='readonly').grid(row=0, column=1, padx=5)
        self.btn_excel = Button(main_frame, text="اختيار ملف الإكسيل", bg='#D2B48C', command=self.browse_excel)
        self.btn_excel.grid(row=0, column=2)

        Label(main_frame, text="مجلد QR (اختياري):", bg='#f0f0f0', font=('Arial', 12)).grid(row=1, column=0, sticky='w', pady=10)
        Entry(main_frame, textvariable=self.qr_folder, width=40, state='readonly').grid(row=1, column=1, padx=5)
        self.btn_qr = Button(main_frame, text="اختيار مجلد (اختياري)", bg='#FFA500', command=self.browse_qr)
        self.btn_qr.grid(row=1, column=2)

        Label(main_frame, text="الصف:", bg='#f0f0f0', font=('Arial', 12)).grid(row=2, column=0, sticky='w', pady=10)
        self.class_combo = ttk.Combobox(main_frame, textvariable=self.selected_class, values=[], state='disabled')
        self.class_combo.grid(row=2, column=1, sticky='we', padx=5)

        Label(main_frame, text="المادة:", bg='#f0f0f0', font=('Arial', 12)).grid(row=3, column=0, sticky='w', pady=10)
        self.subject_combo = ttk.Combobox(main_frame, textvariable=self.selected_subject, values=[], state='disabled')
        self.subject_combo.grid(row=3, column=1, sticky='we', padx=5)

        self.btn_generate = Button(main_frame, text="توليد أوراق الامتحانات", bg='#00008B', fg='white',
                                   font=('Arial', 14), padx=20, pady=10,
                                   command=self.generate, state='disabled')
        self.btn_generate.grid(row=4, column=1, pady=30)

        self.log_text = Text(main_frame, height=5, state='disabled', bg='white')
        self.log_text.grid(row=5, column=0, columnspan=3, sticky='we', pady=10)

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)
            self.btn_excel.config(bg='#4CAF50')
            self.load_classes_and_subjects()
            self.check_ready()

    def browse_qr(self):
        folder = filedialog.askdirectory(title="اختر مجلد (اختياري، لن يُستخدم)")
        if folder:
            self.qr_folder.set(folder)
            self.btn_qr.config(bg='#4CAF50')
            self.check_ready()

    def load_classes_and_subjects(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            return
        try:
            wb = load_workbook(path, data_only=True)
            sheet = wb.active

            classes = set()
            for row in range(2, sheet.max_row + 1):
                val = sheet.cell(row, 26).value
                if val:
                    classes.add(str(val).strip())
            self.classes = sorted(list(classes))
            self.class_combo['values'] = self.classes
            self.class_combo.config(state='readonly')

            subjects = []
            for col in range(5, 20):
                header = sheet.cell(1, col).value
                if header and str(header).strip():
                    subjects.append(str(header).strip())
            self.subjects = subjects
            self.subject_combo['values'] = self.subjects
            self.subject_combo.config(state='readonly')

        except Exception as e:
            self.log(f"❌ خطأ في قراءة البيانات: {e}")

    def check_ready(self):
        if (self.excel_path.get() and self.qr_folder.get() and
            self.selected_class.get() and self.selected_subject.get()):
            self.btn_generate.config(state='normal')
        else:
            self.btn_generate.config(state='disabled')

    def generate(self):
        excel = self.excel_path.get()
        qr = self.qr_folder.get()
        cls = self.selected_class.get()
        sub = self.selected_subject.get()

        if not all([excel, qr, cls, sub]):
            messagebox.showwarning("تنبيه", "يرجى ملء جميع الحقول")
            return

        output_dir = filedialog.askdirectory(title="اختر مجلد حفظ ملف PDF")
        if not output_dir:
            return

        self.log("⏳ جاري التوليد...")
        self.root.update()

        try:
            generate_papers(excel, qr, cls, sub, output_dir)
            self.log("✅ تم الانتهاء بنجاح")
        except Exception as e:
            self.log(f"❌ خطأ: {e}")
            messagebox.showerror("خطأ", str(e))

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        self.log_text.config(state='disabled')

    def on_class_change(self, event):
        self.check_ready()

    def on_subject_change(self, event):
        self.check_ready()

if __name__ == "__main__":
    root = Tk()
    app = App(root)
    app.class_combo.bind('<<ComboboxSelected>>', app.on_class_change)
    app.subject_combo.bind('<<ComboboxSelected>>', app.on_subject_change)
    root.mainloop()
