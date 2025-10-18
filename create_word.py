from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io

app = Flask(__name__)

# ---------------------- Persian Text Processor ----------------------
class PersianTextProcessor:
    """پردازشگر دقیق متن فارسی بدون حذف فرمول"""

    def clean_text(self, text):
        # محافظت از فرمول‌ها
        formulas = re.findall(r'\$\$.*?\$\$|\$.*?\$', text)
        for i, f in enumerate(formulas):
            text = text.replace(f, f"§§{i}§§")

        # نرمال‌سازی نویسه‌ها
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه').replace('ؤ', 'و')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)

        # نیم‌فاصله‌ها
        prefixes = ['می', 'نمی', 'بی', 'به', 'در', 'که']
        for p in prefixes:
            text = re.sub(f'\\b{p} ', f'{p}\u200c', text)

        # بازگرداندن فرمول‌ها بدون تغییر
        for i, f in enumerate(formulas):
            text = text.replace(f"§§{i}§§", f)
        return text.strip()

# ---------------------- Smart Word Generator ----------------------
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_doc()

    def _setup_doc(self):
        s = self.doc.sections[0]
        s.page_height = Inches(11.69)  # A4
        s.page_width = Inches(8.27)
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Inches(1)

    def _set_rtl(self, p):
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    # ---------- نوع خط ----------
    def detect_content_type(self, line):
        line = line.strip()
        if not line:
            return 'empty'
        if re.match(r'^#+', line):
            return 'heading'
        if re.search(r'\$\$.*?\$\$|\$.*?\$', line):
            return 'formula'
        if re.match(r'^شکل\s*\d+', line):
            return 'figure_caption'
        if re.match(r'^جدول\s*\d+', line):
            return 'table_caption'
        return 'text'

    # ---------- اجزای سند ----------
    def add_heading(self, text, level=1):
        text = re.sub(r'^#+\s*', '', text)
        h = self.doc.add_heading(level=level)
        r = h.add_run(text)
        r.bold = True
        r.font.name = 'B Nazanin'
        r.font.size = Pt(18 - level * 2)
        r._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(h)

    def add_formula(self, text):
        """درج فرمول خام، بدون حذف یا تغییر"""
        formulas = re.findall(r'\$\$.*?\$\$|\$.*?\$', text)
        for f in formulas:
            f = f.strip('$').strip()
            p = self.doc.add_paragraph(f)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.name = 'Cambria Math'
            r.font.size = Pt(14)

    def add_caption(self, text):
        """افزودن عنوان شکل یا جدول"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        r = p.add_run(text)
        r.bold = True
        r.font.name = 'B Nazanin'
        r.font.size = Pt(13)
        r._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    def add_text(self, text):
        """پاراگراف معمولی با متن فارسی و انگلیسی"""
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        self._set_rtl(p)
        parts = re.split(r'([A-Za-z0-9,;:.()\[\]{}=+\-*/^%<>])', text)
        for part in parts:
            if not part.strip():
                continue
            if re.match(r'[A-Za-z0-9]', part):
                r = p.add_run(part)
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)
            else:
                r = p.add_run(part)
                r.font.name = 'B Nazanin'
                r.font.size = Pt(14)
                r._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------- پردازش کلی متن ----------
    def process_text(self, text):
        lines = [ln.rstrip() for ln in text.split('\n')]
        clean_lines = []
        for ln in lines:
            # حذف خطوط خالی متوالی
            if ln == '' and (not clean_lines or clean_lines[-1] == ''):
                continue
            clean_lines.append(ln)

        previous_type = None
        for line in clean_lines:
            t = self.detect_content_type(line)
            if t == 'empty':
                # فقط یه بار خط خالی بین دو پاراگراف عادی
                if previous_type in ['text', 'formula']:
                    self.doc.add_paragraph()
            elif t == 'heading':
                self.add_heading(line)
            elif t == 'formula':
                self.add_formula(line)
            elif t in ['figure_caption', 'table_caption']:
                self.add_caption(line)
            else:
                self.add_text(line)
            previous_type = t

    def save_to_stream(self):
        f = io.BytesIO()
        self.doc.save(f)
        f.seek(0)
        return f

# ---------------------- API ----------------------
@app.route('/generate', methods=['POST'])
def generate_word():
    try:
        data = request.get_json()
        text = data.get('text', '')
        if not text:
            return jsonify({'error': 'متن الزامی است'}), 400
        gen = SmartDocumentGenerator()
        gen.process_text(text)
        fstream = gen.save_to_stream()
        return send_file(
            fstream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='document.docx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/')
def home():
    return jsonify({
        'status': 'ok',
        'message': 'ProjectWord Final Persian-Math DOCX Engine ✅',
        'endpoints': ['/health', '/generate']
    })

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'port': 8001})

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8001)
