from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io

app = Flask(__name__)

# ---------------- فارسی‌ساز ----------------
class PersianTextProcessor:
    def clean_text(self, text):
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه').replace('ؤ', 'و')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        prefixes = ['می', 'نمی', 'بی', 'به', 'در', 'که']
        for p in prefixes:
            text = re.sub(f'\\b{p} ', f'{p}\u200c', text)
        return text.strip()

# ---------------- سازنده سند ----------------
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_doc()

    def _setup_doc(self):
        s = self.doc.sections[0]
        s.page_height = Inches(11.69)
        s.page_width = Inches(8.27)
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1)

    def _set_rtl(self, p):
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    # ---------------- شناسایی نوع محتوا ----------------
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

    # ---------------- تیتر ----------------
    def add_heading(self, text, level=1):
        text = re.sub(r'^#+\s*', '', text)
        text = self.text_processor.clean_text(text)

        # پشتیبانی از *...* یا **...**
        segments = re.split(r'(\*{1,2}[^*]+?\*{1,2})', text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(p)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        for seg in segments:
            if not seg.strip():
                continue
            if re.match(r'^\*{1,2}[^*]+?\*{1,2}$', seg):
                clean_seg = re.sub(r'^\*{1,2}|(?<=.)\*{1,2}$', '', seg)
                run = p.add_run(clean_seg)
                run.bold = True
            else:
                run = p.add_run(seg)
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            run.font.size = Pt(18 - level * 2)

    # ---------------- فرمول ----------------
    def add_formula(self, text):
        formulas = re.findall(r'\$\$.*?\$\$|\$.*?\$', text)
        for f in formulas:
            f = f.strip('$').strip()
            p = self.doc.add_paragraph(f)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.runs[0]
            r.font.name = 'Cambria Math'
            r.font.size = Pt(14)

    # ---------------- کپشن ----------------
    def add_caption(self, text):
        text = self.text_processor.clean_text(text)
        segments = re.split(r'(\*{1,2}[^*]+?\*{1,2})', text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)

        for seg in segments:
            if not seg.strip():
                continue
            if re.match(r'^\*{1,2}[^*]+?\*{1,2}$', seg):
                clean_seg = re.sub(r'^\*{1,2}|(?<=.)\*{1,2}$', '', seg)
                run = p.add_run(clean_seg)
                run.bold = True
            else:
                run = p.add_run(seg)
            run.font.name = 'B Nazanin'
            run.font.size = Pt(13)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- متن عادی ----------------
    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        self._set_rtl(p)

        # پشتیبانی از *...* یا **...**
        bold_segments = re.split(r'(\*{1,2}[^*]+?\*{1,2})', text)

        for segment in bold_segments:
            if not segment.strip():
                continue

            # اگر بین ستاره‌هاست => بولد
            if re.match(r'^\*{1,2}[^*]+?\*{1,2}$', segment):
                clean_seg = re.sub(r'^\*{1,2}|(?<=.)\*{1,2}$', '', segment)
                run = p.add_run(clean_seg)
                run.bold = True
                run.font.name = 'B Nazanin'
                run.font.size = Pt(14)
                run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                continue

            # سایر بخش‌ها مانند قبل
            parts = re.split(r'([A-Za-z0-9,;:.()\[\]{}=+\-*/^%<>])', segment)
            for part in parts:
                if not part.strip():
                    continue
                if re.match(r'[A-Za-z0-9]', part):
                    run = p.add_run(part)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)
                else:
                    run = p.add_run(part)
                    run.font.name = 'B Nazanin'
                    run.font.size = Pt(14)
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- پردازش کل متن ----------------
    def process_text(self, text):
        lines = [ln.rstrip() for ln in text.split('\n')]
        clean_lines = []
        for ln in lines:
            if ln == '' and (not clean_lines or clean_lines[-1] == ''):
                continue
            clean_lines.append(ln)

        previous_type = None
        for line in clean_lines:
            t = self.detect_content_type(line)
            if t == 'empty':
                if previous_type in ['text', 'formula']:
                    self.doc.add_paragraph().paragraph_format.space_after = Pt(0)
            elif t == 'heading':
                level = len(re.match(r'^#+', line).group())
                self.add_heading(line, level=min(level, 3))
            elif t == 'formula':
                self.add_formula(line)
            elif t in ['figure_caption', 'table_caption']:
                self.add_caption(line)
            else:
                self.add_text(line)
            previous_type = t

    def save_to_stream(self):
        buf = io.BytesIO()
        self.doc.save(buf)
        buf.seek(0)
        return buf

# ---------------- سرور Flask ----------------
@app.route('/generate', methods=['POST'])
def generate_word():
    try:
        data = request.get_json()
        text = data.get('text', '')
        if not text:
            return jsonify({'error': 'متن الزامی است'}), 400
        gen = SmartDocumentGenerator()
        gen.process_text(text)
        stream = gen.save_to_stream()
        return send_file(
            stream,
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
        'message': 'ProjectWord Final Persian DOCX Engine ✅',
        'endpoints': ['/health', '/generate']
    })

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'port': 8001})

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8001)
