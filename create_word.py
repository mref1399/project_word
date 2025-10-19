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

    # ---------------- تشخیص نوع محتوا ----------------
    def detect_content_type(self, line):
        line = line.strip()
        if not line:
            return 'empty'
        if re.match(r'^#+', line):
            return 'heading'
        if re.search(r'\$\$.*?\$\$|\$.*?\$', line):
            return 'formula'
        if '|' in line and len(line.split('|')) > 2:
            return 'table'
        if re.match(r'^شکل\s*\d+', line):
            return 'figure_caption'
        if re.match(r'^جدول\s*\d+', line):
            return 'table_caption'
        return 'text'

    # ---------------- تیتر ----------------
    def add_heading(self, text, level=1):
        text = re.sub(r'^#+\s*', '', text)
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(p)

        bold_segments = re.split(r'(\*{1,2}[^*]+?\*{1,2})', text)
        for seg in bold_segments:
            if not seg.strip():
                continue
            if re.match(r'^\*{1,2}[^*]+?\*{1,2}$', seg):
                seg = re.sub(r'^\*{1,2}|(?<=.)\*{1,2}$', '', seg)
                run = p.add_run(seg)
                run.bold = True
            else:
                run = p.add_run(seg)
            run.font.name = 'B Nazanin'
            run.font.size = Pt(18 - level * 2)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

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
        p = self.doc.add_paragraph(self.text_processor.clean_text(text))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        for run in p.runs:
            run.font.name = 'B Nazanin'
            run.font.size = Pt(13)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- جدول ----------------
    def add_table(self, lines):
        rows = []
        for line in lines:
            parts = [self.text_processor.clean_text(p.strip()) for p in line.strip('|').split('|')]
            if len(parts) > 1:
                rows.append(parts)

        if not rows:
            return

        table = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = 'Table Grid'
        table.autofit = False

        for i, row_data in enumerate(rows):
            row = table.rows[i]
            for j, cell_data in enumerate(row_data):
                cell = row.cells[j]
                p = cell.paragraphs[0]
                run = p.add_run(cell_data)
                if re.search(r'[A-Za-z0-9]', cell_data):
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(11)
                else:
                    run.font.name = 'B Nazanin'
                    run.font.size = Pt(12)
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._set_rtl(p)
                cell.width = Inches(1.5)

        for cell in table.rows[0].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True

    # ---------------- متن عادی ----------------
    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        self._set_rtl(p)

        bold_segments = re.split(r'(\*{1,2}[^*]+?\*{1,2})', text)
        for seg in bold_segments:
            if not seg.strip():
                continue
            if re.match(r'^\*{1,2}[^*]+?\*{1,2}$', seg):
                seg = re.sub(r'^\*{1,2}|(?<=.)\*{1,2}$', '', seg)
                run = p.add_run(seg)
                run.bold = True
                run.font.name = 'B Nazanin'
                run.font.size = Pt(14)
                run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                continue
            parts = re.split(r'([A-Za-z0-9,;:.()\[\]{}=+\-*/^%<>])', seg)
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
        i = 0
        while i < len(lines):
            line = lines[i]
            t = self.detect_content_type(line)

            if t == 'empty':
                i += 1
                continue

            if t == 'table':
                block = []
                while i < len(lines) and '|' in lines[i]:
                    block.append(lines[i])
                    i += 1
                self.add_table(block)
                continue

            elif t == 'heading':
                level = len(re.match(r'^#+', line).group())
                self.add_heading(line, level=min(level, 3))
            elif t == 'formula':
                self.add_formula(line)
            elif t in ['figure_caption', 'table_caption']:
                self.add_caption(line)
            else:
                self.add_text(line)
            i += 1

    def save_to_stream(self):
        buf = io.BytesIO()
        self.doc.save(buf)
        buf.seek(0)
        return buf

# ---------------- Flask ----------------
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
    return jsonify({'message': 'Persian DOCX Generator with Table ✅'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
