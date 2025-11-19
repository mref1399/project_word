# -*- coding: utf-8 -*-
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, html, re

app = Flask(__name__)

# ---------------- متن‌پاک‌کن ----------------
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = str(text)
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ؤ', 'و').replace('ە', 'ه')
        text = text.replace('\x00', '')                           # حذف Null
        text = re.sub(r'[\x01-\x08\x0b-\x1f\x7f]+', '', text)     # حذف کنترل‌ها
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'(\*\*|__)(.*?)\1', r'\2', text)           # حذف bold/underline نشانه
        return text.strip()

# ---------------- سازنده سند ----------------
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.cleaner = PersianTextProcessor()
        self._setup_doc()

    def _setup_doc(self):
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Inches(11.69), Inches(8.27)
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Inches(1)

    def _set_rtl(self, p):
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    # پارس علامت‌های bold/underline
    def _parse_inline_marks(self, text):
        parts = []
        pattern = r'(\*\*(.*?)\*\*|__(.*?)__)'
        last = 0
        for m in re.finditer(pattern, text):
            if m.start() > last:
                parts.append({'text': text[last:m.start()], 'bold': False, 'underline': False})
            inner = m.group(2) or m.group(3)
            if m.group(2):
                parts.append({'text': inner, 'bold': True, 'underline': False})
            else:
                parts.append({'text': inner, 'bold': False, 'underline': True})
            last = m.end()
        if last < len(text):
            parts.append({'text': text[last:], 'bold': False, 'underline': False})
        return parts if parts else [{'text': text, 'bold': False, 'underline': False}]

    def add_text(self, text):
        text = self.cleaner.clean_text(text)
        p = self.doc.add_paragraph()
        self._set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        for part in self._parse_inline_marks(text):
            safe_text = self.cleaner.clean_text(part['text'])
            run = p.add_run(safe_text)
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            run.font.size = Pt(14)
            run.bold = part['bold']
            run.underline = part['underline']

    def add_table(self, lines):
        rows = [list(filter(None, [x.strip() for x in ln.strip('|').split('|')])) for ln in lines if ln.strip()]
        if not rows:
            return
        table = self.doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
        table.style = 'Table Grid'
        for i, row in enumerate(rows):
            for j, cell_text in enumerate(row):
                cell = table.rows[i].cells[j]
                cell_text = self.cleaner.clean_text(cell_text)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._set_rtl(p)
                for part in self._parse_inline_marks(cell_text):
                    safe_text = self.cleaner.clean_text(part['text'])
                    run = p.add_run(safe_text)
                    run.font.name = 'B Nazanin'
                    run.font.size = Pt(12)
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                    if part['bold'] or i == 0:
                        run.bold = True
                    if part['underline']:
                        run.underline = True
        self.doc.add_paragraph()

    def process_text(self, text):
        lines = text.split('\n')
        block = []
        for ln in lines:
            if '|' in ln:
                block.append(ln)
            else:
                if block:
                    self.add_table(block)
                    block = []
                self.add_text(ln)
        if block:
            self.add_table(block)

    def save_to_stream(self):
        buf = io.BytesIO()
        self.doc.save(buf)
        buf.seek(0)
        return buf

# ---------------- Flask Endpoint ----------------
@app.route('/generate', methods=['POST'])
def generate_docx():
    try:
        data = request.get_json(force=True, silent=True)
        text = data.get('text', '') if data else ''
        gen = SmartDocumentGenerator()
        gen.process_text(text)
        stream = gen.save_to_stream()
        if not stream.getvalue():
            raise ValueError("خروجی خالی است.")
        return send_file(stream,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True,
                         download_name='document.docx')
    except Exception as e:
        safe_error = html.escape(str(e).encode('utf-8', 'ignore').decode('utf-8', 'ignore'))
        safe_error = re.sub(r'[\x01-\x08\x0b-\x1f\x7f]+', '', safe_error)  # آخرین خط فیلتر اشباح
        return jsonify({'error': f'Safe Fail ⛔ {safe_error}'}), 200

@app.route('/')
def home():
    return jsonify({'message': 'Generator running — no XML error guaranteed ✅'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
