from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io, re

app = Flask(__name__)

# ---- متن فارسی پاکسازی ----
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        return text.strip()

# ---- سازنده‌ی سند ----
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_doc()

    def _setup_doc(self):
        s = self.doc.sections[0]
        s.page_width, s.page_height = Inches(8.27), Inches(11.69)
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Inches(1)

    def _set_rtl_paragraph(self, p):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    def _parse_bold_text(self, text):
        parts = []
        pattern = r'\*\*(.*?)\*\*'
        last_end = 0
        for m in re.finditer(pattern, text):
            if m.start() > last_end:
                parts.append({'text': text[last_end:m.start()], 'bold': False})
            parts.append({'text': m.group(1), 'bold': True})
            last_end = m.end()
        if last_end < len(text):
            parts.append({'text': text[last_end:], 'bold': False})
        return parts if parts else [{'text': text, 'bold': False}]

    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        if not text.strip():
            return
        p = self.doc.add_paragraph()
        self._set_rtl_paragraph(p)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        parts = self._parse_bold_text(text)
        for part in parts:
            run = p.add_run(part['text'])
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            run.font.size = Pt(14)
            if part['bold']:
                run.bold = True

    def add_table(self, lines):
        rows = []
        for ln in lines:
            if not ln.strip():
                continue
            parts = [self.text_processor.clean_text(p.strip()) for p in ln.strip('|').split('|')]
            if len(parts) > 1:
                rows.append(parts)

        if not rows:
            return
        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]

        # حذف separator markdown
        if len(rows) > 1 and all(set(cell.strip()) <= {'-', ':', '|'} for cell in rows[1]):
            rows.pop(1)

        # -------- ایجاد جدول کاملاً RTL --------
        tbl = OxmlElement('w:tbl')
        tblPr = OxmlElement('w:tblPr')
        bidiVisual = OxmlElement('w:bidiVisual')
        bidiVisual.set(qn('w:val'), 'false')
        tblPr.insert(0, bidiVisual)
        tbl.append(tblPr)

        # ظاهر کلی جدول
        tblBorders = OxmlElement('w:tblBorders')
        for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '12')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)
        tblPr.append(tblBorders)

        # ساخت سطرها از راست به چپ
        for rindex, row in enumerate(rows):
            tr = OxmlElement('w:tr')
            for cindex in range(len(row)-1, -1, -1):  # از آخر به اول
                cell = OxmlElement('w:tc')
                tcPr = OxmlElement('w:tcPr')

                # سایه برای ردیف اول (هدر)
                if rindex == 0:
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9E2F3')
                    tcPr.append(shd)

                # حاشیه‌های سلول
                borders = OxmlElement('w:tcBorders')
                for b in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{b}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '6')
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), '000000')
                    borders.append(border)
                tcPr.append(borders)

                cell.append(tcPr)

                # پاراگراف داخل سلول
                p = OxmlElement('w:p')
                pPr = OxmlElement('w:pPr')
                align = OxmlElement('w:jc')
                align.set(qn('w:val'), 'center')
                pPr.append(align)
                bidi = OxmlElement('w:bidi')
                bidi.set(qn('w:val'), '1')
                pPr.append(bidi)
                p.append(pPr)

                # متن در xml آماده شود
                t_input = row[cindex]
                parts = self._parse_bold_text(t_input)
                for part in parts:
                    r = OxmlElement('w:r')
                    if part['bold']:
                        bold = OxmlElement('w:b')
                        bold.set(qn('w:val'), 'true')
                        rPr = OxmlElement('w:rPr')
                        rPr.append(bold)
                        fonts = OxmlElement('w:rFonts')
                        fonts.set(qn('w:cs'), 'B Nazanin')
                        fonts.set(qn('w:ascii'), 'B Nazanin')
                        rPr.append(fonts)
                        rPr.append(OxmlElement('w:lang'))
                        r.append(rPr)
                    t = OxmlElement('w:t')
                    t.text = part['text']
                    r.append(t)
                    p.append(r)
                cell.append(p)
                tr.append(cell)
            tbl.append(tr)

        # اضافه شدن جدول به بدنه
        self.doc._body._element.append(tbl)
        self.doc.add_paragraph()

    def process_text(self, text):
        lines = text.split('\n')
        i = 0
        while i < len(lines):
            ln = lines[i]
            if not ln.strip():
                i += 1
                continue
            if '|' in ln and len(ln.split('|')) > 2:
                block = []
                while i < len(lines) and '|' in lines[i]:
                    block.append(lines[i])
                    i += 1
                self.add_table(block)
                continue
            else:
                self.add_text(ln)
            i += 1

    def save_to_stream(self):
        buf = io.BytesIO()
        self.doc.save(buf)
        buf.seek(0)
        return buf

# ---- Flask route ----
@app.route('/generate', methods=['POST'])
def generate_docx():
    try:
        data = request.get_json(force=True)
        text = data.get('text', '')
        gen = SmartDocumentGenerator()
        gen.process_text(text)
        stream = gen.save_to_stream()
        return send_file(stream,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True,
                         download_name='persian_doc.docx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/')
def index():
    return jsonify({'message': 'Persian DOCX Generator — Full RTL compatibility ✅ for Word 2016'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
