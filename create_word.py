from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from zipfile import ZipFile
import tempfile, os, io, shutil, re

app = Flask(__name__)

# 📚 پاک‌سازی دقیق متن فارسی
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        return text.strip()


# 🧩 سازنده هوشمند سند فارسی
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_page()

    def _setup_page(self):
        section = self.doc.sections[0]
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(1)

    def _set_rtl_para(self, p):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    def _parse_bold(self, text):
        pattern = r'\*\*(.*?)\*\*'
        parts = []
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
        if not text.strip(): return
        p = self.doc.add_paragraph()
        self._set_rtl_para(p)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        parts = self._parse_bold(text)
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
            cols = [self.text_processor.clean_text(x.strip()) for x in ln.strip('|').split('|')]
            if len(cols) > 1:
                rows.append(cols)
        if not rows:
            return
        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]

        table = self.doc.add_table(rows=0, cols=cols)
        table.style = 'Table Grid'

        for i, r in enumerate(rows):
            tr = table.add_row().cells
            for j in range(cols):
                cell = tr[j]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                for run_obj in p.runs:
                    p._element.remove(run_obj._element)

                parts = self._parse_bold(r[j])
                for part in parts:
                    run = p.add_run(part['text'])
                    run.font.name = 'B Nazanin'
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                    run.font.size = Pt(13)
                    if part['bold']:
                        run.bold = True

                if i == 0:  # رنگ پس‌زمینه‌ی سطر اول
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9E2F3')
                    cell._tc.get_or_add_tcPr().append(shd)

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
                blk = []
                while i < len(lines) and '|' in lines[i]:
                    blk.append(lines[i])
                    i += 1
                self.add_table(blk)
            else:
                self.add_text(ln)
                i += 1

    # ⚙️ تلفیق با منطق اصلی تو: اصلاح XML بعد از ساخت سند
    def _post_fix_xml(self, stream):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip = ZipFile(stream, 'r')
            zip.extractall(tmpdir)

            xml_path = os.path.join(tmpdir, 'word', 'document.xml')
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(xml_path, parser)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            # حذف تگ‌های RTL اشتباه از جداول
            for tbl in root.findall('.//w:tbl', ns):
                tblPr = tbl.find('w:tblPr', ns)
                if tblPr is not None:
                    for bad_tag in ['w:bidiVisual', 'w:tblDir']:
                        for el in tblPr.findall(bad_tag, ns):
                            tblPr.remove(el)

            # اطمینان از RTL کلی سند
            sectPr = root.find('.//w:sectPr', ns)
            if sectPr is not None:
                rtlGutter = sectPr.find('w:rtlGutter', ns)
                if rtlGutter is None:
                    rtlGutter = etree.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rtlGutter')
                rtlGutter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')

            tree.write(xml_path, pretty_print=True, encoding='utf-8', xml_declaration=True)

            fixed_path = os.path.join(tmpdir, 'data_fixed.docx')
            with ZipFile(fixed_path, 'w') as out_zip:
                for folder, _, files in os.walk(tmpdir):
                    for f in files:
                        path = os.path.join(folder, f)
                        arcname = os.path.relpath(path, tmpdir)
                        out_zip.write(path, arcname)

            with open(fixed_path, 'rb') as f:
                return io.BytesIO(f.read())

    def save_to_stream(self):
        stream = io.BytesIO()
        self.doc.save(stream)
        stream.seek(0)
        fixed_stream = self._post_fix_xml(stream)
        fixed_stream.seek(0)
        return fixed_stream


# 🧠 مسیر Flask
@app.route('/generate', methods=['POST'])
def generate_doc():
    data = request.get_json(force=True)
    text = data.get('text', '')
    gen = SmartDocumentGenerator()
    gen.process_text(text)
    stream = gen.save_to_stream()
    return send_file(
        stream,
        as_attachment=True,
        download_name='persian_doc_final.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.route('/')
def index():
    return jsonify({'msg': 'نسخه تلفیقی Persian DOCX Generator — جهت فارسیِ کامل و جدول استاندارد ✅'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
