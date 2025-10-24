from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import tempfile, zipfile, os, io, re, shutil

app = Flask(__name__)

# ==========================
# 🧩 کلاس پاکسازی متن فارسی
# ==========================
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        return text.strip()


# ==========================
# 📄 سازنده هوشمند ورد فارسی
# ==========================
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_doc()

    # == تنظیمات صفحه و ظاهر ==
    def _setup_doc(self):
        section = self.doc.sections[0]
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(1)

    # == تنظیم پاراگراف راست‌به‌چپ ==
    def _set_rtl_paragraph(self, p):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    # == تفسیر متن بولد با **
    def _parse_bold_text(self, text):
        parts, pattern, last_end = [], r'\*\*(.*?)\*\*', 0
        for m in re.finditer(pattern, text):
            if m.start() > last_end:
                parts.append({'text': text[last_end:m.start()], 'bold': False})
            parts.append({'text': m.group(1), 'bold': True})
            last_end = m.end()
        if last_end < len(text):
            parts.append({'text': text[last_end:], 'bold': False})
        return parts if parts else [{'text': text, 'bold': False}]

    # == افزودن متن پاراگراف ==
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

    # == افزودن جدول (LTR واقعی با متن راست‌چین) ==
    def add_table(self, lines):
        rows = []
        for ln in lines:
            if not ln.strip(): continue
            parts = [self.text_processor.clean_text(p.strip()) for p in ln.strip('|').split('|')]
            if len(parts) > 1:
                rows.append(parts)
        if not rows:
            return
        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]
        if len(rows) > 1 and all(set(x.strip()) <= {'-', ':', '|'} for x in rows[1]):
            rows.pop(1)

        table = self.doc.add_table(rows=0, cols=cols)
        table.style = 'Table Grid'

        for r_idx, row in enumerate(rows):
            tr = table.add_row().cells
            for c_idx in range(cols):
                cell = tr[c_idx]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                parts = self._parse_bold_text(row[c_idx])
                for run_obj in p.runs:
                    p._element.remove(run_obj._element)
                for part in parts:
                    run = p.add_run(part['text'])
                    run.font.name = 'B Nazanin'
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                    run.font.size = Pt(13)
                    if part['bold']:
                        run.bold = True
                # ردیف اول رنگ پس‌زمینه
                if r_idx == 0:
                    shading_elm = OxmlElement("w:shd")
                    shading_elm.set(qn('w:fill'), "D9E2F3")
                    cell._tc.get_or_add_tcPr().append(shading_elm)
        self.doc.add_paragraph()

    # == تفسیر متن کامل ==
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
            else:
                self.add_text(ln)
                i += 1

    # == اعمال اصلاحات XML نهایی برای جهت کلی سند ==
    def fix_global_rtl(self, stream):
        with tempfile.TemporaryDirectory() as tmpdir:
            # unzip محتویات DOCX
            with zipfile.ZipFile(stream, 'r') as z:
                z.extractall(tmpdir)
            xml_path = os.path.join(tmpdir, 'word', 'document.xml')
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(xml_path, parser)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            sectPr = root.find('.//w:sectPr', ns)
            if sectPr is not None:
                rtlGutter = sectPr.find('w:rtlGutter', ns)
                if rtlGutter is None:
                    rtlGutter = etree.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rtlGutter')
                rtlGutter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')

            tree.write(xml_path, encoding='utf-8', xml_declaration=True, standalone='yes')
            fixed_path = os.path.join(tmpdir, "data_fixed.docx")
            with zipfile.ZipFile(fixed_path, 'w', zipfile.ZIP_DEFLATED) as docx:
                for foldername, subfolders, filenames in os.walk(tmpdir):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        arcname = os.path.relpath(file_path, tmpdir)
                        docx.write(file_path, arcname)
            with open(fixed_path, 'rb') as f:
                return io.BytesIO(f.read())

    # == ذخیره نهایی ==
    def save_to_stream(self):
        tmp = io.BytesIO()
        self.doc.save(tmp)
        tmp.seek(0)
        fixed_stream = self.fix_global_rtl(tmp)
        fixed_stream.seek(0)
        return fixed_stream


# ==========================
# 🌐 Flask API Endpoint
# ==========================
@app.route('/generate', methods=['POST'])
def generate_docx():
    try:
        data = request.get_json(force=True)
        text = data.get('text', '')
        generator = SmartDocumentGenerator()
        generator.process_text(text)
        stream = generator.save_to_stream()
        return send_file(
            stream,
            as_attachment=True,
            download_name='persian_doc_final.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/')
def index():
    return jsonify({
        'message': 'نسخه نهایی تولیدکننده ورد فارسی با جهت درست صفحه و جدول ✔️'
    })


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8001)
