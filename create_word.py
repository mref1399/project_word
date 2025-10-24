# -*- coding: utf-8 -*-
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

# ========================================
# 📚 پاک‌سازی دقیق متن فارسی
# ========================================
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        return text.strip()


# ========================================
# 🧩 سازنده هوشمند سند فارسی
# ========================================
class SmartDocumentGenerator:
    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self._setup_page()

    def _setup_page(self):
        section = self.doc.sections[0]
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(1)

    # 📌 تنظیم پاراگراف راست‌به‌چپ
    def _set_rtl_para(self, p):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

    # پردازش پررنگ‌ها (Bold)
    def _parse_bold(self, text):
        pattern = r'\*\*(.*?)\*\*'
        parts, last_end = [], 0
        for m in re.finditer(pattern, text):
            if m.start() > last_end:
                parts.append({'text': text[last_end:m.start()], 'bold': False})
            parts.append({'text': m.group(1), 'bold': True})
            last_end = m.end()
        if last_end < len(text):
            parts.append({'text': text[last_end:], 'bold': False})
        return parts if parts else [{'text': text, 'bold': False}]

    # ✍️ اضافه کردن پاراگراف
    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        if not text.strip():
            return
        p = self.doc.add_paragraph()
        self._set_rtl_para(p)
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        parts = self._parse_bold(text)
        for part in parts:
            run = p.add_run(part['text'])
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            run.font.size = Pt(14)
            if part['bold']:
                run.bold = True

    # ✏️ اضافه کردن جدول
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

        for i, row in enumerate(rows):
            tr = table.add_row().cells
            for j, cell_text in enumerate(row):
                cell = tr[j]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    p._element.remove(r._element)

                parts = self._parse_bold(cell_text)
                for part in parts:
                    run = p.add_run(part['text'])
                    run.font.name = 'B Nazanin'
                    run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                    run.font.size = Pt(13)
                    if part['bold']:
                        run.bold = True

                if i == 0:  # رنگ ردیف اول
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9E2F3')
                    cell._tc.get_or_add_tcPr().append(shd)

        self.doc.add_paragraph()

    # تحلیل خودکار روی متن ورودی
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

    # 🧩 رفع کامل مشکلات راست‌به‌چپ در XML
    def _post_fix_xml(self, input_path, output_path):
        temp_dir = tempfile.mkdtemp()
        with ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        doc_xml = os.path.join(temp_dir, 'word/document.xml')
        parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
        tree = etree.parse(doc_xml, parser)
        root = tree.getroot()

        # حذف تگ‌های جهت‌دهی اشتباه
        for bad_tag in ['w:bidiVisual', 'w:tblDir']:
            for el in root.findall(f".//{bad_tag}", ns):
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)

        # اضافه‌کردن جهت RTL برای تمام پاراگراف‌ها
        for para in root.findall('.//w:p', ns):
            pPr = para.find('w:pPr', ns)
            if pPr is None:
                pPr = etree.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
            bidi = pPr.find('w:bidi', ns)
            if bidi is None:
                bidi = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi')
                bidi.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1')

        # جهت‌دهی جدول‌ها
        for tbl in root.findall('.//w:tbl', ns):
            tblPr = tbl.find('w:tblPr', ns)
            if tblPr is None:
                tblPr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr')
            rtl = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi')
            rtl.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1')

        # راست‌چین کردن کل سند
        sectPr = root.find('.//w:sectPr', ns)
        if sectPr is not None:
            rtlGutter = sectPr.find('w:rtlGutter', ns)
            if rtlGutter is None:
                rtlGutter = etree.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rtlGutter')
            rtlGutter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')

        tree.write(doc_xml, encoding='utf-8', xml_declaration=True, standalone=True)

        # 🎨 تنظیم فونت پیش‌فرض در styles.xml
        styles_xml = os.path.join(temp_dir, 'word/styles.xml')
        if os.path.exists(styles_xml):
            stree = etree.parse(styles_xml, parser)
            sroot = stree.getroot()
            fonts = sroot.findall('.//w:rFonts', ns)
            for f in fonts:
                f.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', 'B Nazanin')
                f.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', 'B Nazanin')
                f.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs', 'B Nazanin')
            stree.write(styles_xml, encoding='utf-8', xml_declaration=True, standalone=True)

        # ذخیره DOCX جدید
        with ZipFile(output_path, 'w', compression=ZipFile.ZIP_DEFLATED) as zip_out:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_out.write(file_path, arcname)

        shutil.rmtree(temp_dir)
        return output_path

    # خروجی نهایی برای پاسخ HTTP
    def save_to_stream(self):
        tmp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        self.doc.save(tmp_input.name)

        tmp_output = tempfile.NamedTemporaryFile(delete=False, suffix='_fixed.docx')
        fixed = self._post_fix_xml(tmp_input.name, tmp_output.name)

        with open(fixed, 'rb') as f:
            data = f.read()

        for path in [tmp_input.name, tmp_output.name, fixed]:
            if os.path.exists(path):
                os.remove(path)

        stream = io.BytesIO(data)
        stream.seek(0)
        return stream


# ========================================
# 🌐 مسیرهای Flask
# ========================================
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
    return jsonify({'msg': '📄 Persian DOCX Generator — نسخه نهایی، کاملاً سازگار با Word ✅'})


# ========================================
# 🚀 اجرای سرور Flask
# ========================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
