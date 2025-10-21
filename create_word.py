from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, re

app = Flask(__name__)

# ---------------- فارسی‌ساز ----------------
class PersianTextProcessor:
    def clean_text(self, text):
        if not text:
            return ''
        text = text.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه').replace('ؤ', 'و')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
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

    def _set_cell_borders(self, cell):
        """تنظیم حاشیه‌های سلول برای ظاهر شکیل"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '12')  # ضخامت حاشیه
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')  # رنگ مشکی
            tcBorders.append(border)
        
        tcPr.append(tcBorders)

    def _set_cell_shading(self, cell, is_header=False):
        """رنگ پس‌زمینه برای سلول"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'D9E2F3' if is_header else 'FFFFFF')  # آبی روشن برای هدر
        tcPr.append(shading)

    def _set_cell_margins(self, cell):
        """فاصله داخلی سلول"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        
        for margin_name in ['top', 'left', 'bottom', 'right']:
            margin = OxmlElement(f'w:{margin_name}')
            margin.set(qn('w:w'), '100')
            margin.set(qn('w:type'), 'dxa')
            tcMar.append(margin)
        
        tcPr.append(tcMar)

    def _parse_bold_text(self, text):
        """تجزیه متن و شناسایی بخش‌های bold شده با **"""
        parts = []
        pattern = r'\*\*(.*?)\*\*'
        last_end = 0
        
        for match in re.finditer(pattern, text):
            # متن قبل از **
            if match.start() > last_end:
                parts.append({'text': text[last_end:match.start()], 'bold': False})
            # متن داخل **
            parts.append({'text': match.group(1), 'bold': True})
            last_end = match.end()
        
        # متن بعد از آخرین **
        if last_end < len(text):
            parts.append({'text': text[last_end:], 'bold': False})
        
        return parts if parts else [{'text': text, 'bold': False}]

    # ---------------- تشخیص نوع ----------------
    def detect_content_type(self, line):
        line = line.strip()
        if not line:
            return 'empty'
        if '|' in line and len(line.split('|')) > 2:
            return 'table'
        if re.match(r'^#+', line):
            return 'heading'
        if re.search(r'\$\$.*?\$\$|\$.*?\$', line):
            return 'formula'
        if re.match(r'^(شکل|جدول)\s*\d+', line):
            return 'caption'
        return 'text'

    # ---------------- تیتر ----------------
    def add_heading(self, text, level=1):
        text = re.sub(r'^#+\s*', '', text)
        # حذف ** از عناوین
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(p)
        run = p.add_run(text)
        run.bold = True
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
            r = p.runs[0]
            r.font.name = 'Cambria Math'
            r.font.size = Pt(14)

    # ---------------- کپشن ----------------
    def add_caption(self, text):
        # حذف ** از کپشن
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        p = self.doc.add_paragraph(self.text_processor.clean_text(text))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        for run in p.runs:
            run.bold = True
            run.font.name = 'B Nazanin'
            run.font.size = Pt(13)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- جدول شکیل با فرمت‌بندی حرفه‌ای ----------------
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

        # اطمینان از یکنواختی تعداد ستون‌ها
        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]

        # حذف خط جداکننده markdown (خط دوم با ---|---|---)
        if len(rows) > 1 and all(set(cell.strip()) <= {'-', ':', '|', ' '} for cell in rows[1]):
            rows.pop(1)

        if not rows:
            return

        try:
            table = self.doc.add_table(rows=len(rows), cols=cols)
            table.style = 'Table Grid'
            
            # تنظیم عرض جدول
            table.autofit = False
            table.allow_autofit = False
            
            for i, row_data in enumerate(rows):
                is_header = (i == 0)  # سطر اول را هدر در نظر می‌گیریم
                
                for j, cell_data in enumerate(row_data):
                    try:
                        cell = table.rows[i].cells[j]
                        
                        # تنظیمات ظاهری سلول
                        self._set_cell_borders(cell)
                        self._set_cell_shading(cell, is_header)
                        self._set_cell_margins(cell)
                        
                        # محتوای سلول با پشتیبانی از **bold**
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(3)
                        p.paragraph_format.space_after = Pt(3)
                        
                        # پردازش متن برای bold
                        parts = self._parse_bold_text(cell_data)
                        
                        for part in parts:
                            run = p.add_run(part['text'])
                            
                            # فونت و اندازه
                            if re.search(r'[A-Za-z0-9]', part['text']):
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(11)
                            else:
                                run.font.name = 'B Nazanin'
                                run.font.size = Pt(12)
                                run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                            
                            # اعمال bold
                            if part['bold'] or is_header:
                                run.bold = True
                            
                            # هدر را برجسته می‌کنیم
                            if is_header:
                                run.font.color.rgb = RGBColor(0, 0, 0)
                        
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        self._set_rtl(p)
                        
                    except Exception as e:
                        continue
            
            # اضافه کردن فاصله بعد از جدول
            self.doc.add_paragraph()
            
        except Exception:
            # اگر جدول واقعاً خراب است، کل داده را به متن عادی تبدیل می‌کند
            joined = "\n".join([" | ".join(r) for r in rows])
            self.add_text(joined)
            return

    # ---------------- متن با پشتیبانی از bold ----------------
    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        if not text:
            return
        
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        self._set_rtl(p)
        
        # پردازش متن برای bold
        parts = self._parse_bold_text(text)
        
        for part in parts:
            run = p.add_run(part['text'])
            run.font.name = 'B Nazanin'
            run.font.size = Pt(14)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            
            if part['bold']:
                run.bold = True

    # ---------------- پردازش کل ----------------
    def process_text(self, text):
        if not text or not isinstance(text, str):
            self.add_text("⚠️ ورودی خالی یا نامعتبر بود.")
            return

        lines = text.split('\n')
        i = 0
        while i < len(lines):
            ln = lines[i]
            t = self.detect_content_type(ln)

            if t == 'empty':
                i += 1
                continue
            elif t == 'table':
                block = []
                while i < len(lines) and '|' in lines[i]:
                    block.append(lines[i])
                    i += 1
                self.add_table(block)
                continue
            elif t == 'heading':
                level = len(re.match(r'^#+', ln).group())
                self.add_heading(ln, level=min(level, 3))
            elif t == 'formula':
                self.add_formula(ln)
            elif t == 'caption':
                self.add_caption(ln)
            else:
                self.add_text(ln)
            i += 1

    def save_to_stream(self):
        buffer = io.BytesIO()
        self.doc.save(buffer)
        buffer.seek(0)
        return buffer

# ---------------- Flask route ----------------
@app.route('/generate', methods=['POST'])
def generate_word():
    try:
        data = request.get_json(force=True, silent=True)
        if not data or 'text' not in data:
            return jsonify({'error': 'متن الزامی است'}), 400
        text = data.get('text', '')
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
        # هرگونه استثناء را مهار می‌کند تا n8n هرگز 500 نگیرد
        return jsonify({'error': f'Safe Fail ⛔ {str(e)}'}), 200

@app.route('/')
def home():
    return jsonify({'message': 'Persian DOCX Generator — Safe Mode ✅'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
