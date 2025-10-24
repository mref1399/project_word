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
        text = re.sub(r'\s+([.,،؛:!؟»\\)])', r'\1', text)
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
            border.set(qn('w:sz'), '12')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tcBorders.append(border)
        
        tcPr.append(tcBorders)

    def _set_cell_shading(self, cell, is_header=False):
        """رنگ پس‌زمینه برای سلول"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'D9E2F3' if is_header else 'FFFFFF')
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

    def _clean_markup(self, text):
        """حذف تمام نشانه‌گذاری‌های فرمت (**, __, ~~)"""
        # حذف ** برای bold
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        # حذف __ برای italic/underline
        text = re.sub(r'__(.*?)__', r'\1', text)
        # حذف ~~ برای strikethrough
        text = re.sub(r'~~(.*?)~~', r'\1', text)
        return text

    def _parse_formatted_text(self, text):
        """تجزیه متن و شناسایی فرمت‌های مختلف (**, __, ~~)"""
        parts = []
        # الگوی ترکیبی برای ** و __ و ~~
        pattern = r'(\*\*.*?\*\*|__.*?__|~~.*?~~)'
        segments = re.split(pattern, text)
        
        for segment in segments:
            if not segment:
                continue
            
            if segment.startswith('**') and segment.endswith('**'):
                # متن bold
                parts.append({'text': segment[2:-2], 'bold': True, 'italic': False, 'strike': False})
            elif segment.startswith('__') and segment.endswith('__'):
                # متن italic/underline
                parts.append({'text': segment[2:-2], 'bold': False, 'italic': True, 'strike': False})
            elif segment.startswith('~~') and segment.endswith('~~'):
                # متن strikethrough
                parts.append({'text': segment[2:-2], 'bold': False, 'italic': False, 'strike': True})
            else:
                # متن عادی
                parts.append({'text': segment, 'bold': False, 'italic': False, 'strike': False})
        
        return parts if parts else [{'text': text, 'bold': False, 'italic': False, 'strike': False}]

    def _parse_chemical_formula(self, text):
        """تجزیه فرمول شیمیایی و شناسایی زیرنویس و بالانویس"""
        parts = []
        # الگوی ترکیبی: H_2O یا H_{10}O یا X^2 یا X^{10}
        pattern = r'([A-Za-z]+)(_\{?\d+\}?|\^\{?\d+\}?)?'
        
        pos = 0
        for match in re.finditer(pattern, text):
            # متن قبل از فرمول
            if match.start() > pos:
                parts.append({
                    'text': text[pos:match.start()],
                    'subscript': False,
                    'superscript': False
                })
            
            element = match.group(1)  # عنصر شیمیایی (مثلاً H)
            modifier = match.group(2)  # زیرنویس یا بالانویس (_2 یا ^2)
            
            # اضافه کردن عنصر
            parts.append({
                'text': element,
                'subscript': False,
                'superscript': False
            })
            
            # اضافه کردن زیرنویس یا بالانویس
            if modifier:
                number = re.sub(r'[_^\{\}]', '', modifier)  # حذف _, ^, {, }
                is_subscript = modifier.startswith('_')
                is_superscript = modifier.startswith('^')
                
                parts.append({
                    'text': number,
                    'subscript': is_subscript,
                    'superscript': is_superscript
                })
            
            pos = match.end()
        
        # متن باقیمانده
        if pos < len(text):
            parts.append({
                'text': text[pos:],
                'subscript': False,
                'superscript': False
            })
        
        return parts if parts else [{'text': text, 'subscript': False, 'superscript': False}]

    # ---------------- تشخیص نوع ----------------
    def detect_content_type(self, line):
        line = line.strip()
        if not line:
            return 'empty'
        if '|' in line and len(line.split('|')) > 2:
            return 'table'
        if re.match(r'^#+', line):
            return 'heading'
        # تشخیص فرمول ریاضی با $ و $$
        if re.search(r'\$\$.*?\$\$|\$.*?\$', line):
            return 'formula'
        # تشخیص فرمول شیمیایی (H_2O, CO_2, etc.)
        if re.search(r'[A-Za-z]+_\{?\d+\}?|\^\{?\d+\}?', line):
            return 'chemical'
        if re.match(r'^(شکل|جدول)\s*\d+', line):
            return 'caption'
        return 'text'

    # ---------------- تیتر ----------------
    def add_heading(self, text, level=1):
        text = re.sub(r'^#+\s*', '', text)
        # حذف تمام نشانه‌گذاری‌ها از عناوین
        text = self._clean_markup(text)
        text = self.text_processor.clean_text(text)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(p)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'B Nazanin'
        run.font.size = Pt(18 - level * 2)
        run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- فرمول ریاضی ----------------
    def add_formula(self, text):
        """نمایش فرمول‌های ریاضی با $ و $$"""
        # جدا کردن فرمول‌ها از متن عادی
        parts = re.split(r'(\$\$.*?\$\$|\$.*?\$)', text)
        
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        for part in parts:
            if not part:
                continue
            
            if part.startswith('$$') and part.endswith('$$'):
                # فرمول display mode
                formula_text = part.strip('$').strip()
                run = p.add_run(formula_text)
                run.font.name = 'Cambria Math'
                run.font.size = Pt(14)
            elif part.startswith('$') and part.endswith('$'):
                # فرمول inline
                formula_text = part.strip('$').strip()
                run = p.add_run(formula_text)
                run.font.name = 'Cambria Math'
                run.font.size = Pt(12)
            else:
                # متن عادی
                run = p.add_run(part)
                run.font.name = 'B Nazanin'
                run.font.size = Pt(14)
                run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- فرمول شیمیایی ----------------
    def add_chemical_formula(self, text):
        """نمایش فرمول‌های شیمیایی با زیرنویس و بالانویس صحیح"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        parts = self._parse_chemical_formula(text)
        
        for part in parts:
            run = p.add_run(part['text'])
            run.font.name = 'Cambria'
            run.font.size = Pt(14)
            
            if part['subscript']:
                run.font.subscript = True
            elif part['superscript']:
                run.font.superscript = True

    # ---------------- کپشن ----------------
    def add_caption(self, text):
        # حذف تمام نشانه‌گذاری‌ها از کپشن
        text = self._clean_markup(text)
        p = self.doc.add_paragraph(self.text_processor.clean_text(text))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        for run in p.runs:
            run.bold = True
            run.font.name = 'B Nazanin'
            run.font.size = Pt(13)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')

    # ---------------- جدول با حذف نشانه‌گذاری ----------------
    def add_table(self, lines):
        rows = []
        for ln in lines:
            if not ln.strip():
                continue
            # حذف ** و __ از محتوای سلول‌ها
            parts = [self._clean_markup(self.text_processor.clean_text(p.strip())) for p in ln.strip('|').split('|')]
            if len(parts) > 1:
                rows.append(parts)

        if not rows:
            return

        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]

        # حذف خط جداکننده markdown
        if len(rows) > 1 and all(set(cell.strip()) <= {'-', ':', '|', ' '} for cell in rows[1]):
            rows.pop(1)

        if not rows:
            return

        try:
            table = self.doc.add_table(rows=len(rows), cols=cols)
            table.style = 'Table Grid'
            table.autofit = False
            table.allow_autofit = False
            
            for i, row_data in enumerate(rows):
                is_header = (i == 0)
                
                for j, cell_data in enumerate(row_data):
                    try:
                        cell = table.rows[i].cells[j]
                        
                        self._set_cell_borders(cell)
                        self._set_cell_shading(cell, is_header)
                        self._set_cell_margins(cell)
                        
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(3)
                        p.paragraph_format.space_after = Pt(3)
                        
                        # اضافه کردن متن ساده (بدون فرمت)
                        run = p.add_run(cell_data)
                        
                        if re.search(r'[A-Za-z0-9]', cell_data):
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(11)
                        else:
                            run.font.name = 'B Nazanin'
                            run.font.size = Pt(12)
                            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
                        
                        if is_header:
                            run.bold = True
                            run.font.color.rgb = RGBColor(0, 0, 0)
                        
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        self._set_rtl(p)
                        
                    except Exception:
                        continue
            
            self.doc.add_paragraph()
            
        except Exception:
            joined = "\n".join([" | ".join(r) for r in rows])
            self.add_text(joined)

    # ---------------- متن با پشتیبانی از فرمت‌ها ----------------
    def add_text(self, text):
        text = self.text_processor.clean_text(text)
        if not text:
            return
        
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        self._set_rtl(p)
        
        # پردازش متن برای فرمت‌ها
        parts = self._parse_formatted_text(text)
        
        for part in parts:
            run = p.add_run(part['text'])
            run.font.name = 'B Nazanin'
            run.font.size = Pt(14)
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            
            if part['bold']:
                run.bold = True
            if part['italic']:
                run.italic = True
            if part['strike']:
                run.font.strike = True

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
            elif t == 'chemical':
                self.add_chemical_formula(ln)
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
        return jsonify({'error': f'Safe Fail ⛔ {str(e)}'}), 200

@app.route('/')
def home():
    return jsonify({'message': 'Persian DOCX Generator — Safe Mode ✅'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
