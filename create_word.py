from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io
import sympy as sp

app = Flask(__name__)

class PersianTextProcessor:
    """پردازشگر متن فارسی (بدون hazm)"""

    def __init__(self):
        pass

    def clean_text(self, text):
        """تمیزکاری هوشمند متن"""
        # نرمال‌سازی حروف عربی
        text = text.replace('ي', 'ی').replace('ك', 'ک')
        text = text.replace('ە', 'ه').replace('ؤ', 'و')
        
        # حذف فاصله‌های اضافی
        text = re.sub(r'\s+', ' ', text)
        
        # اصلاح فاصله‌های اشتباه
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)  # حذف فاصله قبل نقطه
        text = re.sub(r'([(«])\s+', r'\1', text)  # حذف فاصله بعد پرانتز
        
        # اصلاح اعداد فارسی در فرمول‌ها
        text = self.fix_numbers_in_formulas(text)
        
        # اصلاح نیم‌فاصله
        text = self.fix_half_spaces(text)

        return text.strip()

    def fix_numbers_in_formulas(self, text):
        """تبدیل اعداد فارسی به انگلیسی در فرمول‌ها"""
        persian_digits = '۰۱۲۳۴۵۶۷۸۹'
        english_digits = '0123456789'
        trans_table = str.maketrans(persian_digits, english_digits)

        def replace_in_formula(match):
            formula = match.group(0)
            return formula.translate(trans_table)

        # در فرمول‌ها اعداد را تبدیل کن
        text = re.sub(r'\$\$.*?\$\$|\$.*?\$', replace_in_formula, text, flags=re.DOTALL)
        return text

    def fix_half_spaces(self, text):
        """اصلاح نیم‌فاصله‌ها"""
        # بعد از می، نمی، بی و...
        prefixes = ['می', 'نمی', 'بی', 'با', 'از', 'به', 'در', 'که']
        for prefix in prefixes:
            text = re.sub(f'\\b{prefix} ', f'{prefix}\u200c', text)
        return text

class MathProcessor:
    """پردازشگر فرمول‌های ریاضی"""

    @staticmethod
    def is_formula(text):
        """تشخیص فرمول ریاضی"""
        math_patterns = [
            r'\$\$.*?\$\$',  # Display math
            r'\$.*?\$',  # Inline math
            r'[∂∫∑∏√±×÷≈≠≤≥∞αβγδεθλμπρσφω]',  # نمادهای ریاضی
            r'\\[a-zA-Z]+',  # دستورات LaTeX
            r'\^[\{\d]',  # توان
            r'_[\{\d]',  # زیرنویس
        ]

        for pattern in math_patterns:
            if re.search(pattern, text):
                return True
        return False

    @staticmethod
    def clean_formula(formula):
        """تمیزکاری فرمول"""
        formula = formula.strip('$').strip()
        formula = re.sub(r'\s+', ' ', formula)
        formula = re.sub(r'\\left\s*', r'\\left', formula)
        formula = re.sub(r'\\right\s*', r'\\right', formula)
        return formula

    @staticmethod
    def format_formula_for_word(formula):
        """فرمت‌بندی فرمول برای Word"""
        conversions = {
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ',
            r'\\delta': 'δ', r'\\epsilon': 'ε', r'\\theta': 'θ',
            r'\\lambda': 'λ', r'\\mu': 'μ', r'\\pi': 'π',
            r'\\rho': 'ρ', r'\\sigma': 'σ', r'\\tau': 'τ',
            r'\\phi': 'φ', r'\\omega': 'ω', r'\\Omega': 'Ω',
            r'\\partial': '∂', r'\\infty': '∞', r'\\nabla': '∇',
            r'\\times': '×', r'\\div': '÷', r'\\pm': '±',
            r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠',
            r'\\approx': '≈', r'\\int': '∫', r'\\sum': '∑',
            r'\\prod': '∏', r'\\sqrt': '√',
        }

        for latex, unicode_char in conversions.items():
            formula = formula.replace(latex, unicode_char)

        return formula

class SmartDocumentGenerator:
    """تولیدکننده هوشمند سند Word"""

    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self.math_processor = MathProcessor()
        self._setup_document()

    def _setup_document(self):
        """تنظیمات اولیه سند"""
        section = self.doc.sections[0]
        section.page_height = Inches(11.69)  # A4
        section.page_width = Inches(8.27)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    def detect_content_type(self, line):
        """تشخیص نوع محتوا"""
        line = line.strip()
        
        if not line:
            return 'empty'
        
        if re.match(r'^#+\s', line):
            return 'heading'
        
        if self.math_processor.is_formula(line):
            return 'formula'
        
        if re.match(r'^\*\*(.+?)\*\*', line) or re.match(r'^\*(.+?)\*', line):
            return 'bold'
        
        return 'text'

    def add_heading(self, text, level=1):
        """افزودن عنوان"""
        text = re.sub(r'^#+\s*', '', text)
        text = self.text_processor.clean_text(text)
        
        heading = self.doc.add_heading(level=level)
        run = heading.add_run(text)
        run.font.name = 'B Nazanin'
        run.font.size = Pt(18 - level * 2)
        run.bold = True
        run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
        
        heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def add_formula(self, text):
        """افزودن فرمول"""
        formula_match = re.search(r'\$\$(.*?)\$\$|\$(.*?)\$', text, re.DOTALL)
        
        if formula_match:
            formula = formula_match.group(1) or formula_match.group(2)
            formula = self.math_processor.clean_formula(formula)
            formula = self.math_processor.format_formula_for_word(formula)
            
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run = paragraph.add_run(formula)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(14)

    def add_mixed_text_paragraph(self, text):
        """افزودن پاراگراف با متن مختلط"""
        text = self.text_processor.clean_text(text)
        
        # اصلاح بولد و ایتالیک
        text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
        text = re.sub(r'\*(.+?)\*', r'<i>\1</i>', text)
        
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
        # پردازش قطعات
        parts = re.split(r'(<b>.*?</b>|<i>.*?</i>)', text)
        
        for part in parts:
            if not part:
                continue
            
            bold = False
            italic = False
            
            if part.startswith('<b>'):
                bold = True
                part = re.sub(r'</?b>', '', part)
            elif part.startswith('<i>'):
                italic = True
                part = re.sub(r'</?i>', '', part)
            
            # تشخیص اسکریپت
            current_script = None
            current_text = ''
            
            for char in part:
                script = self._detect_script(char)
                
                if script != current_script:
                    if current_text:
                        self._add_run(paragraph, current_text, current_script, bold, italic)
                    current_script = script
                    current_text = char
                else:
                    current_text += char
            
            if current_text:
                self._add_run(paragraph, current_text, current_script, bold, italic)

    def _detect_script(self, char):
        """تشخیص نوع اسکریپت کاراکتر"""
        code = ord(char)
        
        # فارسی/عربی
        if (0x0600 <= code <= 0x06FF or
            0xFB50 <= code <= 0xFDFF or
            0xFE70 <= code <= 0xFEFF):
            return 'persian'
        
        # لاتین
        if (0x0020 <= code <= 0x007E or
            0x00A0 <= code <= 0x00FF):
            return 'latin'

        return 'other'

    def _add_run(self, paragraph, text, script, bold, italic):
        """افزودن run با فرمت مناسب"""
        run = paragraph.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(14)

        if script == 'persian':
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
        else:
            run.font.name = 'Times New Roman'

        return run

    def process_text(self, text):
        """پردازش کامل متن"""
        lines = text.split('\n')

        for line in lines:
            content_type = self.detect_content_type(line)

            if content_type == 'empty':
                self.doc.add_paragraph()

            elif content_type == 'heading':
                level = len(re.match(r'^#+', line).group())
                self.add_heading(line, level=min(level, 3))

            elif content_type == 'formula':
                self.add_formula(line)

            else:
                self.add_mixed_text_paragraph(line)

    def save_to_stream(self):
        """ذخیره در stream"""
        file_stream = io.BytesIO()
        self.doc.save(file_stream)
        file_stream.seek(0)
        return file_stream

@app.route('/generate', methods=['POST'])
def generate_document():
    """API تولید سند"""
    try:
        data = request.get_json()
        text = data.get('text', '')

        if not text:
            return jsonify({'error': 'متن الزامی است'}), 400

        # تولید سند
        generator = SmartDocumentGenerator()
        generator.process_text(text)
        file_stream = generator.save_to_stream()

        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='document.docx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok', 'message': 'Service is running on port 8001'})

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8001)
