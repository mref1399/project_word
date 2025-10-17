from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io
from hazm import Normalizer, word_tokenize
import sympy as sp
from sympy.parsing.latex import parse_latex

app = Flask(__name__)

class PersianTextProcessor:
    """پردازشگر هوشمند متن فارسی"""
    
    def __init__(self):
        self.normalizer = Normalizer(
            correct_spacing=True,
            remove_extra_spaces=True,
            remove_diacritics=False
        )
    
    def clean_text(self, text):
        """تمیزکاری هوشمند متن"""
        # نرمال‌سازی
        text = self.normalizer.normalize(text)
        
        # اصلاح فاصله‌های اشتباه
        text = re.sub(r'\s+([.,،؛:!؟»])', r'\1', text)  # حذف فاصله قبل نقطه
        text = re.sub(r'([«])\s+', r'\1', text)  # حذف فاصله بعد گیومه
        text = re.sub(r'\s+', ' ', text)  # فاصله‌های متوالی
        
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
        text = re.sub(r'\$\$.*?\$\$|\$.*?\$', replace_in_formula, text)
        
        return text
    
    def fix_half_spaces(self, text):
        """اصلاح نیم‌فاصله‌ها"""
        # بعد از می، نمی، بی و...
        prefixes = ['می', 'نمی', 'بی', 'با', 'از', 'به', 'در']
        for prefix in prefixes:
            text = re.sub(f'{prefix} ', f'{prefix}\u200c', text)
        
        return text

class MathProcessor:
    """پردازشگر فرمول‌های ریاضی"""
    
    @staticmethod
    def is_formula(text):
        """تشخیص فرمول ریاضی"""
        math_patterns = [
            r'\$\$.*?\$\$',  # Display math
            r'\$.*?\$',  # Inline math
            r'[∂∫∑∏√±×÷≈≠≤≥∞]',  # نمادهای ریاضی
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
        # حذف $ از اطراف
        formula = formula.strip('$').strip()
        
        # اصلاح فاصله‌های اضافی
        formula = re.sub(r'\s+', ' ', formula)
        
        # اصلاح الگوهای معمول
        formula = re.sub(r'\\left\s*', r'\\left', formula)
        formula = re.sub(r'\\right\s*', r'\\right', formula)
        
        return formula
    
    @staticmethod
    def format_formula_for_word(formula):
        """فرمت‌بندی فرمول برای Word"""
        # تبدیل LaTeX به Unicode (برای خوانایی بهتر)
        conversions = {
            r'\\alpha': 'α',
            r'\\beta': 'β',
            r'\\gamma': 'γ',
            r'\\delta': 'δ',
            r'\\epsilon': 'ε',
            r'\\theta': 'θ',
            r'\\lambda': 'λ',
            r'\\mu': 'μ',
            r'\\pi': 'π',
            r'\\rho': 'ρ',
            r'\\sigma': 'σ',
            r'\\tau': 'τ',
            r'\\phi': 'φ',
            r'\\omega': 'ω',
            r'\\partial': '∂',
            r'\\infty': '∞',
            r'\\nabla': '∇',
            r'\\times': '×',
            r'\\div': '÷',
            r'\\pm': '±',
            r'\\leq': '≤',
            r'\\geq': '≥',
            r'\\neq': '≠',
            r'\\approx': '≈',
            r'\\int': '∫',
            r'\\sum': '∑',
            r'\\prod': '∏',
            r'\\sqrt': '√',
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
        # تنظیم حاشیه‌ها
        sections = self.doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
        
        # تنظیم فونت پیش‌فرض
        style = self.doc.styles['Normal']
        style.font.name = 'B Nazanin'
        style.font.size = Pt(14)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    
    def detect_content_type(self, line):
        """تشخیص نوع محتوا"""
        line = line.strip()
        
        # خالی
        if not line:
            return 'empty'
        
        # عنوان
        if line.startswith('#'):
            return 'heading'
        
        # فرمول display
        if line.startswith('$$') or self.math_processor.is_formula(line):
            return 'formula'
        
        # لیست
        if re.match(r'^[\d\-\*]\s+', line):
            return 'list'
        
        # پاراگراف عادی
        return 'paragraph'
    
    def is_rtl(self, text):
        """تشخیص جهت متن"""
        persian_count = len(re.findall(r'[\u0600-\u06FF]', text))
        latin_count = len(re.findall(r'[A-Za-z]', text))
        return persian_count > latin_count
    
    def add_heading(self, text, level=1):
        """افزودن عنوان"""
        # حذف علامت #
        text = re.sub(r'^#+\s*', '', text)
        text = self.text_processor.clean_text(text)
        
        heading = self.doc.add_heading(text, level=level)
        heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT if self.is_rtl(text) else WD_ALIGN_PARAGRAPH.LEFT
        
        # استایل عنوان
        run = heading.runs[0]
        run.font.name = 'B Nazanin'
        run.font.size = Pt(18 - (level * 2))
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # تنظیم RTL
        if self.is_rtl(text):
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
    
    def add_formula(self, formula_text):
        """افزودن فرمول ریاضی"""
        # تمیزکاری فرمول
        formula = self.math_processor.clean_formula(formula_text)
        formula = self.math_processor.format_formula_for_word(formula)
        
        # افزودن به سند
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = paragraph.add_run(formula)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(14)
        
        # فاصله قبل و بعد
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)
    
    def add_mixed_text_paragraph(self, text):
        """افزودن پاراگراف با متن مخلوط (فارسی + انگلیسی + فرمول)"""
        text = self.text_processor.clean_text(text)
        
        paragraph = self.doc.add_paragraph()
        is_rtl = self.is_rtl(text)
        
        # تنظیم جهت و تراز
        if is_rtl:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.right_to_left = True
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # پردازش متن
        self._process_mixed_text(paragraph, text)
        
        # فاصله بین پاراگراف‌ها
        paragraph.paragraph_format.space_after = Pt(8)
    
    def _process_mixed_text(self, paragraph, text):
        """پردازش متن مخلوط"""
        buffer = ''
        current_script = None
        bold = False
        italic = False
        
        i = 0
        while i < len(text):
            char = text[i]
            
            # تشخیص bold
            if char == '*' and i + 1 < len(text) and text[i + 1] == '*':
                if buffer:
                    self._add_run(paragraph, buffer, current_script, bold, italic)
                    buffer = ''
                bold = not bold
                i += 2
                continue
            
            # تشخیص فرمول inline
            if char == '$':
                if buffer:
                    self._add_run(paragraph, buffer, current_script, bold, italic)
                    buffer = ''
                
                # یافتن پایان فرمول
                i += 1
                formula = ''
                while i < len(text) and text[i] != '$':
                    formula += text[i]
                    i += 1
                
                if i < len(text):
                    i += 1  # رد شدن از $ پایانی
                
                # افزودن فرمول
                formula = self.math_processor.format_formula_for_word(formula)
                run = paragraph.add_run(formula)
                run.font.name = 'Cambria Math'
                run.font.size = Pt(13)
                
                continue
            
            # تشخیص superscript/subscript
            if char in ('^', '_'):
                if buffer:
                    self._add_run(paragraph, buffer, current_script, bold, italic)
                    buffer = ''
                
                is_super = (char == '^')
                i += 1
                
                # خواندن مقدار
                value = ''
                if i < len(text) and text[i] == '{':
                    i += 1
                    while i < len(text) and text[i] != '}':
                        value += text[i]
                        i += 1
                    if i < len(text):
                        i += 1
                elif i < len(text):
                    value = text[i]
                    i += 1
                
                # افزودن run
                run = paragraph.add_run(value)
                run.font.size = Pt(10)
                if is_super:
                    run.font.superscript = True
                else:
                    run.font.subscript = True
                
                continue
            
            # کاراکترهای عادی
            script = self._detect_script(char)
            
            if script != current_script:
                if buffer:
                    self._add_run(paragraph, buffer, current_script, bold, italic)
                    buffer = ''
                current_script = script
            
            buffer += char
            i += 1
        
        # افزودن باقیمانده
        if buffer:
            self._add_run(paragraph, buffer, current_script, bold, italic)
    
    def _detect_script(self, char):
        """تشخیص نوع اسکریپت کاراکتر"""
        code = ord(char)
        
        # فارسی/عربی
        if (0x0600 <= code <= 0x06FF or 
            0x0750 <= code <= 0x077F or 
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
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8001)
