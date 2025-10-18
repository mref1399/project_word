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

# ===============================
#   Persian Text Processor
# ===============================
class PersianTextProcessor:
    """پردازشگر متن فارسی (بدون hazm)"""

    def __init__(self):
        pass

    def clean_text(self, text):
        # نرمال‌سازی
        text = text.replace('ی', 'ی').replace('ک', 'ک')
        text = text.replace('ه', 'ه').replace('ؤ', 'و')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])', r'\1', text)
        text = re.sub(r'([(«])\s+', r'\1', text)
        text = self.fix_numbers_in_formulas(text)
        text = self.fix_half_spaces(text)
        return text.strip()

    def fix_numbers_in_formulas(self, text):
        persian_digits = '۰۱۲۳۴۵۶۷۸۹'
        english_digits = '0123456789'
        trans_table = str.maketrans(persian_digits, english_digits)

        def replace_in_formula(match):
            formula = match.group(0)
            return formula.translate(trans_table)

        return re.sub(r'\$\$.*?\$\$|\$.*?\$', replace_in_formula, text, flags=re.DOTALL)

    def fix_half_spaces(self, text):
        prefixes = ['می', 'نمی', 'بی', 'با', 'از', 'به', 'در', 'که']
        for prefix in prefixes:
            text = re.sub(f'\\b{prefix} ', f'{prefix}\u200c', text)
        return text


# ===============================
#   Math Processor
# ===============================
class MathProcessor:
    """پردازشگر فرمول‌های ریاضی"""

    @staticmethod
    def is_formula(text):
        math_patterns = [
            r'\$\$.*?\$\$',
            r'\$.*?\$',
            r'[∂∫∑∏√±×÷≈≠≤≥∞αβγδεθλμπρσφω]',
            r'\\[a-zA-Z]+',
            r'\^[\{\d]',
            r'_[\{\d]',
        ]
        return any(re.search(p, text) for p in math_patterns)

    @staticmethod
    def clean_formula(formula):
        formula = formula.strip('$').strip()
        formula = re.sub(r'\s+', ' ', formula)
        formula = re.sub(r'\\left\s*', r'\\left', formula)
        formula = re.sub(r'\\right\s*', r'\\right', formula)
        return formula

    @staticmethod
    def format_formula_for_word(formula):
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
        for latex, uni in conversions.items():
            formula = formula.replace(latex, uni)
        return formula


# ===============================
#   Smart Document Generator
# ===============================
class SmartDocumentGenerator:
    """تولیدکننده هوشمند سند Word"""

    def __init__(self):
        self.doc = Document()
        self.text_processor = PersianTextProcessor()
        self.math_processor = MathProcessor()
        self._setup_document()

    def _setup_document(self):
        section = self.doc.sections[0]
        section.page_height = Inches(11.69)
        section.page_width = Inches(8.27)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    def detect_content_type(self, line):
        line = line.strip()
        if not line:
            return 'empty'
        if re.match(r'^#+\s', line):
            return 'heading'
        if self.math_processor.is_formula(line):
            return 'formula'
        return 'text'

    def add_heading(self, text, level=1):
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
        text = self.text_processor.clean_text(text)
        text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
        text = re.sub(r'\*(.+?)\*', r'<i>\1</i>', text)
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        parts = re.split(r'(<b>.*?</b>|<i>.*?</i>)', text)
        for part in parts:
            if not part:
                continue
            bold = '<b>' in part
            italic = '<i>' in part
            part = re.sub(r'</?(b|i)>', '', part)
            script, current_text = None, ''
            for char in part:
                current_script = self._detect_script(char)
                if current_script != script:
                    if current_text:
                        self._add_run(paragraph, current_text, script, bold, italic)
                    script, current_text = current_script, char
                else:
                    current_text += char
            if current_text:
                self._add_run(paragraph, current_text, script, bold, italic)

    def _detect_script(self, char):
        code = ord(char)
        if (0x0600 <= code <= 0x06FF or 0xFB50 <= code <= 0xFEFF):
            return 'persian'
        if (0x0020 <= code <= 0x00FF):
            return 'latin'
        return 'other'

    def _add_run(self, paragraph, text, script, bold, italic):
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
        lines = text.split('\n')
        for line in lines:
            ctype = self.detect_content_type(line)
            if ctype == 'empty':
                self.doc.add_paragraph()
            elif ctype == 'heading':
                level = len(re.match(r'^#+', line).group())
                self.add_heading(line, level=min(level, 3))
            elif ctype == 'formula':
                self.add_formula(line)
            else:
                self.add_mixed_text_paragraph(line)

    def save_to_stream(self):
        file_stream = io.BytesIO()
        self.doc.save(file_stream)
        file_stream.seek(0)
        return file_stream


# ===============================
#   Routes
# ===============================
@app.route('/')
def home():
    """Root endpoint for browser access"""
    return jsonify({
        'status': 'ok',
        'message': 'ProjectWord is running on Darkube 🚀',
        'endpoints': ['/health', '/generate']
    })


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': 'Service is running on port 8001'})


@app.route('/generate', methods=['POST'])
def generate_document():
    try:
        data = request.get_json()
        text = data.get('text', '')
        if not text:
            return jsonify({'error': 'متن الزامی است'}), 400
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


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8001)
