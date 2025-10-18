from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io

app = Flask(__name__)

# ---------------- Persian Text Processor ----------------
class PersianTextProcessor:
    def clean_text(self, text):
        text = text.replace('ي','ی').replace('ك','ک').replace('ە','ه').replace('ؤ','و')
        text = re.sub(r'\s+',' ',text)
        text = re.sub(r'\s+([.,،؛:!؟»\)])',r'\1',text)
        text = re.sub(r'([(«])\s+',r'\1',text)
        text = self.fix_half_spaces(text)
        return text.strip()

    def fix_half_spaces(self, text):
        prefixes = ['می','نمی','بی','با','از','به','در','که']
        for p in prefixes:
            text = re.sub(f'\\b{p} ',f'{p}\u200c',text)
        return text

# ---------------- Math Processor ----------------
class MathProcessor:
    @staticmethod
    def is_formula(t):
        return bool(re.search(r'\$.*?\$|[∑√±×÷≤≥∞∫≈≠αβγδθλμπρσφω]',t))

    @staticmethod
    def clean_formula(f):
        f = f.strip('$').strip(); f=re.sub(r'\s+',' ',f);return f

    @staticmethod
    def format_formula_for_word(f):
        conv={r'\\alpha':'α',r'\\beta':'β',r'\\gamma':'γ',r'\\pi':'π',r'\\times':'×',r'\\div':'÷'}
        for a,b in conv.items():f=f.replace(a,b)
        return f

# ---------------- Smart Document Generator ----------------
class SmartDocumentGenerator:
    def __init__(self):
        self.doc=Document()
        self.text_processor=PersianTextProcessor()
        self.math_processor=MathProcessor()
        self._setup_document()

    def _setup_document(self):
        s=self.doc.sections[0]
        s.page_height=Inches(11.69)
        s.page_width=Inches(8.27)
        s.left_margin=s.right_margin=s.top_margin=s.bottom_margin=Inches(1)

    def _set_rtl(self,p):
        el=p._element; pPr=el.get_or_add_pPr()
        bidi=OxmlElement('w:bidi');bidi.set(qn('w:val'),'1');pPr.append(bidi)

    def detect_content_type(self, line):
        line=line.strip()
        if not line: return 'empty'
        if re.match(r'^#+\s', line): return 'heading'
        if re.match(r'^شکل\s*\d+', line): return 'figure_caption'
        if re.match(r'^جدول\s*\d+', line): return 'table_caption'
        if self.math_processor.is_formula(line): return 'formula'
        return 'text'

    def add_heading(self,text,level=1):
        text=re.sub(r'^#+\s*','',text)
        text=self.text_processor.clean_text(text)
        h=self.doc.add_heading(level=level)
        r=h.add_run(text)
        r.font.name='B Nazanin';r.font.size=Pt(18-level*2);r.bold=True
        r._element.rPr.rFonts.set(qn('w:cs'),'B Nazanin')
        h.alignment=WD_ALIGN_PARAGRAPH.RIGHT
        self._set_rtl(h)

    def add_formula(self,text):
        m=re.search(r'\$\$(.*?)\$\$|\$(.*?)\$',text)
        if m:
            f=m.group(1) or m.group(2)
            f=self.math_processor.clean_formula(f)
            f=self.math_processor.format_formula_for_word(f)
            p=self.doc.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT   # ← چپ‌چین
            r=p.add_run(f)
            r.font.name='Cambria Math';r.font.size=Pt(14)

    def add_caption(self,text,position='bottom'):
        """زیرنویس شکل یا بالانویس جدول"""
        text=self.text_processor.clean_text(text)
        p=self.doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        self._set_rtl(p)
        r=p.add_run(text)
        r.bold=True
        r.font.name='B Nazanin';r.font.size=Pt(12)
        r._element.rPr.rFonts.set(qn('w:cs'),'B Nazanin')

    def add_mixed_text_paragraph(self,text):
        text=self.text_processor.clean_text(text)
        p=self.doc.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY   # ← Justify (چین دوطرفه)
        p.paragraph_format.line_spacing_rule=WD_LINE_SPACING.ONE_POINT_FIVE
        self._set_rtl(p)
        r=p.add_run(text)
        r.font.name='B Nazanin';r.font.size=Pt(14)
        r._element.rPr.rFonts.set(qn('w:cs'),'B Nazanin')

    def process_text(self,text):
        for line in text.split('\n'):
            t=self.detect_content_type(line)
            if t=='empty': self.doc.add_paragraph()
            elif t=='heading': self.add_heading(line,1)
            elif t=='formula': self.add_formula(line)
            elif t=='figure_caption': self.add_caption(line,'bottom')
            elif t=='table_caption': self.add_caption(line,'top')
            else: self.add_mixed_text_paragraph(line)

    def save_to_stream(self):
        buf=io.BytesIO();self.doc.save(buf);buf.seek(0);return buf

# ---------------- Flask App ----------------
@app.route('/generate',methods=['POST'])
def generate_document():
    try:
        data=request.get_json()
        text=data.get('text','')
        if not text: return jsonify({'error':'متن الزامی است'}),400
        gen=SmartDocumentGenerator()
        gen.process_text(text)
        fs=gen.save_to_stream()
        return send_file(fs,mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',as_attachment=True,download_name='document.docx')
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/')
def home():
    return jsonify({'status':'ok','message':'ProjectWord Persian DOCX Ready 🚀','endpoints':['/health','/generate']})

@app.route('/health')
def health():
    return jsonify({'status':'ok','port':8001})

if __name__=='__main__':
    app.run(debug=False,host='0.0.0.0',port=8001)
