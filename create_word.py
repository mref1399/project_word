def add_text(self, text):
    text = self.text_processor.clean_text(text)
    if not text:
        return

    p = self.doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    self._set_rtl(p)

    # پاراگراف پیش‌فرض را فارسی کن تا Style هم منطبق شود
    rPr = p._element.get_or_add_pPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'B Nazanin')
    rPr.append(rFonts)

    parts = self._parse_bold_text(text)
    for part in parts:
        run = p.add_run(part['text'])
        if re.search(r'[A-Za-z0-9]', part['text']):  # English segment
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
            run._element.rPr.rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            run.font.size = Pt(12)
        else:  # Persian segment
            run.font.name = 'B Nazanin'
            run._element.rPr.rFonts.set(qn('w:cs'), 'B Nazanin')
            run.font.size = Pt(14)
        if part.get('bold'):
            run.bold = True
