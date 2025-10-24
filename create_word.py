    # ---- ساخت جدول‌ها ----
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
        cols = max(len(r) for r in rows)
        rows = [r + [''] * (cols - len(r)) for r in rows]
        # حذف separator markdown
        if len(rows) > 1 and all(set(cell.strip()) <= {'-', ':', '|'} for cell in rows[1]):
            rows.pop(1)

        # ایجاد جدول واقعی LTR (بدون bidiVisual یا tblDir)
        tbl = OxmlElement('w:tbl')
        tblPr = OxmlElement('w:tblPr')

        # حاشیه‌ها
        tblBorders = OxmlElement('w:tblBorders')
        for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '12')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)
        tblPr.append(tblBorders)
        tbl.append(tblPr)

        # ساخت سطرها به ترتیب طبیعی (بدون معکوس)
        for rindex, row in enumerate(rows):
            tr = OxmlElement('w:tr')
            for cindex, cell_text in enumerate(row):
                cell = OxmlElement('w:tc')
                tcPr = OxmlElement('w:tcPr')

                # رنگ پس‌زمینه برای هدر
                if rindex == 0:
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), 'D9E2F3')
                    tcPr.append(shd)

                # حاشیه‌های سلول
                borders = OxmlElement('w:tcBorders')
                for b in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{b}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '6')
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), '000000')
                    borders.append(border)
                tcPr.append(borders)
                cell.append(tcPr)

                # پاراگراف داخلی راست‌چین ولی بدون bidi
                p = OxmlElement('w:p')
                pPr = OxmlElement('w:pPr')
                jc = OxmlElement('w:jc')
                jc.set(qn('w:val'), 'right')   # راست‌چین
                pPr.append(jc)
                p.append(pPr)

                # افزودن متن
                parts = self._parse_bold_text(cell_text)
                for part in parts:
                    r = OxmlElement('w:r')
                    rPr = OxmlElement('w:rPr')

                    fonts = OxmlElement('w:rFonts')
                    fonts.set(qn('w:cs'), 'B Nazanin')
                    fonts.set(qn('w:ascii'), 'B Nazanin')
                    rPr.append(fonts)

                    if part['bold']:
                        bold = OxmlElement('w:b')
                        bold.set(qn('w:val'), 'true')
                        rPr.append(bold)

                    r.append(rPr)
                    t = OxmlElement('w:t')
                    t.text = part['text']
                    r.append(t)
                    p.append(r)
                cell.append(p)
                tr.append(cell)
            tbl.append(tr)

        self.doc._body._element.append(tbl)
        self.doc.add_paragraph()
