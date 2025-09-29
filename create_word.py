import io
import base64
import csv
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

app = Flask(__name__)

def set_paragraph_format(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = paragraph.runs[0]
    run.font.name = 'B Nazanin'
    run.font.size = Pt(14)
    paragraph.paragraph_format.line_spacing = 1

@app.route("/create_word", methods=["POST"])
def create_word_endpoint():
    data = request.get_json()
    if not data or "text" not in data or "mapping" not in data:
        return jsonify({"error": "Invalid input"}), 400

    text = data["text"]
    mapping = data["mapping"]

    doc = Document()

    # اضافه کردن متن و جایگذاری تگ‌ها
    lines = text.split("\n")
    for line in lines:
        if line.strip().startswith("[REF_IMG_"):
            ref_id = line.strip()[5:-1]  # گرفتن ID
            img_entry = next((img for img in mapping["images"] if img["id"] == ref_id), None)
            if img_entry:
                img_bytes = base64.b64decode(img_entry["data"])
                img_stream = io.BytesIO(img_bytes)
                doc.add_picture(img_stream, width=None)
                cap_para = doc.add_paragraph(img_entry["caption"])
                set_paragraph_format(cap_para)
        elif line.strip().startswith("[REF_TAB_"):
            ref_id = line.strip()[5:-1]
            tab_entry = next((tab for tab in mapping["tables"] if tab["id"] == ref_id), None)
            if tab_entry:
                csv_bytes = base64.b64decode(tab_entry["data"])
                csv_stream = io.StringIO(csv_bytes.decode('utf-8'))
                df = pd.read_csv(csv_stream)
                table = doc.add_table(rows=1, cols=len(df.columns))
                # header
                hdr_cells = table.rows[0].cells
                for i, col_name in enumerate(df.columns):
                    hdr_cells[i].text = col_name
                # data
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, cell_val in enumerate(row):
                        row_cells[i].text = str(cell_val)
                cap_para = doc.add_paragraph(tab_entry["caption"])
                set_paragraph_format(cap_para)
        else:
            para = doc.add_paragraph(line)
            set_paragraph_format(para)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)

    return send_file(
        out_stream,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="final.docx"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8001)
