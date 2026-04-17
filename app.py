from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import zipfile
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_pdf():
    pdf_file = request.files['pdf']
    excel_file = request.files['excel']

    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)

    pdf_file.save(pdf_path)
    excel_file.save(excel_path)

    # Read Excel
    df = pd.read_excel(excel_path)

    # Validate columns
    if 'page' not in df.columns or 'name' not in df.columns:
        return "Excel must have 'page' and 'name' columns"

    reader = PdfReader(pdf_path)

    output_files = []

    for index, row in df.iterrows():
        page_number = int(row['page']) - 1  # zero-based index
        name = str(row['name'])

        if page_number >= len(reader.pages):
            return f"Page {page_number+1} does not exist in PDF"

        writer = PdfWriter()
        writer.add_page(reader.pages[page_number])

        # Clean filename
        safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_")).strip()
        safe_name = safe_name.replace(" ", "_")

        file_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.pdf")

        with open(file_path, "wb") as f:
            writer.write(f)

        output_files.append(file_path)

    # Create ZIP
    zip_filename = f"{uuid.uuid4()}.zip"
    zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in output_files:
            zipf.write(file, os.path.basename(file))

    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run()