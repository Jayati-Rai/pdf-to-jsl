from flask import Flask, render_template, request, send_file
import pdfplumber
import openpyxl
import os

app = Flask(__name__, template_folder='templates')

@app.route('/')
def home():
    return render_template('pdfReader.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded", 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return "No selected file", 400

    pdf_path = "temp.pdf"
    uploaded_file.save(pdf_path)

    excel_file_path = convert_pdf_to_excel(pdf_path)

    os.remove(pdf_path)  # Clean up temporary PDF file

    return send_file(excel_file_path, as_attachment=True, download_name="converted.xlsx")

def convert_pdf_to_excel(pdf_path):
    # Create an Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "vetan bill"
    with pdfplumber.open(pdf_path) as pdf:
        # Iterate through each page and extract tables
        row_idx = 1
        for page in pdf.pages:
            tables = page.extract_table()
            if tables:
                for row in tables:
                    if row:  # Avoid empty rows
                        sheet.append(row)
                    row_idx += 1

    excel_file_path = "Vetan Bill.xlsm"
    workbook.save(excel_file_path)
    return excel_file_path

if __name__ == '__main__':
    app.run(debug=True)
