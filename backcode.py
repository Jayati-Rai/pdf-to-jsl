from flask import Flask, render_template, request, send_file
from vba_praptra_g import fill_data_into_template
from formula_pages.prapatrag import formula_prapatrag
from formula_pages.nps2g import formula_2G
from copying import fill_data_into_newfile
from formula_pages.tds import formula_tds
from formula_pages.gpf import formula_gpf
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

    return send_file(excel_file_path, as_attachment=True, download_name="Vetan Bill.xlsm")

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
    template_path = "prapatra-g.xlsx"
    fill_data_into_newfile(template_path,excel_file_path,"प्रपत्र-ग","A1:J50")
    fill_data_into_newfile(template_path,excel_file_path,"2G","A1:N92")
    fill_data_into_newfile(template_path,excel_file_path,"105","A1:L26")
    fill_data_into_newfile(template_path,excel_file_path,"gpf challan","A1:U43")
    fill_data_into_newfile(template_path,excel_file_path,"281","A1:AJ48")
    # print("We are okay till here")
    try:
        employee_count_cell = "A" + str(sheet.max_row - 3)
        employee_count_raw = sheet[employee_count_cell].value
        print(f"Raw employee count from {employee_count_cell}: {employee_count_raw}")

        employee_count = int(employee_count_raw)
        print("Validated employee count:", employee_count)

        # formula calling
        
        formula_prapatrag(excel_file_path)
        formula_2G(excel_file_path)
        formula_tds(excel_file_path)
        formula_gpf(excel_file_path)
    except Exception as e:
        print(f"Error processing employee count or inserting rows: {e}")

    return excel_file_path

if __name__ == '__main__':
    app.run(debug=True)
