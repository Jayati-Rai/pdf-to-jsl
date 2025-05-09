# 📄 Automated Vetan Bill Generator

An end-to-end automation tool for generating salary grant sheets (`Vetan Bill`) for aided school teachers and staff, following the official format used in Uttar Pradesh. This tool eliminates repetitive manual work by replicating Excel templates—including merged cells, styling, and formulas—and inserting extracted data automatically.

Currently, only **प्रपत्र-ग** (Prapatra-G) is supported. Full support for all five pages (प्रपत्र-क to प्रपत्र-घ) is coming soon.

## ⚡ Why Use This?

Manually preparing each grant page takes **15–40 minutes** depending on data and format availability. This tool:
- Automates formatting, merged cells, styling, and formulas
- Generates the output in **under 30 seconds**
- Produces a `.xlsm` file compatible with official workflows

## 🛠 Technologies Used

| Tool/Library     | Purpose                                         |
|------------------|--------------------------------------------------|
| `Python`         | Core logic and scripting                        |
| `openpyxl`       | Excel template handling with styles & formulas  |
| `pdfplumber`     | (Optional) Extract data from scanned PDFs       |
| `Flask`          | Web interface for uploading and generating files|
| `VBA support`    | Retains macros in the output `.xlsm` file       |

## 📋 Features

- ✅ Reads Excel templates with merged cells and formatting
- ✅ Copies formulas and styles to the new workbook
- ✅ Automatically creates "प्रपत्र-ग" in the output file
- ✅ Fast and accurate generation
- 🔜 More grant pages (क, ख, घ, ड) coming soon

## 📂 Folder Structure

├── templates/

│ └── prapatra-g.xlsx # Source Excel template

├── static/

│ └── ... # (Optional) Web assets

├── vetan_generator.py # Core logic for template copy

├── app.py # Flask app for UI

├── requirements.txt # Python dependencies

└── README.md

## 🚀 Getting Started

1. **Clone the repository**
   
<pre>
git clone https://github.com/your-username/vetan-bill-generator.git
cd vetan-bill-generator
</pre>

2. **Install dependencies**

<pre>
pip install -r requirements.txt
</pre>

3. **Run the app**
   
<pre> python app.py</pre>

4. **Use the Interface**
   
  Upload your Excel template and choose the destination file.
  Click Generate to create प्रपत्र-ग automatically.

🧪 Sample Code
<pre>
output_wb = openpyxl.load_workbook(output_path, keep_vba=True)
new_sheet = output_wb.create_sheet(title="प्रपत्र-ग")
template_wb = openpyxl.load_workbook(template_path, keep_vba=True)
template_ws = template_wb["प्रपत्र-ग"]
</pre>

**Copy formatting, values, and formulas**
<pre> for row in template_ws.iter_rows():
    for cell in row:
        new_cell = new_sheet[cell.coordinate]
        new_cell.value = cell.value
        new_cell.font = copy(cell.font)
        ...
</pre>

## 🛡 Limitations

• Currently supports only one page (प्रपत्र-ग)

• Requires structured Excel templates

• PDF data extraction logic is in progress


## 📈 Future Plans

• Add support for all grant sheets:प्रपत्र-ग, 105, 2G, Challan etc.

• Build PDF-to-Excel data mapping layer

• Enhance UI with form previews and progress tracking

• Add multilingual support for labels


## 👨‍💻 Author

**Jayati Rai**

📍 Ghazipur, Uttar Pradesh

💼 Android + Python Developer

✉️ jayatirai3@gmail.com



## 📃 License

This project is licensed under the MIT License.

⚠️ This tool is tailored to the grant format used by aided schools in U.P., India. Please adapt the templates if you're using it for a different regional format.
