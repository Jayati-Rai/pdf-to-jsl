import openpyxl
def formula_gpf(output_path):
    
    try:
        output_wb = openpyxl.load_workbook(output_path, keep_vba=True)
        print(f"Loaded {output_path}")

        vetan_bill_ws = output_wb["vetan bill"]
        print(f"Loaded sheet: {vetan_bill_ws.title}")
        gpf_ws = output_wb["gpf challan"]
        print(f"Loaded sheet: {gpf_ws.title}")

        last_row = vetan_bill_ws.max_row
        gpf_ws["R15"] = float(vetan_bill_ws["L" + str(last_row)].value)
        gpf_ws["R16"] = float(vetan_bill_ws["N" + str(last_row-1)].value)

        output_wb.save(output_path)
        print(f"GPF challan edited and saved to {output_path}")
    except KeyError as e:
        print(f"Sheet not found: {e}")
        return