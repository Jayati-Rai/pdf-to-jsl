import openpyxl
def formula_tds(output_path):
    
    try:
        output_wb = openpyxl.load_workbook(output_path, keep_vba=True)
        print(f"Loaded {output_path}")

        vetan_bill_ws = output_wb["vetan bill"]
        print(f"Loaded sheet: {vetan_bill_ws.title}")

        tds_ws = output_wb["281"]
        print(f"Loaded sheet: {tds_ws.title}")
        last_row = vetan_bill_ws.max_row
        tds_ws["U23"]= float(vetan_bill_ws["M" + str(last_row-1)].value)

        output_wb.save(output_path)
        print(f"TDS challan edited and saved to {output_path}")

    except KeyError as e:
        print(f"Sheet not found: {e}")
        return