import openpyxl
from copy import copy
def formula_prapatrag(output_path):
    try:
        output_wb = openpyxl.load_workbook(output_path, keep_vba=True)
        print(f"Loaded {output_path}")

        vetan_bill_ws = output_wb["vetan bill"]
        print(f"Loaded sheet: {vetan_bill_ws.title}")
        
        last_row = vetan_bill_ws.max_row - 1
            
        prapatra_g_ws = output_wb["प्रपत्र-ग"]
       
        prapatra_g_ws["D28"] = float(vetan_bill_ws["F"+str(last_row)].value)
        prapatra_g_ws["D29"] = float(vetan_bill_ws["G"+str(last_row)].value )
        prapatra_g_ws["D30"] = float(vetan_bill_ws["H"+str(last_row)].value)
        prapatra_g_ws["D35"] = float(vetan_bill_ws["L"+str(last_row+1)].value)
        prapatra_g_ws["D36"] = float(vetan_bill_ws["L"+str(last_row+1)].value)
        prapatra_g_ws["D37"] = float(vetan_bill_ws["L"+str(last_row)].value)
        prapatra_g_ws["D39"] = float(vetan_bill_ws["M"+str(last_row)].value)
    

        output_wb.save(output_path)
        print(f"Prapatra-g edited and saved to {output_path}")

    except Exception as e:
        print(f"Error occurred: {e}")
    return