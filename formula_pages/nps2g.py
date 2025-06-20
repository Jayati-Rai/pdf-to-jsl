import openpyxl
def formula_2G(output_path):
    
    try:
        output_wb = openpyxl.load_workbook(output_path, keep_vba=True)
        print(f"Loaded {output_path}")

        vetan_bill_ws = output_wb["vetan bill"]
        print(f"Loaded sheet: {vetan_bill_ws.title}")
        nps2g_ws = output_wb["2G"]
        print(f"Loaded sheet: {nps2g_ws.title}")
        last_row = vetan_bill_ws.max_row - 1
        print(f"Last row in vetan bill: {last_row}")
        nps_row = 24
        for i in range(8, last_row, 2):
            l_value = vetan_bill_ws["L" + str(i + 1)].value
            l_serial = vetan_bill_ws["A" + str(i)].value
            print(l_serial=="Total Amount")
    # Skip if value is None or zero
            try:
                if not(l_value is None or float(l_value) == 0.0 or l_serial == "Total Amount"):
                    nps2g_ws["D" + str(nps_row)] = vetan_bill_ws["C" + str(i)].value
                    nps2g_ws["F" + str(nps_row)] = float(vetan_bill_ws["F" + str(i)].value)
                    nps2g_ws["H" + str(nps_row)] = float(vetan_bill_ws["G" + str(i)].value)
                    nps2g_ws["L" + str(nps_row)] = float(vetan_bill_ws["R" + str(i)].value)
                    print(f"Copied employee at row {i}: Name = {vetan_bill_ws['C' + str(i)].value}")
                    print(f"serial number = {l_serial} :Type = {type(l_serial)}")
                    nps_row += 1
            except (ValueError, TypeError) as e:
                print(f"Invalid number at L{i+1}: {l_value} ({e}), skipping.")
                continue

    # Now safe to copy
        
        for row in range(nps_row, 88):
            nps2g_ws.row_dimensions[row].hidden = True
        

        #editing after deletion 
    
        output_wb.save(output_path)
        print(f"2G edited and saved to {output_path}")

    except Exception as e:
        print(f"Error occurred: {e}")
    return