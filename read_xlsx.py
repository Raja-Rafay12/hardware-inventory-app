import sys
import os
import zipfile
import xml.etree.ElementTree as ET

def read_xlsx_fallback(file_path):
    print("Using zip/xml fallback parser...", file=sys.stderr)
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            # Read shared strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_data = z.read('xl/sharedStrings.xml')
                root = ET.fromstring(ss_data)
                # The namespace is usually http://schemas.openxmlformats.org/spreadsheetml/2006/main
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root.findall('ns:si', ns):
                    # Can have <t> or multiple <r><t>
                    t_elems = si.findall('.//ns:t', ns)
                    text = "".join([t.text for t in t_elems if t.text])
                    shared_strings.append(text)
            
            # Read sheet1
            sheet_data = z.read('xl/worksheets/sheet1.xml')
            root = ET.fromstring(sheet_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            rows = {}
            for row in root.findall('.//ns:row', ns):
                row_idx = int(row.attrib['r'])
                row_data = {}
                for cell in row.findall('ns:c', ns):
                    ref = cell.attrib['r'] # e.g. A1, B2
                    # Get column letter
                    col_letter = ''.join([c for c in ref if c.isalpha()])
                    val_elem = cell.find('ns:v', ns)
                    val = ""
                    if val_elem is not None and val_elem.text:
                        val = val_elem.text
                        t = cell.attrib.get('t')
                        if t == 's': # shared string
                            idx = int(val)
                            if 0 <= idx < len(shared_strings):
                                val = shared_strings[idx]
                    row_data[col_letter] = val
                rows[row_idx] = row_data
            
            # Convert to list of lists or print
            if not rows:
                print("No data found in sheet1.xml")
                return
            
            max_row = max(rows.keys())
            # Find all unique columns
            all_cols = set()
            for r in rows.values():
                all_cols.update(r.keys())
            
            # Sort columns alphabetically by letter length and then letter
            sorted_cols = sorted(list(all_cols), key=lambda x: (len(x), x))
            
            # Print rows
            import csv
            writer = csv.writer(sys.stdout)
            writer.writerow(sorted_cols)
            for r_idx in range(1, max_row + 1):
                if r_idx in rows:
                    row_dict = rows[r_idx]
                    writer.writerow([row_dict.get(col, "") for col in sorted_cols])
    except Exception as e:
        print(f"Fallback parser error: {e}", file=sys.stderr)

def main():
    if len(sys.argv) < 2:
        print("Usage: python read_xlsx.py <file_path>")
        sys.exit(1)
        
    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Error: file not found at {file_path}")
        sys.exit(1)
        
    try:
        import pandas as pd
        print("Using pandas to read excel...", file=sys.stderr)
        df = pd.read_excel(file_path)
        df.to_csv(sys.stdout, index=False)
    except ImportError:
        try:
            import openpyxl
            print("Using openpyxl to read excel...", file=sys.stderr)
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            import csv
            writer = csv.writer(sys.stdout)
            for row in sheet.iter_rows(values_only=True):
                # Filter out None row if all are None
                if any(x is not None for x in row):
                    writer.writerow([str(x) if x is not None else "" for x in row])
        except ImportError:
            read_xlsx_fallback(file_path)

if __name__ == "__main__":
    main()
