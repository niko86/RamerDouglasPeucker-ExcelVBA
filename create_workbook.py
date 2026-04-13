"""
create_workbook.py
==================
Generates a ready-to-use Excel template (.xlsx) for the
Ramer-Douglas-Peucker VBA project.

The resulting file contains:
  Sheet1  — Table1 (X/Y input data pre-loaded from data.csv)
              cell D1: "epsilon" label,      E1: 5.0  (named range "epsilon")
              cell D2: "rdpTargetCount" label, E2: 20 (named range "rdpTargetCount")
  Sheet2  — Table2 (empty X/Y output table)

After running this script:
  1. Open RamerDouglasPeucker_Template.xlsx in Excel.
  2. File > Save As > Excel Macro-Enabled Workbook (.xlsm).
  3. Open the VBE (Alt+F11).
  4. File > Import File → select RamerDouglasPeucker.bas.
  5. In the Project pane double-click ThisWorkbook and paste:

        Private Sub Workbook_Open()
            CreateRDPToolbar
        End Sub

        Private Sub Workbook_BeforeClose(Cancel As Boolean)
            DeleteRDPToolbar
        End Sub

  6. Save, close, and reopen the workbook.
  7. The "RDP Simplify" toolbar appears in the Add-ins tab.

Requirements:
  pip install openpyxl
"""

import csv
import os
import sys

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.workbook.defined_name import DefinedName
except ImportError:
    sys.exit(
        "openpyxl is required. Install it with:\n"
        "    pip install openpyxl"
    )

OUTPUT_FILE = "RamerDouglasPeucker_Template.xlsx"
DATA_FILE   = os.path.join(os.path.dirname(__file__), "data.csv")


def load_data(path):
    """Read X,Y pairs from a CSV file (no header expected)."""
    rows = []
    with open(path, newline="") as f:
        for line in csv.reader(f):
            if len(line) >= 2:
                try:
                    rows.append((float(line[0]), float(line[1])))
                except ValueError:
                    pass  # skip header / non-numeric lines
    return rows


def build_workbook(data):
    wb = openpyxl.Workbook()

    # -----------------------------------------------------------------------
    # Sheet1 — Input
    # -----------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Sheet1"

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="4472C4")

    # Table1 headers
    ws1["A1"] = "X"
    ws1["B1"] = "Y"
    for cell in (ws1["A1"], ws1["B1"]):
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Data rows
    for r, (x, y) in enumerate(data, start=2):
        ws1.cell(row=r, column=1, value=x)
        ws1.cell(row=r, column=2, value=y)

    data_row_count = len(data)
    last_row       = data_row_count + 1   # 1 header + N data rows

    # Excel table over A1:B<last_row>
    tbl1 = Table(
        displayName="Table1",
        ref=f"A1:B{last_row}",
    )
    tbl1.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True,
    )
    ws1.add_table(tbl1)

    # Column widths
    ws1.column_dimensions["A"].width = 12
    ws1.column_dimensions["B"].width = 12

    # -----------------------------------------------------------------------
    # Named-range cells (D/E columns, Sheet1)
    # -----------------------------------------------------------------------
    label_font = Font(italic=True, color="595959")

    ws1["D1"] = "epsilon"
    ws1["D1"].font = label_font
    ws1["E1"] = 5.0
    ws1["E1"].number_format = "0.000"

    ws1["D2"] = "rdpTargetCount"
    ws1["D2"].font = label_font
    ws1["E2"] = 20

    ws1.column_dimensions["D"].width = 18
    ws1.column_dimensions["E"].width = 10

    # Named ranges (workbook-scoped)
    wb.defined_names["epsilon"] = DefinedName(
        "epsilon",
        attr_text="Sheet1!$E$1",
    )
    wb.defined_names["rdpTargetCount"] = DefinedName(
        "rdpTargetCount",
        attr_text="Sheet1!$E$2",
    )

    # -----------------------------------------------------------------------
    # Sheet2 — Output
    # -----------------------------------------------------------------------
    ws2 = wb.create_sheet("Sheet2")

    ws2["A1"] = "X"
    ws2["B1"] = "Y"
    for cell in (ws2["A1"], ws2["B1"]):
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill("solid", fgColor="70AD47")
        cell.alignment = Alignment(horizontal="center")

    # Placeholder row so Excel accepts the table definition
    ws2["A2"] = ""
    ws2["B2"] = ""

    tbl2 = Table(
        displayName="Table2",
        ref="A1:B2",
    )
    tbl2.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium10",
        showRowStripes=True,
    )
    ws2.add_table(tbl2)

    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 12

    return wb


def main():
    if not os.path.exists(DATA_FILE):
        sys.exit(f"Cannot find data file: {DATA_FILE}")

    data = load_data(DATA_FILE)
    if not data:
        sys.exit("No valid X,Y rows found in data.csv")

    print(f"Loaded {len(data)} data points from {DATA_FILE}")

    wb = build_workbook(data)
    wb.save(OUTPUT_FILE)
    print(f"Template saved: {OUTPUT_FILE}")
    print()
    print("Next steps:")
    print("  1. Open the file in Excel")
    print("  2. File > Save As > Excel Macro-Enabled Workbook (.xlsm)")
    print("  3. Alt+F11 > File > Import File > RamerDouglasPeucker.bas")
    print("  4. In ThisWorkbook, paste Workbook_Open and Workbook_BeforeClose")
    print("     (see README.md for the exact code)")
    print("  5. Save, close, reopen — toolbar appears in the Add-ins tab")


if __name__ == "__main__":
    main()
