#!/usr/bin/env python3
"""
Creates two sample Excel files with varied formatting to test merger.py.

    python create_samples.py

Produces:
    sample_Q1_Sales.xlsx   — 3 sheets: Overview, North Region, South Region
    sample_Products.xlsx   — 2 sheets: Inventory, Pricing
"""

import os
import sys

try:
    import win32com.client
    import pywintypes
except ImportError:
    sys.exit("pywin32 not found. Run setup.ps1 first.")


def rgb(r, g, b):
    """Convert RGB to Excel's BGR integer colour format."""
    return r + g * 256 + b * 65536


# Common palette
C_HEADER_BG  = rgb(31,  73,  125)   # dark blue
C_HEADER_FG  = rgb(255, 255, 255)   # white
C_ALT_ROW    = rgb(217, 226, 243)   # light blue
C_TOTAL_BG   = rgb(255, 242, 204)   # pale yellow
C_TOTAL_FG   = rgb(156,  87,   0)   # dark gold
C_GREEN_HDR  = rgb(55,  138,   0)   # dark green
C_BLUE_HDR   = rgb(31,   73, 125)   # dark blue
C_RED        = rgb(192,   0,   0)   # dark red
C_GREEN_CELL = rgb(226, 239, 218)   # mint
C_BORDER     = rgb(128, 128, 128)   # grey


def apply_header(cell, text, bg=C_HEADER_BG, fg=C_HEADER_FG):
    cell.Value = text
    cell.Font.Bold = True
    cell.Font.Color = fg
    cell.Interior.Color = bg
    cell.HorizontalAlignment = -4108   # xlCenter


def apply_border(rng):
    for edge in (7, 8, 9, 10):   # xlEdgeLeft/Right/Top/Bottom
        rng.Borders(edge).LineStyle = 1   # xlContinuous
        rng.Borders(edge).Color = C_BORDER
        rng.Borders(edge).Weight = 2      # xlThin
    rng.Borders(12).LineStyle = 1         # xlInsideVertical
    rng.Borders(12).Color = C_BORDER
    rng.Borders(11).LineStyle = 1         # xlInsideHorizontal
    rng.Borders(11).Color = C_BORDER


# ── Sheet builders ────────────────────────────────────────────────────────────

def build_overview(ws):
    ws.Name = "Overview"
    ws.Tab.Color = rgb(31, 73, 125)

    # Title
    ws.Range("A1:E1").Merge()
    c = ws.Cells(1, 1)
    c.Value = "Q1 Sales Overview"
    c.Font.Bold = True
    c.Font.Size = 14
    c.Font.Color = C_HEADER_FG
    c.Interior.Color = C_HEADER_BG
    c.HorizontalAlignment = -4108

    # Headers
    headers = ["Region", "Jan", "Feb", "Mar", "Total"]
    for col, h in enumerate(headers, 1):
        apply_header(ws.Cells(3, col), h)

    # Data rows
    data = [
        ("North",  82400, 91200, 87600),
        ("South",  74300, 68900, 79100),
        ("East",   61200, 72500, 68800),
        ("West",   88100, 94300, 91700),
    ]
    for row_i, (region, jan, feb, mar) in enumerate(data, 4):
        ws.Cells(row_i, 1).Value = region
        ws.Cells(row_i, 2).Value = jan
        ws.Cells(row_i, 3).Value = feb
        ws.Cells(row_i, 4).Value = mar
        ws.Cells(row_i, 5).Formula = f"=SUM(B{row_i}:D{row_i})"
        if row_i % 2 == 0:
            ws.Range(f"A{row_i}:E{row_i}").Interior.Color = C_ALT_ROW
        for col in range(2, 6):
            ws.Cells(row_i, col).NumberFormat = "#,##0"

    # Totals row
    total_row = 4 + len(data)
    ws.Cells(total_row, 1).Value = "TOTAL"
    ws.Cells(total_row, 1).Font.Bold = True
    for col in range(2, 6):
        ws.Cells(total_row, col).Formula = f"=SUM({chr(64+col)}4:{chr(64+col)}{total_row-1})"
        ws.Cells(total_row, col).Font.Bold = True
        ws.Cells(total_row, col).NumberFormat = "#,##0"
    ws.Range(f"A{total_row}:E{total_row}").Interior.Color = C_TOTAL_BG
    ws.Range(f"A{total_row}:E{total_row}").Font.Color = C_TOTAL_FG

    apply_border(ws.Range(f"A3:E{total_row}"))

    # Column widths
    ws.Columns("A").ColumnWidth = 12
    for col in "BCDE":
        ws.Columns(col).ColumnWidth = 11
    ws.Rows(1).RowHeight = 28


def build_region(ws, name, color, sales_data):
    ws.Name = name
    ws.Tab.Color = color

    ws.Range("A1:D1").Merge()
    c = ws.Cells(1, 1)
    c.Value = f"{name} — Monthly Detail"
    c.Font.Bold = True
    c.Font.Size = 13
    c.Font.Color = C_HEADER_FG
    c.Interior.Color = color
    c.HorizontalAlignment = -4108

    for col, h in enumerate(["Sales Rep", "Product", "Units", "Revenue"], 1):
        apply_header(ws.Cells(3, col), h, bg=color)

    for row_i, (rep, product, units, revenue) in enumerate(sales_data, 4):
        ws.Cells(row_i, 1).Value = rep
        ws.Cells(row_i, 2).Value = product
        ws.Cells(row_i, 3).Value = units
        ws.Cells(row_i, 4).Value = revenue
        ws.Cells(row_i, 4).NumberFormat = "$#,##0"
        if row_i % 2 == 0:
            ws.Range(f"A{row_i}:D{row_i}").Interior.Color = C_GREEN_CELL if color == C_GREEN_HDR else C_ALT_ROW

    apply_border(ws.Range(f"A3:D{3 + len(sales_data)}"))
    ws.Columns("A:B").ColumnWidth = 16
    ws.Columns("C:D").ColumnWidth = 12
    ws.Rows(1).RowHeight = 24


def build_inventory(ws):
    ws.Name = "Inventory"
    ws.Tab.Color = rgb(0, 112, 192)

    ws.Range("A1:F1").Merge()
    c = ws.Cells(1, 1)
    c.Value = "Product Inventory"
    c.Font.Bold = True
    c.Font.Size = 14
    c.Font.Color = C_HEADER_FG
    c.Interior.Color = rgb(0, 112, 192)
    c.HorizontalAlignment = -4108

    for col, h in enumerate(["SKU", "Product Name", "Category", "In Stock", "Reorder Point", "Status"], 1):
        apply_header(ws.Cells(3, col), h, bg=rgb(0, 112, 192))

    products = [
        ("SKU-001", "Wireless Mouse",      "Electronics",  245, 50,  "OK"),
        ("SKU-002", "USB-C Hub",           "Electronics",   38, 40,  "LOW"),
        ("SKU-003", "Standing Desk Mat",   "Furniture",    112, 20,  "OK"),
        ("SKU-004", "Monitor Arm",         "Furniture",     67, 15,  "OK"),
        ("SKU-005", "Noise Cancel. Headp", "Electronics",   12, 30,  "LOW"),
        ("SKU-006", "Webcam HD",           "Electronics",   89, 25,  "OK"),
        ("SKU-007", "Ergonomic Chair",     "Furniture",     23, 10,  "OK"),
        ("SKU-008", "Laptop Stand",        "Accessories",  156, 35,  "OK"),
    ]
    for row_i, (sku, name, cat, stock, reorder, status) in enumerate(products, 4):
        ws.Cells(row_i, 1).Value = sku
        ws.Cells(row_i, 2).Value = name
        ws.Cells(row_i, 3).Value = cat
        ws.Cells(row_i, 4).Value = stock
        ws.Cells(row_i, 5).Value = reorder
        ws.Cells(row_i, 6).Value = status
        if row_i % 2 == 0:
            ws.Range(f"A{row_i}:F{row_i}").Interior.Color = C_ALT_ROW
        if status == "LOW":
            ws.Cells(row_i, 6).Font.Bold = True
            ws.Cells(row_i, 6).Font.Color = C_RED

    apply_border(ws.Range(f"A3:F{3 + len(products)}"))
    for col, w in zip("ABCDEF", [10, 22, 14, 10, 14, 8]):
        ws.Columns(col).ColumnWidth = w
    ws.Rows(1).RowHeight = 26


def build_pricing(ws):
    ws.Name = "Pricing"
    ws.Tab.Color = rgb(112, 48, 160)

    ws.Range("A1:E1").Merge()
    c = ws.Cells(1, 1)
    c.Value = "Product Pricing"
    c.Font.Bold = True
    c.Font.Size = 14
    c.Font.Color = C_HEADER_FG
    c.Interior.Color = rgb(112, 48, 160)
    c.HorizontalAlignment = -4108

    for col, h in enumerate(["SKU", "Product Name", "Cost", "Retail Price", "Margin %"], 1):
        apply_header(ws.Cells(3, col), h, bg=rgb(112, 48, 160))

    pricing = [
        ("SKU-001", "Wireless Mouse",      18.50,  39.99),
        ("SKU-002", "USB-C Hub",           22.00,  49.99),
        ("SKU-003", "Standing Desk Mat",   14.75,  34.99),
        ("SKU-004", "Monitor Arm",         31.00,  69.99),
        ("SKU-005", "Noise Cancel. Headp", 68.00, 149.99),
        ("SKU-006", "Webcam HD",           29.50,  79.99),
        ("SKU-007", "Ergonomic Chair",    185.00, 449.99),
        ("SKU-008", "Laptop Stand",        12.25,  29.99),
    ]
    for row_i, (sku, name, cost, retail) in enumerate(pricing, 4):
        ws.Cells(row_i, 1).Value = sku
        ws.Cells(row_i, 2).Value = name
        ws.Cells(row_i, 3).Value = cost
        ws.Cells(row_i, 3).NumberFormat = "$#,##0.00"
        ws.Cells(row_i, 4).Value = retail
        ws.Cells(row_i, 4).NumberFormat = "$#,##0.00"
        ws.Cells(row_i, 5).Formula = f"=ROUND((D{row_i}-C{row_i})/D{row_i}*100,1)"
        ws.Cells(row_i, 5).NumberFormat = '0.0"%"'
        if row_i % 2 == 0:
            ws.Range(f"A{row_i}:E{row_i}").Interior.Color = rgb(237, 233, 245)

    apply_border(ws.Range(f"A3:E{3 + len(pricing)}"))
    for col, w in zip("ABCDE", [10, 22, 12, 14, 10]):
        ws.Columns(col).ColumnWidth = w
    ws.Rows(1).RowHeight = 26


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))

    print("Starting Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # ── File 1: Q1 Sales ─────────────────────────────────────────────────
        path1 = os.path.join(out_dir, "sample_Q1_Sales.xlsx")
        print("Creating sample_Q1_Sales.xlsx  (3 sheets: Overview, North Region, South Region)")
        wb = excel.Workbooks.Add()

        # Remove extra default sheets, keep exactly 3
        while wb.Sheets.Count < 3:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        while wb.Sheets.Count > 3:
            wb.Sheets(wb.Sheets.Count).Delete()

        build_overview(wb.Sheets(1))
        build_region(wb.Sheets(2), "North Region", C_GREEN_HDR, [
            ("Alice Johnson", "Widget Pro",   142, 18460),
            ("Alice Johnson", "Gadget Plus",   87, 13050),
            ("Bob Martinez",  "Widget Pro",   210, 27300),
            ("Bob Martinez",  "Super Tool",    65,  9750),
            ("Carol White",   "Gadget Plus",  198, 29700),
        ])
        build_region(wb.Sheets(3), "South Region", C_BLUE_HDR, [
            ("David Lee",   "Widget Pro",   118, 15340),
            ("David Lee",   "Super Tool",    93, 13950),
            ("Emma Davis",  "Gadget Plus",  176, 26400),
            ("Emma Davis",  "Widget Pro",    84, 10920),
            ("Frank Brown", "Super Tool",   201, 30150),
        ])

        wb.SaveAs(path1, FileFormat=51)
        wb.Close()
        print(f"  Saved: {path1}")

        # ── File 2: Products ──────────────────────────────────────────────────
        path2 = os.path.join(out_dir, "sample_Products.xlsx")
        print("Creating sample_Products.xlsx  (2 sheets: Inventory, Pricing)")
        wb = excel.Workbooks.Add()

        while wb.Sheets.Count < 2:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        while wb.Sheets.Count > 2:
            wb.Sheets(wb.Sheets.Count).Delete()

        build_inventory(wb.Sheets(1))
        build_pricing(wb.Sheets(2))

        wb.SaveAs(path2, FileFormat=51)
        wb.Close()
        print(f"  Saved: {path2}")

    finally:
        excel.Quit()

    print()
    print("Done. Now run:")
    print(f"  python merger.py \"{out_dir}\"")


if __name__ == "__main__":
    main()
