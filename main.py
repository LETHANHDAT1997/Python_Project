from Python_Excel_Lib import Excel_WorkBook
from Python_Excel_Style import Excel_Style

# Initialize workbook and style manager
excel = Excel_WorkBook("styled_data_sample.xlsx", "ASPHALT")
style_manager = Excel_Style()

# Data from the image (simplified for brevity, you can expand as needed)
headers = [
    "DISTRICT NAME", "STREET NAME", "BRAVO", "FORMAN", "BACKFILL Forman", "Date", "WORK ORDER", 
    "Excavation", "", "", "", "Sand", "", "SUBBASE", "", "Gravel", "", "Old Material", "", "", "", "", ""
]
sub_headers = [
    "", "", "", "", "", "", "", "L", "W", "D", "m³", "D", "m3", "D", "m3", "D", "m3", "D", "m3", "L", "Remarks", "DONE"
]
data_rows = [
    ["K 14", "FROM ROAD 30", "111", "AHMED JAMAL", "Mudassir Imran", "1-Sep-20", "6511245944", "2", "2", "1", "4", "0.5", "1", "0.5", "1.4", "0", "0", "0", "0", "NO"],
    ["UMM SALEEM", "AI Ain AL AZIZIYAH", "111", "AHMED JAMAL", "Mudassir Imran", "1-Sep-20", "6516597070", "2.5", "2", "1", "5", "0.7", "1.4", "0.8", "7.2", "0", "0", "0", "0", "NO"],
    # Add more rows as needed...
    ["SAFA", "KHALIDA BINTE HESHAM", "710", "GHUFRAN", "Ibrahim Rehan", "1-Sep-20", "6515191235", "3.5", "2", "1.5", "10.5", "0.5", "3.5", "1", "36", "0", "0", "0", "0", "NO"]
]

# Write headers
excel.write_row("ASPHALT", 1, headers)
excel.write_row("ASPHALT", 2, sub_headers)

# Add AutoFilter for the header row (A1:W1)
excel.add_sort_filter("ASPHALT", "A1:W1")

# Write data
for i, row in enumerate(data_rows, start=3):
    excel.write_row("ASPHALT", i, row)

# Write totals (simplified, you can calculate dynamically if needed)
excel.write_cell("ASPHALT", "H70", "TOTAL")
excel.write_cell("ASPHALT", "I70", "374.145")
excel.write_cell("ASPHALT", "K70", "121.517")
excel.write_cell("ASPHALT", "M70", "241.928")
excel.write_cell("ASPHALT", "O70", "11.3")

# Write summary table
summary_data = [
    ["Excavation", "374.145"], ["Sand", "121.517"], ["Subbase", "241.928"], 
    ["Gravel", "11.3"], ["Old Material", "0"]
]
for i, (label, value) in enumerate(summary_data, start=72):
    excel.write_cell("ASPHALT", f"I{i}", label)
    excel.write_cell("ASPHALT", f"J{i}", value)

# Create styles using Excel_Style
# Header style
header_font = style_manager.create_font(bold=True, color="FFFFFF")
header_fill = style_manager.create_pattern_fill(fill_type="solid", start_color="92D050")
header_border = style_manager.create_border(
    left_style="thin", right_style="thin", top_style="thin", bottom_style="thin",
    left_color="000000", right_color="000000", top_color="000000", bottom_color="000000"
)
header_alignment = style_manager.create_alignment(horizontal="center", vertical="center")

# Data style
data_border = style_manager.create_border(
    left_style="thin", right_style="thin", top_style="thin", bottom_style="thin",
    left_color="000000", right_color="000000", top_color="000000", bottom_color="000000"
)
data_alignment = style_manager.create_alignment(horizontal="center", vertical="center")
number_format = "0.0"

# Total row styles
total_font = style_manager.create_font(bold=True, color="FF0000")  # For "TOTAL"
total_value_font = style_manager.create_font(bold=True, color="000000")  # For total values

# Summary table styles
summary_colors = ["92D050", "C0C0C0", "FF6666", "FFFF99", "D3D3D3"]  # Excavation, Sand, Subbase, Gravel, Old Material
summary_fills = [style_manager.create_pattern_fill(fill_type="solid", start_color=color) for color in summary_colors]
summary_alignment = style_manager.create_alignment(horizontal="center", vertical="center")

# Apply styles
# Headers
excel.format_cells("ASPHALT", "A1:W2", font=header_font, pattern_fill=header_fill, border=header_border, alignment=header_alignment)

# Data cells
excel.format_cells("ASPHALT", "A3:W69", border=data_border, alignment=data_alignment)
excel.format_cells("ASPHALT", "I3:I69", number_format=number_format)
excel.format_cells("ASPHALT", "J3:J69", number_format=number_format)
excel.format_cells("ASPHALT", "K3:K69", number_format=number_format)
excel.format_cells("ASPHALT", "M3:M69", number_format=number_format)
excel.format_cells("ASPHALT", "O3:O69", number_format=number_format)
excel.format_cells("ASPHALT", "Q3:Q69", number_format=number_format)

# Total row
excel.format_cells("ASPHALT", "H70", font=total_font)
excel.format_cells("ASPHALT", "I70:O70", font=total_value_font, number_format=number_format)

# Summary table
for i, fill in enumerate(summary_fills, start=72):
    excel.format_cells("ASPHALT", f"I{i}:J{i}", pattern_fill=fill, border=data_border, alignment=summary_alignment)

# Conditional formatting for "DONE" column (S)
yes_fill = style_manager.create_pattern_fill(fill_type="solid", start_color="92D050")
no_fill = style_manager.create_pattern_fill(fill_type="solid", start_color="FF6666")
style_manager.create_conditional_formatting(
    condition_type="cellIs",
    cell_range="S3:S69",
    workbook=excel.workbook,
    sheet_name="ASPHALT",
    operator="equal",
    value='"YES"',
    fill=yes_fill  # Sửa từ font thành fill
)
style_manager.create_conditional_formatting(
    condition_type="cellIs",
    cell_range="S3:S69",
    workbook=excel.workbook,
    sheet_name="ASPHALT",
    operator="equal",
    value='"NO"',
    fill=no_fill  # Sửa từ font thành fill
)

# Adjust column widths for better visibility
for col in range(1, 24):  # Columns A to W
    excel.set_column_width("ASPHALT", col, 15)

# Freeze the header row
excel.freeze_panes("ASPHALT", "G2")  # Freeze the header row

# Save the workbook
excel.save()
excel.close()