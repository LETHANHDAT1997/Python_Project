# Excel_WorkBook Class Documentation

## Overview
The `Excel_WorkBook` class is a Python utility for interacting with Excel files using the `openpyxl` library. It provides methods to create, manipulate, and format Excel workbooks and sheets, with robust error handling and support for various Excel operations.

## Dependencies
- Python 3.x
- `openpyxl` library
- `os`, `re`, and `uuid` standard libraries

## Class Initialization
```python
excel = Excel_WorkBook(str_path_file_excel, str_name_sheet="Sheet")
```
- **Purpose**: Initializes an Excel workbook, either by loading an existing file or creating a new one.
- **Parameters**:
  - `str_path_file_excel` (str): Path to the Excel file.
  - `str_name_sheet` (str, optional): Name of the sheet to work with (default: "Sheet").
- **Behavior**:
  - If the file exists, loads it and checks for the specified sheet. If the sheet doesn't exist, creates it.
  - If the file doesn't exist, creates a new workbook with the specified sheet.
  - Sets the specified sheet as active.
- **Exceptions**: Raises an exception if initialization fails (e.g., invalid file path).

## Methods

### 1. `__check_name_sheet__(str_name_sheet)`
- **Purpose**: Checks if a sheet exists in the workbook.
- **Parameters**:
  - `str_name_sheet` (str): Name of the sheet to check.
- **Returns**: `True` if the sheet exists, `False` otherwise.
- **Internal Use**: Used by other methods to validate sheet names.

### 2. `__validate_cell_reference(cell_ref)`
- **Purpose**: Validates a cell reference (e.g., "A1", "B2").
- **Parameters**:
  - `cell_ref` (str): Cell reference to validate.
- **Returns**: `True` if valid (e.g., matches pattern `^[A-Z]+[1-9][0-9]*$`), `False` otherwise.
- **Internal Use**: Ensures cell references are correctly formatted.

### 3. `create_sheet(str_name_sheet, overwrite=False)`
- **Purpose**: Creates a new sheet in the workbook.
- **Parameters**:
  - `str_name_sheet` (str): Name of the new sheet.
  - `overwrite` (bool, optional): If `True`, overwrites existing sheet with the same name (default: `False`).
- **Returns**: `True` if successful, `False` if the sheet exists and `overwrite=False` or on error.
- **Example**:
  ```python
  excel.create_sheet("NewSheet", overwrite=True)
  ```

### 4. `set_active_sheet(str_name_sheet)`
- **Purpose**: Sets the specified sheet as the active sheet.
- **Parameters**:
  - `str_name_sheet` (str): Name of the sheet to activate.
- **Returns**: `True` if successful, `False` if the sheet doesn't exist.
- **Example**:
  ```python
  excel.set_active_sheet("Sheet1")
  ```

### 5. `get_sheet(str_name_sheet)`
- **Purpose**: Retrieves the specified sheet object.
- **Parameters**:
  - `str_name_sheet` (str): Name of the sheet.
- **Returns**: Sheet object if it exists, `None` otherwise.
- **Example**:
  ```python
  sheet = excel.get_sheet("Sheet1")
  ```

### 6. `get_sheet_names()`
- **Purpose**: Returns a list of all sheet names in the workbook.
- **Returns**: List of sheet names.
- **Example**:
  ```python
  sheets = excel.get_sheet_names()
  print(sheets)  # e.g., ['Sheet1', 'Sheet2']
  ```

### 7. `write_column(str_name_sheet, column, list_content, start_row=1)`
- **Purpose**: Writes data to a specified column.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `column` (str or int): Column letter (e.g., "A") or index (e.g., 1).
  - `list_content` (list): Data to write.
  - `start_row` (int, optional): Starting row (default: 1).
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.write_column("Sheet1", "A", ["Header", "Data1", "Data2"], start_row=1)
  ```

### 8. `write_row(str_name_sheet, row, list_content, start_column=1)`
- **Purpose**: Writes data to a specified row.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `row` (int): Row number.
  - `list_content` (list): Data to write.
  - `start_column` (int, optional): Starting column index (default: 1).
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.write_row("Sheet1", 1, ["Header1", "Header2", "Header3"])
  ```

### 9. `write_cell(str_name_sheet, cell_ref, content)`
- **Purpose**: Writes data to a specific cell.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_ref` (str or tuple): Cell reference (e.g., "A1" or `(row, column)`).
  - `content`: Data to write (any type supported by `openpyxl`).
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.write_cell("Sheet1", "B7", 25000)
  excel.write_cell("Sheet1", (7, 3), 23000)
  ```

### 10. `read_cell(str_name_sheet, cell_ref)`
- **Purpose**: Reads data from a specific cell.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_ref` (str or tuple): Cell reference (e.g., "A1" or `(row, column)`).
- **Returns**: Cell value if successful, `None` on error.
- **Example**:
  ```python
  value = excel.read_cell("Sheet1", "B7")
  print(value)  # e.g., 25000
  ```

### 11. `read_range(str_name_sheet, start_cell, end_cell)`
- **Purpose**: Reads a range of cells (e.g., "A1:C3").
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `start_cell` (str): Top-left cell of the range (e.g., "A1").
  - `end_cell` (str): Bottom-right cell of the range (e.g., "C3").
- **Returns**: List of lists containing cell values, or `None` on error.
- **Example**:
  ```python
  data = excel.read_range("Sheet1", "A1", "C3")
  print(data)  # e.g., [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
  ```

### 12. `set_column_width(str_name_sheet, column, width)`
- **Purpose**: Sets the width of a column.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `column` (str or int): Column letter (e.g., "A") or index (e.g., 1).
  - `width` (float): Width in Excel units.
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.set_column_width("Sheet1", "A", 20)
  ```

### 13. `set_row_height(str_name_sheet, row, height)`
- **Purpose**: Sets the height of a row.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `row` (int): Row number.
  - `height` (float): Height in Excel units.
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.set_row_height("Sheet1", 1, 30)
  ```

### 14. `format_cells(str_name_sheet, cell_range, pattern_fill=None, font=None, border=None, alignment=None, number_format=None)`
- **Purpose**: Applies formatting to a range of cells.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_range` (str): Range of cells (e.g., "B7:C7").
  - `pattern_fill` (PatternFill, optional): Fill style (e.g., solid color).
  - `font` (Font, optional): Font style (e.g., size, bold).
  - `border` (Border, optional): Border style.
  - `alignment` (Alignment, optional): Text alignment.
  - `number_format` (str, optional): Number format (e.g., "#,##0.00").
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
  pattern_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
  font = Font(name='Tahoma', size=11, bold=True)
  border = Border(left=Side(style="double"), right=Side(style="double"), top=Side(style="double"), bottom=Side(style="double"))
  alignment = Alignment(horizontal='center', vertical='center')
  excel.format_cells("Sheet1", "B7:C7", pattern_fill=pattern_fill, font=font, border=border, alignment=alignment, number_format="#,##0.00")
  ```

### 15. `insert_image(str_name_sheet, cell_ref, image_path, scale_width=1.0, scale_height=1.0)`
- **Purpose**: Inserts an image into the specified cell.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_ref` (str): Cell to anchor the image (e.g., "A1").
  - `image_path` (str): Path to the image file.
  - `scale_width` (float, optional): Width scale factor (default: 1.0).
  - `scale_height` (float, optional): Height scale factor (default: 1.0).
- **Returns**: `True` if successful, `False` if the sheet or image doesn't exist or on error.
- **Example**:
  ```python
  excel.insert_image("Sheet1", "A1", "image.png", scale_width=0.5, scale_height=0.5)
  ```

### 16. `set_formula(str_name_sheet, cell_ref, formula)`
- **Purpose**: Sets an Excel formula in a specific cell.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_ref` (str or tuple): Cell reference (e.g., "D7" or `(7, 4)`).
  - `formula` (str): Excel formula (e.g., "=SUM(B7:C7)").
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.set_formula("Sheet1", "D7", "=SUM(B7:C7)")
  ```

### 17. `freeze_panes(str_name_sheet, cell_ref)`
- **Purpose**: Freezes panes at the specified cell.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_ref` (str): Cell to freeze panes at (e.g., "B2").
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.freeze_panes("Sheet1", "B2")
  ```

### 18. `add_sort_filter(str_name_sheet, cell_range)`
- **Purpose**: Adds sort and filter functionality to a cell range.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_range` (str): Range of cells (e.g., "A1:C7").
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.add_sort_filter("Sheet1", "A1:C7")
  ```

### 19. `merge_cells(str_name_sheet, cell_range)`
- **Purpose**: Merges cells in the specified range.
- **Parameters**:
  - `str_name_sheet` (str): Target sheet name.
  - `cell_range` (str): Range of cells to merge (e.g., "B2:D4").
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.merge_cells("Sheet1", "B2:D4")
  ```

### 20. `save(path_save=None)`
- **Purpose**: Saves the workbook to the specified or original path.
- **Parameters**:
  - `path_save` (str, optional): Path to save the file (default: original path).
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.save()  # Saves to original path
  excel.save("new_file.xlsx")  # Saves to new path
  ```

### 21. `close()`
- **Purpose**: Closes the workbook.
- **Returns**: `True` if successful, `False` on error.
- **Example**:
  ```python
  excel.close()
  ```

## Example Usage
```python
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# Initialize workbook
excel = Excel_WorkBook("example.xlsx", "DataSheet")

# Write data
excel.write_column("DataSheet", "A", ["Name", "Alice", "Bob"])
excel.write_row("DataSheet", 1, ["Name", "Age", "Score"])
excel.write_cell("DataSheet", "B2", 25)

# Apply formatting
pattern_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
font = Font(name='Tahoma', size=11, bold=True)
border = Border(left=Side(style="double"), right=Side(style="double"), top=Side(style="double"), bottom=Side(style="double"))
alignment = Alignment(horizontal='center', vertical='center')
excel.format_cells("DataSheet", "A1:C3", pattern_fill=pattern_fill, font=font, border=border, alignment=alignment)

# Set formula and formatting
excel.set_formula("DataSheet", "C4", "=SUM(C2:C3)")
excel.set_column_width("DataSheet", "A", 20)
excel.set_row_height("DataSheet", 1, 30)
excel.add_sort_filter("DataSheet", "A1:C3")
excel.freeze_panes("DataSheet", "B2")

# Save and close
excel.save()
excel.close()
```

## Notes
- All methods include error handling and print informative messages for debugging.
- Cell references can be provided as strings (e.g., "A1") or tuples (e.g., `(row, column)`).
- The class supports advanced Excel features like formatting, formulas, and image insertion.
- Always ensure the `openpyxl` library is installed (`pip install openpyxl`).
- Image insertion requires a valid image file path and may not work in Pyodide environments due to file I/O restrictions.