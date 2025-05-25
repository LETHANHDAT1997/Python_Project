# Excel Style Information

This document provides a comprehensive explanation of the `style_info` dictionary returned by the `get_cell_style` method in the `Excel_Style` class from `Python_Excel_Style.py`. It details the structure and valid values for all style attributes to help you set cell styles correctly in `openpyxl` without errors. Each attribute is described with its type and valid values to ensure accurate usage when styling Excel cells.

## Structure of `style_info`

The `style_info` dictionary contains styling information for a specific cell in an Excel workbook. Below is the complete structure with detailed comments listing valid values for each attribute.

```python
style_info = {
    'font': {
        'name': str or None,           # Font name, e.g., 'Calibri', 'Arial', 'Tahoma', 'Times New Roman'
        'size': float or None,         # Font size, a float value, typically from 1.0 to 72.0 (e.g., 8.0, 11.0, 12.0)
        'bold': bool or None,          # Bold text, valid values: True, False
        'italic': bool or None,        # Italic text, valid values: True, False
        'underline': str or None,      # Underline style, valid values: 'single', 'double', 'singleAccounting', 'doubleAccounting'
        'color': str or None,          # Font color, RGB format (e.g., 'FFFF0000' for red, 'FF000000' for black)
        'strikethrough': bool or None, # Strikethrough text, valid values: True, False
        'vertAlign': str or None       # Vertical alignment, valid values: 'superscript', 'subscript', 'baseline'
    },
    'fill': {
        'fill_type': str or None,      # Fill type, valid values: 'solid', 'darkDown', 'darkGray', 'darkGrid', 'darkHorizontal', 
                                       # 'darkTrellis', 'darkUp', 'darkVertical', 'gray0625', 'gray125', 
                                       # 'lightDown', 'lightGray', 'lightGrid', 'lightHorizontal', 
                                       # 'lightTrellis', 'lightUp', 'lightVertical', 'mediumGray'
        'start_color': str or None,    # Start color, RGB format (e.g., 'FFFFFF00' for yellow)
        'end_color': str or None       # End color, RGB format (e.g., 'FF0000FF' for blue)
    },
    'border': {
        'left_style': str or None,     # Left border style, valid values: 'thin', 'medium', 'thick', 'dashed', 'dotted', 
                                       # 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 
                                       # 'slantDashDot'
        'left_color': str or None,     # Left border color, RGB format (e.g., 'FF000000' for black)
        'right_style': str or None,    # Right border style, same values as left_style
        'right_color': str or None,    # Right border color, RGB format
        'top_style': str or None,      # Top border style, same values as left_style
        'top_color': str or None,      # Top border color, RGB format
        'bottom_style': str or None,   # Bottom border style, same values as left_style
        'bottom_color': str or None,   # Bottom border color, RGB format
        'diagonal': bool or None,      # Diagonal border, valid values: True, False
        'diagonal_style': str or None, # Diagonal border style, same values as left_style
        'diagonal_color': str or None  # Diagonal border color, RGB format
    },
    'alignment': {
        'horizontal': str or None,     # Horizontal alignment, valid values: 'general', 'left', 'center', 'right', 
                                       # 'fill', 'justify', 'centerContinuous', 'distributed'
        'vertical': str or None,       # Vertical alignment, valid values: 'top', 'center', 'bottom', 'justify', 'distributed'
        'text_rotation': int or None,  # Text rotation angle, integer from 0 to 180 or 255 (255 for vertical text)
        'wrap_text': bool or None,     # Wrap text, valid values: True, False
        'shrink_to_fit': bool or None, # Shrink to fit, valid values: True, False
        'indent': int or None          # Indent level, non-negative integer (e.g., 0, 1, 2, typically < 15)
    },
    'number_format': str or None,       # Number format, examples: 'General', '#,##0', '#,##0.00', '0.00%', 
                                       # 'mm-dd-yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 
                                       # 'h:mm:ss AM/PM', 'mm:ss', '0.00E+00', '@' (text), or custom formats
    'protection': {
        'locked': bool or None,        # Cell locked, valid values: True, False
        'hidden': bool or None         # Cell hidden, valid values: True, False
    }
}
```

## Detailed Explanation of Attributes and Valid Values

Below is a detailed explanation of each attribute in `style_info`, including their types and valid values, to help you set styles correctly in the `Excel_Style` class methods (e.g., `create_font`, `create_pattern_fill`, `create_border`, `create_alignment`, `create_protection`, `create_named_style`) without causing errors in `openpyxl`.

### Font Attributes
- **`name`**:
  - **Type**: `str` or `None`
  - **Valid Values**: Any valid font name installed in Excel, e.g., `'Calibri'`, `'Arial'`, `'Tahoma'`, `'Times New Roman'`, `'Verdana'`.
  - **Note**: Use common fonts to ensure compatibility across systems.
- **`size`**:
  - **Type**: `float` or `None`
  - **Valid Values**: A float representing font size, typically from 1.0 to 72.0 (e.g., 8.0, 11.0, 12.0).
  - **Note**: Values outside this range may cause rendering issues in Excel.
- **`bold`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (bold), `False` (not bold).
- **`italic`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (italic), `False` (not italic).
- **`underline`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'single'`: Single underline.
    - `'double'`: Double underline.
    - `'singleAccounting'`: Single accounting underline (further from text).
    - `'doubleAccounting'`: Double accounting underline.
  - **Note**: Invalid values (e.g., `'triple'`) will raise an error or be ignored.
- **`color`**:
  - **Type**: `str` or `None`
  - **Valid Values**: RGB color in format `'AARRGGBB'` (e.g., `'FFFF0000'` for red, `'FF000000'` for black).
  - **Note**: The first two characters (`AA`) represent the alpha channel (usually `'FF'` for opaque).
- **`strikethrough`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (strikethrough), `False` (no strikethrough).
- **`vertAlign`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'superscript'`: Text above the baseline (e.g., for exponents).
    - `'subscript'`: Text below the baseline (e.g., for indices).
    - `'baseline'`: Default alignment.
  - **Note**: Use only these values to ensure compatibility.

### Fill Attributes
- **`fill_type`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'solid'`: Uniform color fill.
    - Pattern fills: `'darkDown'`, `'darkGray'`, `'darkGrid'`, `'darkHorizontal'`, `'darkTrellis'`, `'darkUp'`, `'darkVertical'`, `'gray0625'`, `'gray125'`, `'lightDown'`, `'lightGray'`, `'lightGrid'`, `'lightHorizontal'`, `'lightTrellis'`, `'lightUp'`, `'lightVertical'`, `'mediumGray'`.
  - **Note**: For pattern fills, ensure `start_color` and `end_color` are set appropriately for visibility.
- **`start_color`**:
  - **Type**: `str` or `None`
  - **Valid Values**: RGB color in format `'AARRGGBB'` (e.g., `'FFFFFF00'` for yellow).
- **`end_color`**:
  - **Type**: `str` or `None`
  - **Valid Values**: RGB color in format `'AARRGGBB'` (e.g., `'FF0000FF'` for blue).

### Border Attributes
- **`left_style`, `right_style`, `top_style`, `bottom_style`, `diagonal_style`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'thin'`: Thin border.
    - `'medium'`: Medium border.
    - `'thick'`: Thick border.
    - `'dashed'`: Dashed border.
    - `'dotted'`: Dotted border.
    - `'double'`: Double border.
    - `'hair'`: Hairline border (very thin).
    - `'mediumDashDot'`: Medium dash-dot border.
    - `'mediumDashDotDot'`: Medium dash-dot-dot border.
    - `'mediumDashed'`: Medium dashed border.
    - `'slantDashDot'`: Slanted dash-dot border.
  - **Note**: Invalid border styles will raise an error in `openpyxl`.
- **`left_color`, `right_color`, `top_color`, `bottom_color`, `diagonal_color`**:
  - **Type**: `str` or `None`
  - **Valid Values**: RGB color in format `'AARRGGBB'` (e.g., `'FF000000'` for black).
- **`diagonal`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (diagonal border enabled), `False` (no diagonal border).

### Alignment Attributes
- **`horizontal`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'general'`: Default alignment (left for text, right for numbers).
    - `'left'`: Left alignment.
    - `'center'`: Center alignment.
    - `'right'`: Right alignment.
    - `'fill'`: Repeat content to fill the cell.
    - `'justify'`: Justify text.
    - `'centerContinuous'`: Continuous center (used for merged cells).
    - `'distributed'`: Distribute text evenly.
  - **Note**: Some values like `'fill'` or `'distributed'` are less common.
- **`vertical`**:
  - **Type**: `str` or `None`
  - **Valid Values**:
    - `'top'`: Top alignment.
    - `'center'`: Vertical center alignment.
    - `'bottom'`: Bottom alignment.
    - `'justify'`: Vertical justify.
    - `'distributed'`: Vertical distribute.
  - **Note**: Ensure the value matches the intended vertical alignment.
- **`text_rotation`**:
  - **Type**: `int` or `None`
  - **Valid Values**: Integer from 0 to 180 (rotation angle in degrees) or 255 (for vertical text).
  - **Note**: Values outside this range will cause errors or incorrect rendering.
- **`wrap_text`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (wrap text), `False` (no wrap).
- **`shrink_to_fit`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (shrink text to fit cell), `False` (no shrink).
- **`indent`**:
  - **Type**: `int` or `None`
  - **Valid Values**: Non-negative integer (e.g., 0, 1, 2, typically less than 15).
  - **Note**: Large indent values may not display clearly in Excel.

### Number Format
- **`number_format`**:
  - **Type**: `str` or `None`
  - **Valid Values** (common examples):
    - `'General'`: Default format.
    - `'#,##0'`: Integer with thousand separators.
    - `'#,##0.00'`: Number with two decimal places.
    - `'0.00%'`: Percentage with two decimal places.
    - `'mm-dd-yy'`: Date in month-day-year format (2-digit year).
    - `'d-mmm-yy'`: Date with abbreviated month and 2-digit year.
    - `'d-mmm'`: Day and abbreviated month.
    - `'mmm-yy'`: Abbreviated month and year.
    - `'h:mm AM/PM'`: Time in 12-hour format.
    - `'h:mm:ss AM/PM'`: Time with seconds in 12-hour format.
    - `'mm:ss'`: Minutes and seconds.
    - `'0.00E+00'`: Scientific notation.
    - `'@'`: Text format.
    - Custom formats following Excel syntax (e.g., `"0.0 \"kg\""`, `"$#,##0.00"`).
  - **Note**: Incorrect formats will cause errors or display issues in Excel.

### Protection Attributes
- **`locked`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (cell is locked, cannot be edited when sheet is protected), `False` (cell is unlocked).
  - **Note**: Locking takes effect only when the worksheet is protected.
- **`hidden`**:
  - **Type**: `bool` or `None`
  - **Valid Values**: `True` (cell content is hidden when sheet is protected), `False` (cell content is visible).
  - **Note**: Hiding takes effect only when the worksheet is protected.

## Usage Example

When setting styles using methods like `create_font`, `create_pattern_fill`, `create_border`, `create_alignment`, `create_protection`, or `create_named_style`, use the valid values listed above to avoid errors. Below is an example of setting a cell style:

```python
from Python_Excel_Style import Excel_Style

style_manager = Excel_Style()

# Create font
font = style_manager.create_font(
    name='Arial',
    size=12.0,           # Valid float value
    bold=True,
    italic=False,
    underline='single',   # Valid underline style
    color='FFFF0000',     # Valid RGB color
    strikethrough=False,
    vert_align='superscript'  # Valid vertical alignment
)

# Create fill
fill = style_manager.create_pattern_fill(
    fill_type='solid',    # Valid fill type
    start_color='FFFFFF00',  # Valid RGB color
    end_color='FFFFFF00'
)

# Create border
border = style_manager.create_border(
    left_style='thin',    # Valid border style
    left_color='FF000000',  # Valid RGB color
    right_style='thin',
    right_color='FF000000',
    top_style='thin',
    top_color='FF000000',
    bottom_style='thin',
    bottom_color='FF000000',
    diagonal=False,
    diagonal_style=None,
    diagonal_color=None
)

# Create alignment
alignment = style_manager.create_alignment(
    horizontal='center',  # Valid horizontal alignment
    vertical='center',    # Valid vertical alignment
    text_rotation=45,     # Valid rotation angle
    wrap_text=True,
    shrink_to_fit=False,
    indent=1              # Valid indent level
)

# Create protection
protection = style_manager.create_protection(
    locked=True,          # Valid protection value
    hidden=False
)

# Create named style
style_manager.create_named_style(
    name='custom_style',
    font=font,
    fill=fill,
    border=border,
    alignment=alignment,
    number_format='#,##0.00',  # Valid number format
    protection=protection
)
```

## Notes
- **Validation**: Always use the valid values listed above to prevent errors in `openpyxl`. Incorrect values (e.g., `underline='triple'`, `horizontal='middle'`) will raise exceptions or cause rendering issues.
- **Checking Styles**: Use the `get_cell_style` method to verify the applied styles of a cell and ensure they match your expectations.
- **Handling `None`**: Attributes that are not set will return `None` in `style_info`. When setting styles, you can omit attributes by passing `None`.
- **Protection Context**: The `locked` and `hidden` attributes in `protection` only take effect when the worksheet is protected in Excel (e.g., via `worksheet.protection.enable()`).

This document serves as a complete reference for styling Excel cells using the `Excel_Style` class, ensuring you can apply styles accurately and avoid errors.