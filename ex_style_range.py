import xlwings as xw

def ex_style_range(sheet, cell_range, 
                   font_color=None, font_name=None, font_size=None, 
                   bold=None, italic=None, underline=None, strikethrough=None,
                   border=None, border_style='thin', border_type='surround',
                   background_color=None, 
                   horizontal_alignment=None, vertical_alignment=None,
                   number_format=None):
    """
    Applies various styles to a specified range of cells in an Excel worksheet.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - sheet: object
        The Excel worksheet object where the cell range will be styled.
    - cell_range: str
        The range of cells to be styled (e.g., 'A1:B10').
    - font_color: int, optional
        The RGB color value for the font. Defaults to None.
    - font_name: str, optional
        The name of the font to be applied. Defaults to None.
    - font_size: int, optional
        The size of the font to be applied. Defaults to None.
    - bold: bool, optional
        If True, applies bold styling to the font. Defaults to None.
    - italic: bool, optional
        If True, applies italic styling to the font. Defaults to None.
    - border: bool, optional
        If True, applies borders to the specified range. If None, no borders are applied. Defaults to None.
    - border_style: str, optional
        The style of the border to apply (e.g., 'thin'). Defaults to 'thin'.
    - border_type: str, optional
        The type of border to apply ('surround' for surrounding the range or 'all' for all cells). Defaults to 'surround'.
    - background_color: int, optional
        The RGB color value for the background. Defaults to None.
    - horizontal_alignment: str, optional
        The horizontal alignment (e.g., 'left', 'center', 'right'). Defaults to None.
    - vertical_alignment: str, optional
        The vertical alignment (e.g., 'top', 'center', 'bottom'). Defaults to None.
    - number_format: str, optional
        The number format to apply (e.g., '0.00', '$#,##0.00'). Defaults to None.
    - strikethrough: bool, optional
        If True, applies strikethrough styling to the font. Defaults to None.
    - underline: bool, optional
        If True, applies underline styling to the font. Defaults to None.

    Returns:
    - None
        The function does not return any value. It directly modifies the styles of the specified range.
    """

    rng = sheet[cell_range]

    # Thiết lập màu chữ
    if font_color is not None:
        rng.font.color = font_color

    # Thiết lập tên font chữ
    if font_name is not None:
        rng.font.name = font_name

    # Thiết lập kích cỡ chữ
    if font_size is not None:
        rng.font.size = font_size

    # Thiết lập kiểu chữ
    if bold is not None:
        rng.font.bold = bold
    if italic is not None:
        rng.font.italic = italic

    # Thiết lập viền
    if border is not None:
        if border:
            if border_type == 'surround':
                # Thêm viền bao quanh vùng
                for side in ['top', 'bottom', 'left', 'right']:
                    rng.api.Borders[getattr(xw.constants.XlBordersIndex, f'xlEdge{side.capitalize()}')].LineStyle = getattr(xw.constants.XlLineStyle, f'xlLineStyle{border_style.capitalize()}')
            elif border_type == 'all':
                # Thêm viền cho tất cả các ô trong vùng
                for side in ['top', 'bottom', 'left', 'right']:
                    rng.api.Borders[getattr(xw.constants.XlBordersIndex, f'xlEdge{side.capitalize()}')].LineStyle = getattr(xw.constants.XlLineStyle, f'xlLineStyle{border_style.capitalize()}')
                for row in rng.rows:
                    for cell in row:
                        for side in ['top', 'bottom', 'left', 'right']:
                            cell.api.Borders[getattr(xw.constants.XlBordersIndex, f'xlEdge{side.capitalize()}')].LineStyle = getattr(xw.constants.XlLineStyle, f'xlLineStyle{border_style.capitalize()}')

    # Thiết lập màu nền
    if background_color is not None:
        rng.fill.solid()
        rng.fill.fore_color = background_color

    # Thiết lập căn chỉnh
    if horizontal_alignment is not None:
        rng.api.HorizontalAlignment = getattr(xw.constants.XlHAlign, f'xlHAlign{horizontal_alignment.capitalize()}')
    if vertical_alignment is not None:
        rng.api.VerticalAlignment = getattr(xw.constants.XlVAlign, f'xlVAlign{vertical_alignment.capitalize()}')

    # Thiết lập định dạng số
    if number_format is not None:
        rng.number_format = number_format

    # Thiết lập gạch chéo và gạch dưới
    if strikethrough is not None:
        rng.font.strikethrough = strikethrough
    if underline is not None:
        rng.font.underline = underline

# Ví dụ sử dụng
if __name__ == "__main__":
    wb = xw.Book()  # Mở một workbook mới
    sheet = wb.sheets[0]  # Lấy sheet đầu tiên
    ex_style_range(sheet, 'A1:B2', font_color='red', font_name='Arial', font_size=12, bold=True, italic=None, border=True, border_style='thin', border_type='surround')
