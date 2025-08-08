import xlwings as xw
import logging
import time

logging.basicConfig(
        level=logging.DEBUG,
        format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_insert_textbox(sheet, shape_name, textbox_content, 
                      position='A1', width=100, height=20, 
                      orientation=1, placement=3, 
                      font_name='Verdana', font_size=10, 
                      bold=False, italic=False, underline=False, 
                      text_color=None, text_alignment='left', 
                      fill_color=None, line_color=0x000000, 
                      line_weight=1, 
                      left_margin_cm=0.2, right_margin_cm=0.2, 
                      top_margin_cm=0.1, bottom_margin_cm=0.1, 
                      auto_size=True, shadow=False, 
                      shadow_color=None, shadow_offset=1, 
                      text_rotation=0, text_wrap=False, 
                      z_order=1, locked=False):
    
    """
    Inserts a textbox into a specified Excel sheet with customizable properties.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - sheet: object
        The Excel worksheet object where the textbox will be inserted.
    - shape_name: str
        The name of the textbox shape to be created.
    - textbox_content: str
        The content text to display inside the textbox.
    - position: str, optional
        The cell reference (e.g., 'A1') where the textbox will be positioned. Defaults to 'A1'.
    - width: float, optional
        The width of the textbox in points. Defaults to 100.
    - height: float, optional
        The height of the textbox in points. Defaults to 20.
    - orientation: int, optional
        The orientation of the textbox (1 for horizontal). Defaults to 1.
    - placement: int, optional
        The placement option for the textbox (3 for move and size with cells). Defaults to 3.
    - font_name: str, optional
        The font name for the textbox text. Defaults to 'Verdana'.
    - font_size: int, optional
        The font size for the textbox text. Defaults to 10.
    - bold: bool, optional
        If True, text will be bold. Defaults to False.
    - italic: bool, optional
        If True, text will be italic. Defaults to False.
    - underline: bool, optional
        If True, text will be underlined. Defaults to False.
    - text_color: int, optional
        The RGB color value for the text. Defaults to None.
    - text_alignment: str, optional
        The alignment of the text ('left', 'center', or 'right'). Defaults to 'left'.
    - fill_color: int, optional
        The RGB color value for the textbox background. Defaults to None.
    - line_color: int, optional
        The RGB color value for the textbox border. Defaults to 0x000000 (black).
    - line_weight: int, optional
        The weight of the border line. Defaults to 1.
    - left_margin_cm: float, optional
        Left margin in centimeters. Defaults to 0.2 cm.
    - right_margin_cm: float, optional
        Right margin in centimeters. Defaults to 0.2 cm.
    - top_margin_cm: float, optional
        Top margin in centimeters. Defaults to 0.1 cm.
    - bottom_margin_cm: float, optional
        Bottom margin in centimeters. Defaults to 0.1 cm.
    - auto_size: bool, optional
        If True, the textbox will adjust its size automatically. Defaults to True.
    - shadow: bool, optional
        If True, adds a shadow to the textbox. Defaults to False.
    - shadow_color: int, optional
        The RGB color value for the shadow. Defaults to None.
    - shadow_offset: int, optional
        The offset for the shadow. Defaults to 1.
    - text_rotation: int, optional
        The rotation angle for the text in degrees. Defaults to 0.
    - text_wrap: bool, optional
        If True, enables text wrapping within the textbox. Defaults to False.
    - z_order: int, optional
        The z-order of the textbox (1 for bring to front, 0 for send to back). Defaults to 1.
    - locked: bool, optional
        If True, locks the textbox to prevent editing. Defaults to False.

    Returns:
    - None
        The function does not return any value. It logs the status of the insertion operation.

    Logs:
    - Logs debug information when starting the textbox insertion process.
    - Logs an error message if an exception occurs during the insertion process.
    """

    logging.debug(f"Starting to insert textbox '{shape_name}' in sheet '{sheet.name}' at position '{position}'.")

    try:
        logging.debug(f"Processing sheet: {sheet.name}")

        left = sheet.range(position).left
        top = sheet.range(position).top

        # Tạo textbox
        textbox = sheet.api.Shapes.AddTextbox(
            Orientation=orientation,
            Left=left,
            Top=top,
            Width=width,
            Height=height
        )

        textbox.Name = shape_name
        textbox.TextFrame.Characters().Text = textbox_content

        text_frame = textbox.TextFrame
        text_frame.Characters().Font.Name = font_name
        text_frame.Characters().Font.Size = font_size

        # Đặt kiểu chữ
        text_frame.Characters().Font.Bold = bold
        text_frame.Characters().Font.Italic = italic
        text_frame.Characters().Font.Underline = underline

        # Căn chỉnh văn bản
        if text_alignment == 'center':
            text_frame.HorizontalAlignment = 1  # Center
        elif text_alignment == 'right':
            text_frame.HorizontalAlignment = 3  # Right
        else:
            text_frame.HorizontalAlignment = 2  # Left

        # Đặt AutoSize và Placement
        textbox.TextFrame.AutoSize = auto_size  
        textbox.Placement = placement

        # Đặt lề
        text_frame.MarginLeft = left_margin_cm * 28.35 
        text_frame.MarginRight = right_margin_cm * 28.35
        text_frame.MarginTop = top_margin_cm * 28.35
        text_frame.MarginBottom = bottom_margin_cm * 28.35

        # Đặt màu nền
        if fill_color:
            textbox.Fill.ForeColor.RGB = fill_color

        # Đặt màu văn bản
        if text_color:
            text_frame.Characters().Font.Color = text_color

        # Đặt đường viền
        line = textbox.Line
        line.ForeColor.RGB = line_color 
        line.Weight = line_weight

        # Đặt bóng
        if shadow:
            textbox.Shadow = True
            if shadow_color:
                textbox.ShadowColor = shadow_color
            textbox.ShadowOffset = shadow_offset

        # Đặt độ xoay văn bản
        textbox.TextFrame.Orientation = text_rotation

        # Đặt tự động xuống dòng
        text_frame.WrapText = text_wrap

        # Đặt thứ tự hiển thị
        textbox.ZOrder(z_order)  # 1 = mời lên trên cùng, 0 = mời xuống dưới cùng

        # Đặt khóa
        textbox.Locked = locked

        logging.debug(f"Added textbox '{shape_name}' with content '{textbox_content}' to sheet '{sheet.name}'.")

    except Exception as e:
        logging.error(f"Error processing sheet '{sheet.name}': {e}")
        return f"Error processing sheet '{sheet.name}': {e}"

if __name__ == '__main__':
    import os
    import sys

    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    # from functions.excel_io import *

    file_path = r"C:\Users\KNT15083\Downloads\241220\FY24_Q3_UV2小林-3殿宛_1812 _original\RN02753\検討書\J2-24-P094.xlsx"

