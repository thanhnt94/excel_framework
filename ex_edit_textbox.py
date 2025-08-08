import xlwings as xw
import logging
import time

logging.basicConfig(
        level=logging.DEBUG,
        format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_edit_textbox(sheet, shape_name, new_text):
    """
    Edits the text of a specified textbox shape in an Excel worksheet.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - sheet: object
        The Excel worksheet object containing the textbox to be edited.
    - shape_name: str
        The name of the textbox shape to be edited.
    - new_text: str
        The new text to set for the specified textbox.

    Returns:
    - None
        The function does not return any value. It logs the status of the edit operation.

    Logs:
    - Logs debug information when starting the editing process.
    - Logs an info message when the shape is successfully updated with the new text.
    - Logs an error message if the specified shape does not exist in the sheet or if an error occurs during the editing process.
    """
    logging.debug(f"Starting to edit textbox '{shape_name}' in sheet '{sheet.name}'.")
    try:
        if shape_name in [s.Name for s in sheet.api.Shapes]:
            shape = sheet.api.Shapes(shape_name)
            
            original_top = shape.Top
            original_left = shape.Left
            
            shape.TextFrame.Characters().Text = new_text
            
            shape.TextFrame.AutoSize = True
            
            shape.Top = original_top
            shape.Left = original_left
            
            logging.info(f"Successfully updated shape '{shape_name}' in the sheet '{sheet.name}' with new text: '{new_text}'.")
        else:
            logging.error(f"Shape '{shape_name}' does not exist in the sheet '{sheet.name}'.")
    except Exception as e:
        logging.error(f"An error occurred while editing shape '{shape_name}' in the sheet '{sheet.name}': {e}")
