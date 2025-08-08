import xlwings as xw
import logging
import time

logging.basicConfig(
        level=logging.DEBUG,
        format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_copy_textbox(source_sheet, target_sheet, shape_name, coordinates):
    """
    Copies a textbox shape from a source Excel sheet to a target Excel sheet at specified coordinates.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - source_sheet: object
        The Excel worksheet object from which the textbox will be copied.
    - target_sheet: object
        The Excel worksheet object to which the textbox will be pasted.
    - shape_name: str
        The name of the textbox shape to be copied.
    - coordinates: str or tuple
        The target position for the pasted textbox, specified as a cell reference (e.g., 'A1') or as a tuple of (top, left) coordinates.

    Returns:
    - None
        The function does not return any value. It logs the status of the copy operation.

    Logs:
    - Logs debug information when starting the copy process and when the shape is found.
    - Logs a warning if the specified shape is not found in the source sheet.
    - Logs an info message when the position of the pasted shape is successfully set.
    - Logs an error message if an exception occurs during the copying process.
    """

    def cell_to_coordinates(sheet, cell_reference):
        """
        Convert a cell reference (like 'A1') to its top and left coordinates.
        """
        cell = sheet.Range(cell_reference)
        top = cell.Top
        left = cell.Left
        return top, left

    logging.debug(f"Starting to copy shape '{shape_name}' from '{source_sheet.name}' to '{target_sheet.name}' at coordinates '{coordinates}'.")
    try:
        shapes_list = [s.Name for s in source_sheet.api.Shapes]
        if shape_name in shapes_list:
            logging.debug(f'Found shape "{shape_name}" in "{source_sheet.name}".')
            time.sleep(0.1)
            source_sheet.shapes[shape_name].api.Copy()
            time.sleep(0.1)
            target_sheet.api.Paste()
            logging.debug(f'Successfully pasted shape "{shape_name}" to {target_sheet.name}.')

            pasted_shape = target_sheet.api.Shapes(target_sheet.api.Shapes.Count)
            pasted_shape.TextFrame.AutoSize = True
            
            # Check if coordinates is a cell reference or a tuple
            if isinstance(coordinates, str):  # If it's a cell reference
                coordinates = cell_to_coordinates(target_sheet, coordinates)

            # Set position using the provided coordinates
            pasted_shape.Top = coordinates[0]  # Set Top to the first coordinate
            pasted_shape.Left = coordinates[1]  # Set Left to the second coordinate
            
            logging.info(f'Set position of pasted "{shape_name}" to coordinates "{coordinates}" with (Top: {pasted_shape.Top}, Left: {pasted_shape.Left}).')
        else:
            logging.warning(f"Shape '{shape_name}' not found in the source sheet. Available shapes: {shapes_list}.")
    except Exception as e:
        logging.error(f"An error occurred while copying shape '{shape_name}': {e}")

if __name__ == '__main__':
    import os
    import sys

    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    # from functions.excel_io import *

    file_path = r"C:\Users\KNT15083\Downloads\241220\FY24_Q3_UV2小林-3殿宛_1812 _original\RN02753\検討書\J2-24-P094.xlsx"

