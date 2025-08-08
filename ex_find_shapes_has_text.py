import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

import warnings
warnings.filterwarnings("ignore")

from PIL import Image
Image.MAX_IMAGE_PIXELS = None

import openpyxl
from openpyxl import load_workbook

def ex_find_shapes_has_text(sheet, search_text, exact_match=False):
    """
    Searches for specified text within shapes in an Excel worksheet and returns the names of matching shapes.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - sheet: object
        The Excel worksheet object in which to search for shapes.
    - search_text: str
        The text to search for within the shapes.
    - exact_match: bool, optional
        If True, searches for an exact match of the search_text. Defaults to False for partial matches.

    Returns:
    - list
        A list of names of shapes that contain the specified text. Returns an empty list if no matches are found.

    Logs:
    - Logs debug information for each shape checked, including its name, type, and position.
    - Logs an info message when a matching shape is found, indicating whether it was an exact or partial match.
    - Logs a warning if no shapes contain the specified text.
    """

    logging.debug("Starting the search for text in shapes.")
    
    found_shapes = []  # List to store names of shapes containing the text

    for shape in sheet.api.Shapes:
        logging.debug(f"Shape Name: '{shape.Name}', Type: {shape.Type}, Position: ({shape.Left}, {shape.Top})")
        try:
            text = shape.TextFrame.Characters().Text
            logging.debug(f"Checking shape: '{shape.Name}', Content: '{text}'")

            if exact_match:
                if text == search_text:
                    found_shapes.append(shape.Name)
                    message = f"Found in shape (exact match): '{text}', Name: '{shape.Name}', Position: ({shape.Left}, {shape.Top})"
                    logging.info(message)
            else:
                if search_text in text:
                    found_shapes.append(shape.Name)
                    message = f"Found in shape (partial match): '{text}', Name: '{shape.Name}', Position: ({shape.Left}, {shape.Top})"
                    logging.info(message)
        except Exception as e:
            logging.debug(f"Error while accessing shape '{shape.Name}': {e}")

    if found_shapes:
        return found_shapes  # Return the list of found shape names
    else:
        logging.warning(f"Text '{search_text}' not found in any shape in the sheet.")
        return []  # Return an empty list if no shapes were found

if __name__ == "__main__":
    file_path = r"C:\Users\KNT15083\Downloads\521\summary.xlsx"
