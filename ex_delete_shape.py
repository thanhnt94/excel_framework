import xlwings as xw
import logging

logging.basicConfig(
        level=logging.DEBUG,
        format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_delete_shape(sheet, shape_name):
    """
    Deletes a specified shape from an Excel worksheet.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - sheet: object
        The Excel worksheet object from which the shape will be deleted.
    - shape_name: str
        The name of the shape to be deleted.

    Returns:
    - None
        The function does not return any value. It logs the status of the deletion operation.

    Logs:
    - Logs debug information when starting the deletion process.
    - Logs an info message when the shape is successfully deleted.
    - Logs an info message if the specified shape is not found in the sheet.
    - Logs an error message if an exception occurs during the deletion process.
    """
    logging.debug(f"Starting to delete shape '{shape_name}' from sheet '{sheet.name}'.")
    shape_found = False

    try:
        for shape in sheet.api.Shapes:
            if shape.Name == shape_name:
                shape.Delete()
                logging.info(f'Deleted shape "{shape_name}" from sheet "{sheet.name}".')
                shape_found = True

        if not shape_found:
            logging.info(f'Shape "{shape_name}" not found in sheet "{sheet.name}".')
    except Exception as e:
        logging.error(f"An error occurred while deleting shape '{shape_name}': {e}")
