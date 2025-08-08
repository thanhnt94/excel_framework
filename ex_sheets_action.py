import logging
logging.basicConfig(
        level=logging.DEBUG,
        format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

import xlwings as xw

from openpyxl import load_workbook

from PIL import Image
Image.MAX_IMAGE_PIXELS = None

def ex_sheets_action(wb, action):
    """
    Performs specified actions on the sheets of a given Excel workbook.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - wb: object
        The workbook object on which the action will be performed.
    - action: str
        The action to perform on the workbook's sheets. Possible actions include:
        - 'count_total': Count the total number of sheets.
        - 'count_hidden': Count the number of hidden sheets.
        - 'count_visible': Count the number of visible sheets.
        - 'list_sheets': List the names of all sheets.
        - 'list_hidden': List the names of hidden sheets.
        - 'list_visible': List the names of visible sheets.
        - 'delete_hidden': Delete all hidden sheets.

    Returns:
    - int, list, or bool
        Depending on the action:
        - Returns an integer for count actions.
        - Returns a list of sheet names for listing actions.
        - Returns the number of deleted sheets for the delete action.
        - Returns False for unknown actions or if an error occurs.

    Logs:
    - Logs debug information when performing actions on the workbook.
    - Logs an error message if an unknown action is specified or if an error occurs during the action.
    - Logs info messages when hidden sheets are deleted.
    """
    logging.debug(f"Performing action '{action}' on workbook: {wb.name}")
    
    try:
        if action == 'count_total':
            return len(wb.sheets)
        
        elif action == 'count_hidden':
            return sum(1 for sheet in wb.sheets if sheet.api.Visible == 0)

        elif action == 'count_visible':
            return sum(1 for sheet in wb.sheets if sheet.api.Visible == -1)

        elif action == 'list_sheets':
            return [sheet.name for sheet in wb.sheets]

        elif action == 'list_hidden':
            return [sheet.name for sheet in wb.sheets if sheet.api.Visible == 0]

        elif action == 'list_visible':
            return [sheet.name for sheet in wb.sheets if sheet.api.Visible == -1]

        elif action == 'delete_hidden':
            hidden_sheets = [sheet for sheet in wb.sheets if sheet.api.Visible == 0]
            for sheet in hidden_sheets:
                sheet.delete()
            logging.info(f"Deleted {len(hidden_sheets)} hidden sheets.")
            return len(hidden_sheets)

        else:
            logging.error(f"Unknown action: {action}")
            return False

    except Exception as e:
        logging.error(f"Error performing action '{action}' on workbook '{wb.name}': {e}")
        return False
    
if __name__ == "__main__":
    app = xw.App(visible=False)
    wb = app.books.open(r"C:\Users\KNT15083\Downloads\FY24_Q3_UV2小林-3殿宛_1812 _original\RN02753\検討書\TR-V2-S24023.xlsx")

    wb.close()
    app.quit()
