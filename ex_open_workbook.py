import xlwings as xw
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

import xlwings as xw
import logging

def ex_open_workbook(file_path, read_only=False, password=None):
    """
    Opens an Excel workbook and returns the application and workbook objects.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - file_path: str
        The path to the Excel file to be opened.
    - read_only: bool, optional
        If True, opens the workbook in read-only mode. Defaults to False.
    - password: str, optional
        The password required to open the workbook, if applicable.

    Returns:
    - tuple
        A tuple containing the application object and the workbook object if opened successfully.
        Returns (False, False) if the workbook cannot be opened.

    Logs:
    - Logs debug information when starting to open the Excel file.
    - Logs an info message when the Excel file is opened successfully.
    - Logs an error message if the file cannot be opened, including specific messages for password-related issues.
    """
    logging.debug(f"Starting to open the Excel file '{file_path}' with read_only={read_only}.")

    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        workbook = app.books.open(file_path, password=password, read_only=read_only, ignore_read_only_recommended=True)
        logging.info(f"The Excel file '{file_path}' is opened successfully.")
        return app, workbook

    except Exception as e:
        error_message = ""

        if 'password' in str(e).lower():
            error_message = f"Failed to open the Excel file '{file_path}' because this file has been set with a password and cannot be opened."
            logging.error(error_message)
        else:
            error_message = f"Failed to open the Excel file '{file_path}'. Error: {str(e)}"
            logging.error(error_message)

        if 'workbook' in locals():
            workbook.close()

    return False, False

if __name__ == '__main__':
    try:
        file_path = r"C:\Users\KNT15083\Documents\Book1111.xlsx"
        file_path = r"C:\Users\KNT15083\Downloads\fffsda.xlsx"
        file_path = r"C:\Users\KNT15083\Documents\Database1.xlsx"
        file_path = r"C:\Users\KNT15083\Downloads\Q3設通リスト_LT_柏原-3殿宛.xlsx"

    except Exception as e:
        pass

