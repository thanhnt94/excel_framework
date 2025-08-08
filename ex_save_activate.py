import xlwings as xw
import pygetwindow as gw
import logging
import time

logging.basicConfig(
    level=logging.DEBUG,
    format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

def ex_save_activate(file_path):
    """
    Saves the currently active workbook in the active Excel application to a specified file path.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - file_path: str
        The path where the active workbook will be saved.

    Returns:
    - None
        The function does not return any value. It logs the status of the save operation.

    Logs:
    - Logs debug information while checking for the active Excel window.
    - Logs an info message when Excel is the active application.
    - Logs debug information when saving the workbook.
    - Logs an info message if the workbook is saved successfully.
    - Logs a warning if a TypeError is encountered during saving, and attempts to save again.
    - Logs debug information when the workbook is closed.
    """
    # Define the name of the Excel application window
    excel_window_name = "Excel"
    logging.debug("Checking for active Excel window...")

    # Get the active window
    active_window = gw.getActiveWindow()

    # Check if the active window title contains "Excel"
    while active_window is None or excel_window_name not in active_window.title:
        logging.debug("Waiting for Excel to be the active application...")
        time.sleep(1)
        active_window = gw.getActiveWindow()

    logging.info("Excel is now the active application.")

    # Connect to the active Excel application
    app = xw.apps.active
    # Get the active workbook
    wb = app.books.active

    # Save the active workbook to the specified file path
    logging.debug(f"Saving workbook to {file_path}...")

    try:
        wb.save(file_path)
        logging.info(f"Workbook saved successfully to {file_path}.")
    except TypeError:
        logging.warning("TypeError encountered while saving. Attempting to save without compatibility checks...")
        wb.save(file_path)
        logging.info(f"Workbook saved successfully to {file_path} after handling TypeError.")

    # Close the workbook
    wb.close()
    app.quit()

    logging.debug("Workbook closed.")

if __name__ == '__main__':
    file_path = r"C:\exeBuild\2D Drawing Download\save.xlsx"
    ex_save_activate(file_path)

