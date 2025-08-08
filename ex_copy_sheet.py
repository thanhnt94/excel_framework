import openpyxl
import xlwings as xw
import logging

logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s | %(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_copy_sheet(source_file_path, destination_file_path, sheet_identifier=None, paste_position=None):
    """
    Copies a specified sheet from a source Excel workbook to a destination Excel workbook.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - source_file_path: str
        The path to the source Excel file from which to copy the sheet.
    - destination_file_path: str
        The path to the destination Excel file where the sheet will be copied.
    - sheet_identifier: int or str, optional
        The index (1-based) or name of the sheet to copy from the source workbook. 
        If not provided, the active sheet will be copied.
    - paste_position: int, optional
        The position (1-based) in the destination workbook where the sheet will be pasted. 
        If not provided, the sheet will be pasted at the end.

    Returns:
    - None
        The function does not return any value. It logs the status of the copy operation.

    Logs:
    - Logs an error message if the source file or sheet cannot be found.
    - Logs an error message if the paste position is invalid.
    - Logs an info message when the sheet is successfully copied and the destination workbook is saved.
    """
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        
        # Open the source workbook
        source_wb = app.books.open(source_file_path)
        # Open or create the destination workbook
        try:
            destination_wb = app.books.open(destination_file_path)
        except FileNotFoundError:
            destination_wb = app.books.add()
        
        # Determine the sheet to copy
        if sheet_identifier is None:
            # Copy the active sheet if no identifier is provided
            source_sheet = source_wb.api.ActiveSheet
        elif isinstance(sheet_identifier, int):
            # Use the sheet index (1-based)
            if 1 <= sheet_identifier <= len(source_wb.sheets):
                source_sheet = source_wb.sheets[sheet_identifier - 1]  # Convert to 0-based index
            else:
                logging.error("Invalid sheet index. Must be between 1 and the number of sheets.")
                return
        else:
            # Use the sheet name
            if sheet_identifier not in [sheet.name for sheet in source_wb.sheets]:
                logging.error(f"Sheet '{sheet_identifier}' does not exist in the source workbook.")
                return
            source_sheet = source_wb.sheets[sheet_identifier]
        
        # Copy the sheet to the destination workbook
        if paste_position is None:
            # Paste at the end if no position is specified
            source_sheet.copy(after=destination_wb.sheets[destination_wb.sheets.count])
        else:
            # Check if the specified paste position is valid
            if paste_position < 1 or paste_position > len(destination_wb.sheets) + 1:
                logging.error("Invalid paste position. Must be between 1 and the number of sheets + 1.")
                return
            
            # Paste at the specified position
            source_sheet.copy(after=destination_wb.sheets[paste_position - 1])
        
        logging.info(f"Sheet copied from '{source_file_path}' to '{destination_file_path}'.")

        # Save the destination workbook
        destination_wb.save(destination_file_path)

    except Exception as e:
        logging.error(f"An error occurred: {e}")

    finally:
        # Clean up
        source_wb.close()
        destination_wb.close()
        app.quit()

if __name__ == "__main__":
    pass
