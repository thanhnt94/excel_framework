import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

def ex_close_workbook(app, wb, save_on_close=True):
    """
    Closes an Excel workbook and optionally saves changes before closing.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-23

    Parameters:
    - app: object
        The Excel application object that is managing the workbook.
    - wb: object
        The workbook object to be closed.
    - save_on_close: bool, optional
        If True, saves changes to the workbook before closing. Defaults to True.

    Returns:
    - None
        The function does not return any value. It logs the status of the close operation.

    Logs:
    - Logs debug information when starting to close the workbook.
    - Logs an info message when the workbook is saved successfully.
    - Logs an error message if saving the workbook fails.
    - Logs debug information when the workbook is closed successfully.
    - Logs an error message if closing the workbook fails.
    - Logs debug information about whether the Excel application is quitting or remaining open.
    """
    
    logging.debug(f"Starting to close the workbook '{wb.name}' with save_on_close set to {save_on_close}.")
    wb_name = wb.name  

    try:
        if save_on_close:
            try:
                wb.save()
                logging.debug(f"Successfully saved the workbook '{wb_name}'.")
            except Exception as save_error:
                logging.error(f"Failed to save the workbook '{wb_name}'. Error: {save_error}")
                raise
        else:
            logging.debug(f"Skipped saving the workbook '{wb_name}'.")

        wb.close()
        logging.debug(f"Successfully closed the workbook '{wb_name}'.")

    except Exception as e:
        logging.error(f"Failed to close the workbook '{wb_name}'. Error: {e}")
        raise

    finally:
        if app.books.count == 0:
            app.quit()
            logging.debug('Successfully quit the Excel application.')
        else:
            logging.debug('The Excel application is still running with other workbooks open.')

if __name__ == '__main__':
    try:
        file_path = r"C:\Users\KNT15083\Documents\Book1111.xlsx"
        file_path = r"C:\Users\KNT15083\Downloads\fffsda.xlsx"
        file_path = r"C:\Users\KNT15083\Documents\Database1.xlsx"
        file_path = r"C:\Users\KNT15083\Downloads\Q3設通リスト_LT_柏原-3殿宛.xlsx"

    except Exception as e:
        pass

