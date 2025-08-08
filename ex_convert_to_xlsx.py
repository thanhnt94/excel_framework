import os
import xlwings as xw
import logging

logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s | %(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_convert_to_xlsx(xls_file_path, xlsx_file_path=None, remove_xls_file=False):
    """
    Converts an Excel file from .xls format to .xlsx format and optionally deletes the original .xls file.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - xls_file_path: str
        The path to the .xls file that needs to be converted.
    - xlsx_file_path: str, optional
        The path where the converted .xlsx file will be saved. If not provided, it will be created from the xls_file_path.
    - remove_xls_file: bool, optional
        If True, the original .xls file will be deleted after conversion. Defaults to False.

    Returns:
    - None
        The function does not return any value. It logs the status of the conversion operation.

    Logs:
    - Logs an info message when the conversion is successful.
    - Logs an info message if the original .xls file is deleted after conversion.
    - Logs an error message if an exception occurs during the conversion process.
    """
    try:
        # If xlsx_file_path is not provided, create it from xls_file_path
        if xlsx_file_path is None:
            xlsx_file_path = os.path.splitext(xls_file_path)[0] + '.xlsx'
        
        app = xw.App(visible=False) 
        app.display_alerts = False
        wb = app.books.open(xls_file_path)

        wb.save(xlsx_file_path)
        logging.info(f"Successfully converted {xls_file_path} to {xlsx_file_path}.")

        wb.close()
        app.quit()

        # Remove the original xls file if remove_xls_file is True
        if remove_xls_file:
            os.remove(xls_file_path)
            logging.info(f"Original file {xls_file_path} has been deleted.")

    except Exception as e:
        logging.error(f"Error during conversion: {e}")

if __name__ == "__main__":
    try:
        excel_file = r"C:\Users\KNT15083\Downloads\sontung\Out\RN01815\検討書\J2-24-P068.xlsx"
        output_folder = r"C:\Users\KNT15083\Downloads\sontung\Out\RN01815\検討書"
        file_name = r"\hahaha\kakakak"

        # result, info = excel_check_file_status(excel_file, info=True)
        # if result:
        #     print(info)
        #     app, wb = xw_open_workbook(excel_file)
        #     excel_to_pdf(excel_file, output_folder, output_file_name=file_name)
        #     xw_close_workbook(app, wb)
        # else:
        #     print(info)
    except Exception as e:
        pass


    
    
