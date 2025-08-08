import os
import xlwings as xw
import logging

logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s | %(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_convert_to_pdf(excel_file_path, pdf_file_path=None):
    """
    Converts an Excel file to a PDF file and saves it to the specified path.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - excel_file_path: str
        The path to the Excel file that needs to be converted.
    - pdf_file_path: str, optional
        The path where the converted PDF file will be saved. If not provided, it will be created from the excel_file_path.

    Returns:
    - None
        The function does not return any value. It logs the status of the conversion operation.

    Logs:
    - Logs an error message if the input Excel file does not exist.
    - Logs an error message if the output folder does not exist.
    - Logs an info message when the conversion to PDF is completed successfully.
    - Logs an error message if an exception occurs during the conversion process.
    """
    # Check if the input Excel file exists
    if not os.path.exists(excel_file_path):
        logging.error("Excel file does not exist.")
        return

    # If pdf_file_path is not provided, create it from excel_file_path
    if pdf_file_path is None:
        pdf_file_path = os.path.splitext(excel_file_path)[0] + '.pdf'
    
    # Ensure the output folder exists
    pdf_folder = os.path.dirname(pdf_file_path)
    if not os.path.exists(pdf_folder) and pdf_folder != '':
        logging.error("Output folder does not exist.")
        return

    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        workbook = app.books.open(excel_file_path, ignore_read_only_recommended=True)
        workbook.to_pdf(pdf_file_path)
        logging.info(f"Conversion to PDF completed! File saved at: {pdf_file_path}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

    finally:
        workbook.close()
        app.quit()

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


    
    
