import logging
import subprocess


logging.basicConfig(
    level=logging.DEBUG,
    format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

def ex_close_all():
    """
    Force closes all running instances of Excel on the system.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Returns:
    - None
        The function does not return any value. It logs the status of the force close operation.

    Logs:
    - Logs debug information when starting the force close process.
    - Logs an info message if Excel is successfully force closed.
    - Logs an error message if an error occurs during the process.
    - Logs a warning if no running Excel application is found to close.
    """
    logging.debug("Starting the process to force close Excel.")
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], check=True)
        logging.info("Excel has been force closed successfully.")
    except subprocess.CalledProcessError as e:
        logging.error(f"An error occurred while trying to close Excel: {e}")
    except FileNotFoundError:
        logging.warning("No running Excel application found to close.")

if __name__ == '__main__':
    ex_close_all()

