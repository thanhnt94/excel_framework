import psutil
import pygetwindow as gw
import win32process
import logging


logging.basicConfig(
    level=logging.DEBUG,
    format='%(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
)

def ex_close_hidden():
    """
    Terminates any hidden instances of Excel processes running on the system.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Returns:
    - None
        The function does not return any value. It logs the status of the termination process.
        
    Logs:
    - Logs debug information for each Excel process checked and whether it was terminated or not.
    - Logs errors if there are issues accessing any processes.
    - Logs an info message if no hidden Excel processes were found to terminate.
    """
        
    logging.debug("Starting the process to terminate hidden Excel instances.")
    
    def is_excel_window_visible(pid):
        windows = gw.getWindowsWithTitle('Excel')
        for window in windows:
            if window._hWnd:
                hwnd = window._hWnd
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid == pid:
                    return True
        return False

    any_terminated = False

    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'EXCEL.EXE':
            try:
                if not is_excel_window_visible(proc.info['pid']):
                    proc.terminate()
                    logging.info(f'Terminated hidden Excel process: {proc.info["name"]} (PID: {proc.info["pid"]})')
                    any_terminated = True
                else:
                    logging.debug(f'Excel process (PID: {proc.info["pid"]}) is visible and will not be terminated.')
            except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                logging.error(f'Error accessing process (PID: {proc.info["pid"]}): {e}')

    if not any_terminated:
        logging.info('No hidden Excel processes were found to terminate.')

if __name__ == '__main__':
    ex_close_hidden()

