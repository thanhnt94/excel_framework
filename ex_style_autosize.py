import xlwings as xw

def ex_style_autosize(sheet, index, fit_type='column'):
    """
    Adjusts the width of a column or the height of a row in an Excel worksheet.

    Parameters:
    - sheet: object
        The Excel worksheet object for which the fit operation will be performed.
    - index: int or str
        The index (1-based) or letter of the column or the index (1-based) of the row to fit.
    - fit_type: str, optional
        The type of fit operation: 'column' to fit the column width, 'row' to fit the row height. Defaults to 'column'.

    Returns:
    - None
    """
    # If index is an integer, convert it to the corresponding letter for columns
    if fit_type == 'column' and isinstance(index, int):
        index = xw.utils.get_column_letter(index)

    # Fit column width
    if fit_type == 'column':
        sheet.range(index + ':' + index).autofit()
    
    # Fit row height
    elif fit_type == 'row' and isinstance(index, int):
        sheet.range(index).autofit()

    else:
        raise ValueError("Invalid fit_type. Use 'column' or 'row'.")
