import logging

logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s | %(module)s | %(lineno)d | %(funcName)s | %(levelname)s | %(message)s',
    )

def ex_print_action(sheet, 
                     action='check',  # 'check' or 'set'
                     left_margin=0, right_margin=0, top_margin=0, bottom_margin=0,
                     header_margin=0, footer_margin=0,
                     left_header='', center_header='', right_header='',
                     left_footer='', center_footer='', right_footer='',
                     center_horizontally=False, center_vertically=False,
                     paper_size=9,  # 9 corresponds to A4
                     fit_to_pages_wide=1, fit_to_pages_tall=1):
    """
    Checks or updates the print settings for a specified Excel worksheet.

    Author: NGUYEN TIEN THANH / KNT15083
    Last Updated: 2025-01-22

    Parameters:
    - sheet: object
        The Excel worksheet object for which the print settings will be checked or updated.
    - action: str, optional
        The action to perform: 'check' to verify settings, 'set' to update them. Defaults to 'check'.
    - left_margin: float, optional
        The left margin in points. Defaults to 0.
    - right_margin: float, optional
        The right margin in points. Defaults to 0.
    - top_margin: float, optional
        The top margin in points. Defaults to 0.
    - bottom_margin: float, optional
        The bottom margin in points. Defaults to 0.
    - header_margin: float, optional
        The header margin in points. Defaults to 0.
    - footer_margin: float, optional
        The footer margin in points. Defaults to 0.
    - left_header: str, optional
        The text for the left header. Defaults to an empty string.
    - center_header: str, optional
        The text for the center header. Defaults to an empty string.
    - right_header: str, optional
        The text for the right header. Defaults to an empty string.
    - left_footer: str, optional
        The text for the left footer. Defaults to an empty string.
    - center_footer: str, optional
        The text for the center footer. Defaults to an empty string.
    - right_footer: str, optional
        The text for the right footer. Defaults to an empty string.
    - center_horizontally: bool, optional
        If True, centers the printout horizontally. Defaults to False.
    - center_vertically: bool, optional
        If True, centers the printout vertically. Defaults to False.
    - paper_size: int, optional
        The paper size (default is 9 for A4). Defaults to 9.
    - fit_to_pages_wide: int, optional
        The number of pages wide to fit the printout. Defaults to 1.
    - fit_to_pages_tall: int, optional
        The number of pages tall to fit the printout. Defaults to 1.

    Returns:
    - list
        Returns a list of error messages if any occur during the process; otherwise, returns an empty list.

    Logs:
    - Logs debug information when checking or updating print settings.
    - Logs warning messages if there are errors during the process.
    - Logs an info message when the print settings are successfully checked or updated.
    """
    error_list = []  # List to store errors
    try:
        logging.debug(f"{action.capitalize()} print settings for sheet '{sheet.name}'.")

        if action == 'check':
            # Check margins
            if sheet.api.PageSetup.LeftMargin != left_margin:
                error_list.append(f"Left margin is not {left_margin} (current: {sheet.api.PageSetup.LeftMargin})")
            if sheet.api.PageSetup.RightMargin != right_margin:
                error_list.append(f"Right margin is not {right_margin} (current: {sheet.api.PageSetup.RightMargin})")
            if sheet.api.PageSetup.TopMargin != top_margin:
                error_list.append(f"Top margin is not {top_margin} (current: {sheet.api.PageSetup.TopMargin})")
            if sheet.api.PageSetup.BottomMargin != bottom_margin:
                error_list.append(f"Bottom margin is not {bottom_margin} (current: {sheet.api.PageSetup.BottomMargin})")
            if sheet.api.PageSetup.HeaderMargin != header_margin:
                error_list.append(f"Header margin is not {header_margin} (current: {sheet.api.PageSetup.HeaderMargin})")
            if sheet.api.PageSetup.FooterMargin != footer_margin:
                error_list.append(f"Footer margin is not {footer_margin} (current: {sheet.api.PageSetup.FooterMargin})")

            # Check headers
            if sheet.api.PageSetup.LeftHeader != left_header:
                error_list.append(f"Left header is not '{left_header}' (current: '{sheet.api.PageSetup.LeftHeader}')")
            if sheet.api.PageSetup.CenterHeader != center_header:
                error_list.append(f"Center header is not '{center_header}' (current: '{sheet.api.PageSetup.CenterHeader}')")
            if sheet.api.PageSetup.RightHeader != right_header:
                error_list.append(f"Right header is not '{right_header}' (current: '{sheet.api.PageSetup.RightHeader}')")

            # Check footers
            if sheet.api.PageSetup.LeftFooter != left_footer:
                error_list.append(f"Left footer is not '{left_footer}' (current: '{sheet.api.PageSetup.LeftFooter}')")
            if sheet.api.PageSetup.CenterFooter != center_footer:
                error_list.append(f"Center footer is not '{center_footer}' (current: '{sheet.api.PageSetup.CenterFooter}')")
            if sheet.api.PageSetup.RightFooter != right_footer:
                error_list.append(f"Right footer is not '{right_footer}' (current: '{sheet.api.PageSetup.RightFooter}')")

            # Check centering
            if sheet.api.PageSetup.CenterHorizontally != center_horizontally:
                error_list.append(f"Center horizontally is not {center_horizontally}")
            if sheet.api.PageSetup.CenterVertically != center_vertically:
                error_list.append(f"Center vertically is not {center_vertically}")

            # Check paper size
            if sheet.api.PageSetup.PaperSize != paper_size:
                error_list.append(f"Paper size is not {paper_size} (current: {sheet.api.PageSetup.PaperSize})")

            # Check fit to pages
            if sheet.api.PageSetup.FitToPagesWide != fit_to_pages_wide:
                error_list.append(f"FitToPagesWide is not {fit_to_pages_wide} (current: {sheet.api.PageSetup.FitToPagesWide})")
            if sheet.api.PageSetup.FitToPagesTall != fit_to_pages_tall:
                error_list.append(f"FitToPagesTall is not {fit_to_pages_tall} (current: {sheet.api.PageSetup.FitToPagesTall})")

            if error_list:
                logging.warning(f"Print settings check completed with errors for sheet '{sheet.name}': {error_list}")
                return error_list  # Return error list if there are errors
            else:
                logging.info(f"Print settings check completed successfully for sheet '{sheet.name}'.")
                return []  # Return empty list if no errors

        elif action == 'set':
            # Set margins
            try:
                if left_margin is not None:
                    sheet.api.PageSetup.LeftMargin = left_margin
                if right_margin is not None:
                    sheet.api.PageSetup.RightMargin = right_margin
                if top_margin is not None:
                    sheet.api.PageSetup.TopMargin = top_margin
                if bottom_margin is not None:
                    sheet.api.PageSetup.BottomMargin = bottom_margin
                if header_margin is not None:
                    sheet.api.PageSetup.HeaderMargin = header_margin
                if footer_margin is not None:
                    sheet.api.PageSetup.FooterMargin = footer_margin
            except Exception as e:
                error_list.append(f"Error setting margins: {e}")

            # Set Header
            try:
                if left_header is not None:
                    sheet.api.PageSetup.LeftHeader = left_header
                if center_header is not None:
                    sheet.api.PageSetup.CenterHeader = center_header
                if right_header is not None:
                    sheet.api.PageSetup.RightHeader = right_header
            except Exception as e:
                error_list.append(f"Error setting headers: {e}")

            # Set Footer
            try:
                if left_footer is not None:
                    sheet.api.PageSetup.LeftFooter = left_footer
                if center_footer is not None:
                    sheet.api.PageSetup.CenterFooter = center_footer
                if right_footer is not None:
                    sheet.api.PageSetup.RightFooter = right_footer
            except Exception as e:
                error_list.append(f"Error setting footers: {e}")

            # Disable center on page
            try:
                if center_horizontally is not None:
                    sheet.api.PageSetup.CenterHorizontally = center_horizontally
                if center_vertically is not None:
                    sheet.api.PageSetup.CenterVertically = center_vertically
            except Exception as e:
                error_list.append(f"Error setting center on page: {e}")

            # Set Paper Size
            try:
                if paper_size is not None:
                    sheet.api.PageSetup.PaperSize = paper_size
            except Exception as e:
                error_list.append(f"Error setting paper size: {e}")

            # Fit to specified pages
            try:
                sheet.api.PageSetup.Zoom = False
                if fit_to_pages_wide is not None:
                    sheet.api.PageSetup.FitToPagesWide = fit_to_pages_wide
                if fit_to_pages_tall is not None:
                    sheet.api.PageSetup.FitToPagesTall = fit_to_pages_tall
            except Exception as e:
                error_list.append(f"Error fitting to pages: {e}")

            if error_list:
                logging.warning(f"Print settings updated with errors for sheet '{sheet.name}': {error_list}")
                return error_list  # Return error list if there are errors
            else:
                logging.info(f"Print settings updated successfully for sheet '{sheet.name}'.")
                return []  # Return empty list if no errors

    except Exception as sheet_error:
        logging.error(f"Error processing print settings for sheet '{sheet.name}': {sheet_error}")
        return [f"Error processing print settings for sheet '{sheet.name}': {sheet_error}"]
