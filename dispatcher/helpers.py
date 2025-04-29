from .auto_dispatch import auto_dispatch, EdgeAutomation

def safe_dispatch(browser, url, *args):
    """
    A wrapper for auto_dispatch that handles various argument counts.
    
    This function ensures that auto_dispatch is called correctly regardless of
    how many arguments are provided.
    """
    # Ensure we have the right type of browser
    if not isinstance(browser, EdgeAutomation):
        raise TypeError("Browser must be an instance of EdgeAutomation")
    
    # Call auto_dispatch with only the required arguments
    return auto_dispatch(browser, url)
