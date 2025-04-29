from approver.Objects.WebDriverFunctions import clickable_ID, clickable_XPath, presence_ID, presence_XPath, presence_tag, switch_frame_id
from approver.variables.Elements import *
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import tkinter as tk
from tkinter import messagebox
from dispatcher.variables.Elements import work_inbox

with_scroll = True

def auto_approve(browser, url, run_counter=1):
    '''
    This is the main block of code that would be executed to automatically approve tickets.
    Note: most, if not all, locator values (element IDs, XPaths, etc) are found in variables>Elements.py
    
    *** If possible, pleaese change the locator values in variables>Elements.py instead! ***
    parameters:
        browser - this is the EdgeAutomation object that will be used frequenty
        url - this is a string for the url to the SAP Cup Tool web application
    '''
    browser.initialize_driver()
    try:
        global with_scroll
        global wait
        global original_window
        
        # Increase timeout to give more time for elements to load
        wait = WebDriverWait(browser.driver, webdriver_timeout)
        
        # Add explicit check to ensure browser is initialized
        if browser.driver is None:
            raise Exception("Browser driver not initialized properly")
            
        browser.open_browser(url, html_body, presence_tag)
        
        # Add a small delay to ensure page is fully loaded
        time.sleep(2)
        
        iframe = browser.deal_with_element(iframe_tag, presence_tag)
        iframe_id = browser.get_element_attribute(iframe, 'id')
        browser.deal_with_element(iframe_id, switch_frame_id)
        
        # Add explicit wait for work inbox button
        work_inbox_button = wait.until(presence_XPath(work_inbox))
        '''The work inbox button ID may differ across users. 
        When I made this, my work inbox button ID was "WD58" while sir's was "WD5C". Please take note of this as this is very important!
        '''
        
        # Store the original window handle before clicking anything
        original_window = browser.driver.current_window_handle
        
        # Click the work inbox button
        print("Clicking work inbox button...")
        work_inbox_button.click()
        print("Work Inbox button clicked.")
        
        # Wait longer for the new window to open - this is a critical point where timeouts occur
        time.sleep(5)
        
        # IMPORTANT: Different approach for handling new windows - don't rely on browser.deal_with_element method
        # Instead, explicitly check and wait for new windows
        
        # First, check current window handles after button click
        current_handles = browser.driver.window_handles
        print(f"Current window handles: {len(current_handles)}")
        
        # If only one window is present, we need to wait longer or try clicking again
        if len(current_handles) < 2:
            print("No new window detected. Trying again...")
            
            # Try clicking the button again
            try:
                work_inbox_button = wait.until(presence_XPath(work_inbox))
                work_inbox_button.click()
                print("Work Inbox button clicked a second time.")
                time.sleep(5)  # Wait again
                
                # Check window handles again
                current_handles = browser.driver.window_handles
                print(f"Window handles after second attempt: {len(current_handles)}")
                
                if len(current_handles) < 2:
                    # Try to navigate to the inbox URL directly if available
                    # If you have a direct URL to the inbox, you can add it here
                    print("Still no new window. Trying direct navigation if possible...")
            except Exception as e:
                print(f"Error clicking work inbox button again: {e}")
        
        # Safety check - if we still don't have a new window, we need to stop
        if len(browser.driver.window_handles) < 2:
            raise Exception("Failed to open ticket window after multiple attempts")
        
        # Now we know we have at least 2 windows
        windows = browser.driver.window_handles
        
        # Find the new window and switch to it
        new_window = None
        for window in windows:
            if window != original_window:
                new_window = window
                # Try-except block to handle potential session issues
                try:
                    browser.driver.switch_to.window(window)
                    print("Switched to new window.")
                    break
                except Exception as e:
                    print(f"Error switching to window: {e}")
                    # If we failed to switch, try to recover by focusing on original window
                    browser.driver.switch_to.window(original_window)
                    raise Exception("Failed to switch to new window")
        
        if new_window is None:
            raise Exception("Failed to identify the new window")
            
        # After switching to the new window, wait for its content to load
        time.sleep(2)
        
        try:
            browser.deal_with_element(ticket_window, switch_frame_id)
        except Exception as e:
            print(f"Error switching to frame in ticket window: {e}")
            # Try refreshing the page if frame switching fails
            browser.driver.refresh()
            time.sleep(3)
            browser.deal_with_element(ticket_window, switch_frame_id)
            
        # Now proceed with the rest of the automation
        drag_scrollbar_to_bottom(browser)
        time.sleep(1)
        
        last_ticket_rr = get_last_rr_in_tbody(browser)
        
        # Validate we got a valid ticket number
        if last_ticket_rr is None:
            print("No tickets found in the inbox. Ending automation cycle.")
            return
            
        last_ticket_rr = int(last_ticket_rr)
        
        scroll_up_indc = 0
        while last_ticket_rr > 0:
            if with_scroll:
                print(f"Processing ticket {last_ticket_rr}.")
                
            try:
                dispatch_ticket(browser, last_ticket_rr)
            except Exception as e:
                print(f"Failed to dispatch ticket {last_ticket_rr}: {e}")
                
                # Add session recovery logic
                if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
                    print("Session error detected. Attempting to recover...")
                    
                    # Check which windows are still available
                    try:
                        available_windows = browser.driver.window_handles
                        if available_windows:
                            # Switch to any available window
                            browser.driver.switch_to.window(available_windows[0])
                            print(f"Recovered by switching to window: {available_windows[0]}")
                        else:
                            # If no windows are available, we need to restart the browser
                            print("No windows available. Restarting browser session...")
                            browser.initialize_driver()
                            browser.open_browser(url, html_body, presence_tag)
                            # We should probably exit this cycle and restart from the beginning
                            break
                    except Exception as recovery_error:
                        print(f"Recovery attempt failed: {recovery_error}")
                        # If recovery failed, restart the browser
                        browser.initialize_driver()
                        browser.open_browser(url, html_body, presence_tag)
                        break
                        
            last_ticket_rr -= 1
            scroll_up_indc += 1
            if scroll_up_indc % 10 == 0 and scroll_up_indc > 0:
                scroll_up(browser)
            print(f"{scroll_up_indc} ticket/s processed.")
            
        while run_counter < 5:
            # Check if session is still valid before proceeding
            try:
                # Quick check if driver is still responsive
                browser.driver.current_url
                
                time.sleep(3)
                refresh_page(browser)
                print(f"Finished cycle {run_counter}. Refreshing and restarting process.")
                run_counter += 1
                time.sleep(5)
            except Exception as e:
                if "invalid session id" in str(e).lower():
                    print("Session expired during refresh cycle. Restarting browser...")
                    browser.initialize_driver()
                    browser.open_browser(url, html_body, presence_tag)
                    # Reset counter since we're essentially starting over
                    run_counter = 1
                else:
                    raise e
                    
    except Exception as e:
        message = f"Error during automation: {str(e)}"
        browser.log(message, browser.script_logs)
        print(message)
        
        # Display error message in a popup
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showerror("Automation Error", message)
        root.destroy()
        
    finally:
        # Ensure we properly close the browser to prevent orphaned processes
        try:
            if hasattr(browser, 'driver') and browser.driver is not None:
                browser.driver.quit()
        except Exception as close_error:
            print(f"Error closing browser: {close_error}")
            
        print("Finished automation")


def refresh_page(browser):
    '''
    Refresh the page by clicking on the "WDF9-text" button.
    '''
    try:
        # First check if session is still valid
        try:
            browser.driver.current_url  # This will fail if session is invalid
        except Exception as e:
            if "invalid session id" in str(e).lower():
                raise Exception("Browser session is no longer valid. Cannot refresh page.")
        
        refresh_button = wait.until(presence_ID("WDF9-text"))
        refresh_button.click()
        time.sleep(5)
        print("Page refreshed by clicking on WDF9-text button.")
        drag_scrollbar_to_bottom(browser)
        time.sleep(0.5)
        last_ticket_rr = get_last_rr_in_tbody(browser)
        
        if last_ticket_rr is None:
            print("No tickets found.")
            return
            
        last_ticket_rr = int(last_ticket_rr)
        scroll_up_indc = 0
        
        while last_ticket_rr > 0:
            if with_scroll:
                print(f"Processing ticket {last_ticket_rr}.")
                
            try:
                dispatch_ticket(browser, last_ticket_rr)
            except Exception as e:
                print(f"Failed to dispatch ticket {last_ticket_rr}: {e}")
                
                # Add session recovery logic
                if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
                    print("Session error detected during refresh cycle. Breaking out...")
                    break
                    
            last_ticket_rr -= 1
            scroll_up_indc += 1  
            
            # Check if we need to scroll
            if scroll_up_indc % 10 == 0 and scroll_up_indc > 0:  
                scroll_up(browser) 
            print(f"{scroll_up_indc} ticket/s processed.")
            
    except Exception as e:
        print(f"Failed to refresh the page: {e}")


def drag_scrollbar_to_bottom(browser):
    '''
    Drags the scrollbar to the bottom and ensures it stays there.
    '''
    global with_scroll
    try:
        # First check if session is still valid
        try:
            browser.driver.current_url  # This will fail if session is invalid
        except Exception as e:
            if "invalid session id" in str(e).lower():
                raise Exception("Browser session is no longer valid. Cannot drag scrollbar.")
        
        for _ in range(3):
            scrollbar_handle = browser.deal_with_element(scroll_handle, clickable_XPath)
            scrollV_bar = browser.deal_with_element(scroll_track, clickable_XPath)
            
            handle_height = scrollbar_handle.size['height']
            bar_height = scrollV_bar.size['height']
            max_scroll_distance = bar_height - handle_height
            
            action = ActionChains(browser.driver)
            action.click_and_hold(scrollbar_handle).move_by_offset(0, max_scroll_distance).release().perform()
            time.sleep(0.5)
            
            handle_position = scrollbar_handle.location['y']
            track_position = scrollV_bar.location['y'] + bar_height - handle_height
            
            if handle_position >= track_position - 5:
                print("Scrollbar successfully dragged to the bottom.")
                break
        else:
            print("Scrollbar failed to stay at the bottom after multiple attempts.")
            with_scroll = False
    except Exception as e:
        print(f"An error occurred during drag operation: {e}")
        with_scroll = False


def scroll_up(browser):
    '''
    Clicks the empty space above the scrollbar to scroll up slightly.
    '''
    try:
        # First check if session is still valid
        try:
            browser.driver.current_url  # This will fail if session is invalid
        except Exception as e:
            if "invalid session id" in str(e).lower():
                raise Exception("Browser session is no longer valid. Cannot scroll up.")
                
        browser.search_click_element("WD72-scrollV-spc", clickable_ID)
        time.sleep(0.5)
        print("Clicked on the scroll space to move the scrollbar up.")
    except Exception as e:
        print(f"Error clicking scroll space: {e}")


def get_last_rr_in_tbody(browser):
    '''
    This function gets the last visible rr value from the table body, dynamically adjusting the scroll if necessary.
    '''
    try:
        # First check if session is still valid
        try:
            browser.driver.current_url  # This will fail if session is invalid
        except Exception as e:
            if "invalid session id" in str(e).lower():
                raise Exception("Browser session is no longer valid. Cannot get RR values.")
                
        wait.until(presence_XPath(ticket_table))
        ''' The tbody @id xpath may differ across users. 
        When I made this, my xpath was "WD72" while sir's was "WD73". Please take note of this as this is very important! 
        '''
        print("Attempting to retrieve all elements with 'rr' attribute.")
        elements = browser.get_elements_with_attribute('rr')
        
        if elements:
            visible_rr_values = [
                int(browser.get_element_attribute(e, 'rr'))
                for e in elements if browser.get_element_attribute(e, 'rr')]
                
            if visible_rr_values:
                last_visible_rr = max(visible_rr_values)
                print(f"Found last visible rr: {last_visible_rr}")
                return last_visible_rr
            else:
                print("No visible rr values found.")
                return None
        
        else:
            print("No elements with 'rr' attribute found.")
            return None
        
    except Exception as e:
        print(f"Error during rr retrieval: {e}")
        return None


def dispatch_ticket(browser, rr_num):
    '''
    This is the main function with 2 main goals:
        1. Approve a single ticket
        2. Log the ticket number
    Parameters:
        1. Browser - this is the EdgeAutomation object that will be used frequently
        2. rr_num - this is the rr value that will be used by selenium to identify the ticket being forwarded
        3. retries - this is just the number of retry runs allowed for this function
    '''
    try:
        # First check if session is still valid
        try:
            browser.driver.current_url  # This will fail if session is invalid
        except Exception as e:
            if "invalid session id" in str(e).lower():
                raise Exception("Browser session is no longer valid. Cannot dispatch ticket.")
                
        ticket_window = browser.driver.current_window_handle
        
        # Click on the ticket to open it
        print(f"Clicking on ticket with rr={rr_num}")
        browser.search_click_element(approver_ticket.format(rr_num=rr_num), clickable_XPath)
        
        # Use a more generous wait time for the new window to open
        time.sleep(3)
        
        # Get all window handles AFTER clicking
        window_handles = browser.driver.window_handles
        print(f"Window handles after clicking ticket: {len(window_handles)}")
        
        # If we don't have the expected number of windows, try a different approach
        if len(window_handles) < 3:  # We expect at least 3 windows at this point
            print("Did not get expected number of windows. Trying alternative approach...")
            
            # Try clicking the ticket again
            try:
                browser.search_click_element(approver_ticket.format(rr_num=rr_num), clickable_XPath)
                time.sleep(3)
                window_handles = browser.driver.window_handles
                print(f"Window handles after second attempt: {len(window_handles)}")
            except Exception as click_error:
                print(f"Error clicking ticket again: {click_error}")
        
        # If we still don't have enough windows, we can't proceed with this ticket
        if len(window_handles) < 3:
            raise Exception(f"Failed to open submit window for ticket {rr_num} after multiple attempts")
            
        # Find the submit window - it should be neither the ticket window nor the original window
        submit_window_handle = None
        for handle in window_handles:
            if handle != ticket_window and handle != original_window:
                submit_window_handle = handle
                
                # Try switching to the submit window
                try:
                    browser.driver.switch_to.window(handle)
                    print("Switched to the submit window.")
                    break
                except Exception as switch_error:
                    print(f"Error switching to submit window: {switch_error}")
                    continue
                
        # If we couldn't find or switch to the submit window, we can't proceed
        if submit_window_handle is None or browser.driver.current_window_handle != submit_window_handle:
            raise Exception("Could not switch to submit window")
            
        # Wait for the submit window to load
        time.sleep(2)
        
        # Now try to interact with the frame in the submit window
        try:
            browser.deal_with_element(submit_window, switch_frame_id)
        except Exception as frame_error:
            print(f"Error switching to frame in submit window: {frame_error}")
            
            # Try refreshing and trying again
            browser.driver.refresh()
            time.sleep(2)
            browser.deal_with_element(submit_window, switch_frame_id)
                                              
        retries = 0
        
        while retries < 3:  # Increased max retries
            try:
                # Different approach to find the submit button - try multiple locator strategies
                try:
                    # Try the XPath approach first
                    submit_button = wait.until(clickable_XPath('//div[@title="Submit" and contains(@class, "lsButton") and contains(@class, "lsButton--design-emphasized")]'))
                except Exception as xpath_error:
                    print(f"Could not find submit button with XPath: {xpath_error}")
                    
                    # If XPath fails, try ID approach if you have one
                    try:
                        submit_button = wait.until(presence_ID("WD1F"))  # Or WD20 depending on user
                        print("Found submit button using ID")
                    except:
                        # If both approaches fail, try a more general approach
                        submit_button = wait.until(clickable_XPath('//div[contains(@title, "Submit")]'))
                        print("Found submit button using general approach")
                
                # Now click the button
                submit_button.click()
                print("Submit button clicked.")
                break
            except Exception as e:
                retries += 1
                print(f"Attempt {retries} failed to click submit button: {e}")
                time.sleep(2)  # Increased wait time
                
        # If we failed after all retries, try to recover
        if retries >= 3:
            print("Failed to click submit button after multiple attempts, attempting recovery...")
            # Try to find submit button with alternative methods if available
            try:
                # Try JavaScript click if available
                if hasattr(browser, 'js_click') and callable(getattr(browser, 'js_click')):
                    submit_elements = browser.driver.find_elements_by_xpath('//div[contains(@title, "Submit")]')
                    if submit_elements:
                        browser.js_click(submit_elements[0])
                        print("Used JavaScript click as recovery method")
            except Exception as js_error:
                print(f"JavaScript click recovery failed: {js_error}")
            
        # Wait after submit button click
        time.sleep(2)
        
        # Use try-except for CTRL+Enter in case of issues
        try:
            browser.perform_CTRL_Enter()
            print("Performed CTRL+Enter")
        except Exception as e:
            print(f"Error performing CTRL+Enter: {e}")
            # Try alternative method if available
            try:
                # Try to find and click "Yes" button directly if it exists
                yes_button = browser.driver.find_element_by_xpath('//button[contains(text(), "Yes") or @value="Yes"]')
                yes_button.click()
                print("Clicked Yes button directly")
            except Exception as yes_error:
                print(f"Could not find Yes button: {yes_error}")
            
        # Wait longer after confirmation
        time.sleep(3)
        
        # Get ticket number with error handling
        try:
            ticket_number = get_ticket_number(browser)
            print("Ticket number = ", ticket_number)
            if ticket_number:
                browser.log(ticket_number, browser.approve_logs)
            else:
                print("Warning: Could not retrieve ticket number")
        except Exception as e:
            print(f"Error getting ticket number: {e}")
       
        # Close the submit window with error handling
        try:
            browser.driver.close()
            print("Submit window closed successfully")
        except Exception as e:
            print(f"Error closing submit window: {e}")
            
        # Switch back to ticket window with error handling
        try:
            browser.driver.switch_to.window(ticket_window)
            print("Switched back to ticket window.")
            
            # Check if we can interact with the iframe
            try:
                iframe = browser.deal_with_element(iframe_tag, presence_tag)
                iframe_id = browser.get_element_attribute(iframe, 'id')
                browser.deal_with_element(iframe_id, switch_frame_id)
            except Exception as iframe_error:
                print(f"Error interacting with iframe: {iframe_error}")
                # Try to refresh the page or recover in some way
                browser.driver.refresh()
                time.sleep(2)
                # Re-try to interact with iframe
                iframe = browser.deal_with_element(iframe_tag, presence_tag)
                iframe_id = browser.get_element_attribute(iframe, 'id')
                browser.deal_with_element(iframe_id, switch_frame_id)
                
        except Exception as e:
            print(f"Error switching back to ticket window: {e}")
            
            # Try to recover by finding the ticket window in available handles
            try:
                available_handles = browser.driver.window_handles
                for handle in available_handles:
                    if handle == ticket_window:
                        browser.driver.switch_to.window(handle)
                        print("Recovered ticket window")
                        break
                else:
                    # If ticket window not found, switch to original window if available
                    if original_window in available_handles:
                        browser.driver.switch_to.window(original_window)
                        print("Could not find ticket window, switched to original window")
                    else:
                        print("Could not find any valid window to switch to")
            except Exception as recovery_error:
                print(f"Recovery attempt failed: {recovery_error}")
                  
    except Exception as e:
        print(f"Error during ticket approval: {e}")
        
        # Session recovery attempt
        if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
            raise Exception(f"Session error during dispatch of ticket {rr_num}: {e}")


def get_ticket_number(browser):
    try:
        element = browser.deal_with_element(ticket_title_xpath, presence_XPath)
        text_content = element.text
        match = re.search(r"\b\d{6}\b", text_content)
        if match:
            return match.group(0)
        else:
            print("Access request number not found.")
            return None
    except Exception as e:
        print(f"Error getting ticket number: {e}")
        return None