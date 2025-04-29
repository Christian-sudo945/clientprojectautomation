import os
import csv
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge import service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from tkinter import Tk, messagebox
from approver.Objects.WebDriverFunctions import presence_tag, clickable_XPath
from variables.Messages import *

__all__ = ['EdgeAutomation', 'auto_dispatch']

class EdgeAutomation:
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.dispatcher_dir = os.path.dirname(self.script_dir)
        self.driver_path = os.path.join(self.dispatcher_dir, 'edgedriver_win64', 'msedgedriver.exe')
        self.edge_options = None
        self.driver = None
        self.script_logs = os.path.join(self.dispatcher_dir, 'logs','script_logs.csv')
        self.dispatch_logs = os.path.join(self.dispatcher_dir, 'logs','dispatch_logs.csv')
        self.webdriver_timeout = 20
        self.ticket_numbers = []
        self.max_tickets = 100

    def log(self, message, file_path):
        file_exists = os.path.isfile(file_path)
        with open(file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow(['Timestamp', 'Message'])
            writer.writerow([datetime.now(), message])

    def dispatcher_log(self, ticket_num, dispatcher, approver):
        file_exists = os.path.isfile(self.dispatch_logs)
        with open(self.dispatch_logs, mode='a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow(['Timestamp', 'Ticket Number', 'Dispatcher', 'Approver'])
            writer.writerow([datetime.now(), ticket_num, dispatcher, approver])

    def check_driver(self):
        if not self.driver:
            message = "No driver initialized"
            self.log(message, self.script_logs)
            print(message)
            self.show_error_message(message)
            return
        else:
            message = "Driver initialized"
            self.log(message, self.script_logs)
            print(message)

    def show_error_message(self, message):
        messagebox.showerror("Error", message)

    def read_link_from_file(self):
        try:
            with open(os.path.join(self.dispatcher_dir, "link.txt"), "r") as file:
                link = file.read().strip()
            return link
        except FileNotFoundError:
            self.show_error_message(f"Error: Link file not found. {self.dispatcher_dir}")
            return ""

    def initialize_driver(self):
        self.edge_options = webdriver.EdgeOptions()
        self.edge_options.use_chromium = True
        self.edge_options.add_argument("start-maximized")

        binary_location = self.read_link_from_file()
        if binary_location:
            self.edge_options.binary_location = binary_location
            s = service.Service(self.driver_path)
            self.driver = webdriver.Edge(service=s, options=self.edge_options)
        else:
            self.show_error_message(f"No valid binary location found in the text file. {binary_location}")
            return False
        return True

    def open_browser(self, link, element, condition_function):
        self.check_driver()
        self.driver.get(link)
        element = self.deal_with_element(element, condition_function)
        
        if element:
            message = "Browser opened."
        else:
            message = "Browser failed to open."
        self.log(message, self.script_logs)

    def deal_with_element(self, element, condition_function):
        self.check_driver()
        try:
            found_element = WebDriverWait(self.driver, self.webdriver_timeout).until(
                condition_function(element)
            )
            message = success_message.format(element=element)
            print(message)
            self.log(message, self.script_logs)
            return found_element
        except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
            log_message = EXCEPTION_MESSAGES[type(e)].format(element=element)
            print(log_message)
            self.log(log_message, self.script_logs)
            if isinstance(e, StaleElementReferenceException):
                return self.deal_with_element(element, condition_function)
            return None
            
    def get_element_attribute(self, element, type):
        return element.get_attribute(type)
    
    def get_elements_with_attribute(self, attribute_name):
        try:
            elements = self.driver.find_elements(By.XPATH, f'//*[@{attribute_name}]')
            return elements
        except Exception as e:
            print(f"An error occurred: {e}")
            return []
        
    def search_click_element(self, element, function):
        try:
            found_element = self.deal_with_element(element, function)
            if found_element:
                self.click_element(found_element)
        except (StaleElementReferenceException) as e:
            log_message = EXCEPTION_MESSAGES[type(e)].format(element=element)
            print(log_message)
            self.log(log_message, self.script_logs)
            if isinstance(e, StaleElementReferenceException):
                return self.search_click_element(element, function)
            
    def click_element(self, element):
        try:
            message = f"Element '{element}' clicked."
            element.click()
            self.log(message, self.script_logs)
            print(message)
        except Exception as e:
            message = f"Error: {str(e)}"
            self.log(message, self.script_logs)
            print(message)

    def type_in_element(self, element, text):
        element.send_keys(text)

    def perform_CTRL_Enter(self):
        ActionChains(self.driver).send_keys(Keys.CONTROL, Keys.ENTER).perform()

    def get_text_from_span(self, locator_strategy, locator_value):
        try:
            span_element = self.driver.find_element(locator_strategy, locator_value)
            span_text = span_element.text
            return span_text
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def add_ticket_number(self, ticket_num):
        if len(self.ticket_numbers) < self.max_tickets:
            self.ticket_numbers.append(ticket_num)
            return True
        return False
        
    def clear_tickets(self):
        self.ticket_numbers = []

    def get_ticket_count(self):
        return len(self.ticket_numbers)

def auto_dispatch(browser, url):
    """
    Main function to automate the dispatch process
    
    Parameters:
    browser - The EdgeAutomation instance
    url - The URL to navigate to
    """
    browser.initialize_driver()
    try:
        html_body = "html"
        browser.open_browser(url, html_body, presence_tag)
    except Exception as e:
        message = f"Error during dispatch: {str(e)}"
        browser.log(message, browser.script_logs)
        print(message)
        root = Tk()
        root.withdraw()
        messagebox.showerror("Dispatch Error", message)
        root.destroy()
    finally:
        try:
            if hasattr(browser, 'driver') and browser.driver:
                browser.driver.quit()
        except Exception as close_error:
            print(f"Error closing browser: {close_error}")

if __name__ == "__main__":
    dispatcher = EdgeAutomation()
    auto_dispatch(dispatcher, "YOUR_URL_HERE")