import os
import csv
import re
import winreg
import urllib.request
import urllib.error
import zipfile
import io
import time
import socket
import http.client
import ssl
import subprocess
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge import service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from tkinter import messagebox

from variables.Messages import *

'''
This class is for the EdgeAutomation object (constructor and functions)

'''

class EdgeAutomation:
    def __init__(self):
        #this file's path
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        #root folder path
        self.dispatcher_dir = os.path.dirname(self.script_dir)
        self.root_dir = os.path.dirname(self.dispatcher_dir)
        self.driver_dir = os.path.join(self.root_dir, 'edgedriver_win64')
        self.driver_path = os.path.join(self.driver_dir, 'msedgedriver.exe')
        self.edge_options = None
        self.driver = None
        #path to CSV of logs
        self.script_logs = os.path.join(self.dispatcher_dir, 'logs','script_logs.csv')
        self.dispatch_logs = os.path.join(self.dispatcher_dir, 'logs','dispatch_logs.csv')
        self.webdriver_timeout = 20
        
        # Create logs directory if it doesn't exist
        os.makedirs(os.path.join(self.dispatcher_dir, 'logs'), exist_ok=True)
        
        # Create driver directory if it doesn't exist
        os.makedirs(self.driver_dir, exist_ok=True)

    def get_edge_version(self):
        try:
            # Access registry to get Edge version
            key_path = r"SOFTWARE\Microsoft\Edge\BLBeacon"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            value, _ = winreg.QueryValueEx(key, "version")
            winreg.CloseKey(key)
            return value
        except Exception as e:
            self.log(f"Error getting Edge version: {e}", self.script_logs)
            return None
    
    def download_edge_driver(self, version):
        """
        Downloads the Edge WebDriver matching the user's Edge browser version.
        Uses multiple download methods and retry logic for reliability.
        """
        # Extract major version
        major_version = version.split('.')[0]
        driver_version = None
        
        # Create temp directory for downloads if it doesn't exist
        temp_dir = os.path.join(self.root_dir, 'temp_downloads')
        os.makedirs(temp_dir, exist_ok=True)
        
        # Log the start of the download process
        self.log(f"Starting download process for Edge WebDriver (Edge version {version})", self.script_logs)
        
        # Method 1: Try to get the driver version from Microsoft's CDN
        try:
            # Construct download URL
            driver_version_url = f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}"
            
            # Create a context with relaxed SSL verification 
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            
            self.log(f"Attempting to get WebDriver version from {driver_version_url}", self.script_logs)
            
            # Add proper headers to avoid being blocked
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Cache-Control': 'max-age=0'
            }
            
            req = urllib.request.Request(driver_version_url)
            for key, value in headers.items():
                req.add_header(key, value)
            
            # Try with urllib first
            with urllib.request.urlopen(req, timeout=15, context=ctx) as response:
                driver_version = response.read().decode('utf-8').strip()
            
            if driver_version:
                self.log(f"Successfully retrieved WebDriver version: {driver_version}", self.script_logs)
        except Exception as e:
            self.log(f"Error getting driver version with urllib: {e}", self.script_logs)
            
            # Method 2: Try using alternative HTTP client
            try:
                conn = http.client.HTTPSConnection("msedgedriver.azureedge.net", timeout=15, context=ctx)
                conn.request("GET", f"/LATEST_RELEASE_{major_version}", headers=headers)
                response = conn.getresponse()
                if response.status == 200:
                    driver_version = response.read().decode('utf-8').strip()
                    self.log(f"Successfully retrieved WebDriver version via HTTP client: {driver_version}", self.script_logs)
                conn.close()
            except Exception as e2:
                self.log(f"Error getting driver version with HTTP client: {e2}", self.script_logs)
        
        # If we still don't have a version, use the edge version as fallback
        if not driver_version:
            driver_version = version
            self.log(f"Using Edge version as WebDriver version fallback: {driver_version}", self.script_logs)
        
        # Define the download URL for the driver
        driver_url = f"https://msedgedriver.azureedge.net/{driver_version}/edgedriver_win64.zip"
        local_zip_path = os.path.join(temp_dir, "edgedriver_win64.zip")
        
        self.log(f"Attempting to download WebDriver from {driver_url}", self.script_logs)
        
        # Method 1: Try to download with urllib
        download_success = False
        try:
            req = urllib.request.Request(driver_url)
            for key, value in headers.items():
                req.add_header(key, value)
                
            with urllib.request.urlopen(req, timeout=60, context=ctx) as response:
                data = response.read()
                with open(local_zip_path, 'wb') as f:
                    f.write(data)
            download_success = os.path.exists(local_zip_path)
            self.log("WebDriver download successful with urllib", self.script_logs)
        except Exception as e:
            self.log(f"Error downloading WebDriver with urllib: {e}", self.script_logs)
        
        # Method 2: Try to download with powershell if urllib failed
        if not download_success:
            try:
                self.log("Attempting download with PowerShell...", self.script_logs)
                command = [
                    'powershell',
                    '-Command',
                    f"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '{driver_url}' -OutFile '{local_zip_path}' -UseBasicParsing"
                ]
                process = subprocess.run(command, capture_output=True, text=True)
                if process.returncode == 0:
                    download_success = os.path.exists(local_zip_path)
                    self.log("WebDriver download successful with PowerShell", self.script_logs)
                else:
                    self.log(f"PowerShell download error: {process.stderr}", self.script_logs)
            except Exception as e:
                self.log(f"Error using PowerShell for download: {e}", self.script_logs)
        
        # Extract if download was successful
        if download_success:
            try:
                self.log(f"Extracting WebDriver to {self.driver_dir}", self.script_logs)
                with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(self.driver_dir)
                self.log("WebDriver extracted successfully", self.script_logs)
                
                # Verify the driver executable exists
                if os.path.exists(self.driver_path):
                    self.log(f"WebDriver executable verified at {self.driver_path}", self.script_logs)
                    return True
                else:
                    self.log(f"WebDriver executable not found at {self.driver_path} after extraction", self.script_logs)
            except Exception as e:
                self.log(f"Error extracting WebDriver: {e}", self.script_logs)
        
        # Final fallback: Check if we have an MSEdgeDriver in the project directory
        self.log("Checking for existing WebDriver in system PATH...", self.script_logs)
        try:
            # Try to find msedgedriver in PATH
            edge_driver_path = shutil.which("msedgedriver.exe")
            if edge_driver_path:
                self.log(f"Found existing WebDriver at {edge_driver_path}", self.script_logs)
                # Copy it to our driver directory
                shutil.copy2(edge_driver_path, self.driver_path)
                return True
        except Exception as e:
            self.log(f"Error checking for WebDriver in PATH: {e}", self.script_logs)
        
        return False

    def log(self, message, file_path):
        """
        Appends a log message to the script_log CSV file, adding a header if the file is new.

        parameters:
            message - The message to log
            file_path - The path to the CSV file where the log will be appended
        """
        file_exists = os.path.isfile(file_path)
        
        with open(file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            
            # Write header only if the file does not exist
            if not file_exists:
                writer.writerow(['Timestamp', 'Message'])
            
            # Write the log entry with the current timestamp
            writer.writerow([datetime.now(), message])

    def dispatcher_log(self, ticket_num, dispatcher, approver):
        """
        Appends a log entry to the dispatch_logs CSV file
        """
        file_exists = os.path.isfile(self.dispatch_logs)
        
        with open(self.dispatch_logs, mode='a', newline='') as file:
            writer = csv.writer(file)
            
            # Write header only if the file does not exist
            if not file_exists:
                writer.writerow(['Timestamp', 'Ticket Number', 'Dispatcher', 'Approver'])
            
            # Write the log entry with the current timestamp
            writer.writerow([datetime.now(), ticket_num, dispatcher, approver])

    def check_driver(self):
        '''
        A function to check if the driver is initialized and exists
        '''
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
            link_file = os.path.join(self.dispatcher_dir, "link.txt")
            print(f"[LOG] Looking for link file at: {link_file}")
            if os.path.exists(link_file):
                with open(link_file, "r") as file:
                    link = file.read().strip()
                    print(f"[LOG] Found driver path in link.txt: {link}")
                return link
            else:
                print("[LOG] No link.txt file found")
                return ""
        except FileNotFoundError:
            self.show_error_message(f"Error: Link file not found. {self.dispatcher_dir}")
            return ""
        except Exception as e:
            print(f"[LOG] Error reading link.txt: {e}")
            return ""

    def initialize_driver(self):
        try:
            print("[LOG] Initializing Edge WebDriver")
            edge_version = self.get_edge_version()
            if not edge_version:
                print("[LOG] Could not detect Edge version")
                self.show_error_message("Could not detect Microsoft Edge version. Please make sure Edge is installed.")
                return False
            
            print(f"[LOG] Detected Edge version: {edge_version}")
            
            driver_from_txt = self.read_link_from_file()
            if driver_from_txt and os.path.exists(driver_from_txt):
                self.driver_path = driver_from_txt
                print(f"[LOG] Using WebDriver from link.txt: {self.driver_path}")
            else:
                print(f"[LOG] Checking WebDriver at: {self.driver_path}")
            
            if not os.path.exists(self.driver_path):
                print("[LOG] WebDriver not found at specified path")
                return False
            
            self.edge_options = webdriver.EdgeOptions()
            self.edge_options.add_argument("--log-level=3")
            self.edge_options.add_argument("--silent")
            self.edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            self.edge_options.add_experimental_option("useAutomationExtension", False)
            
            # REMOVED headless options to make browser visible
            # self.edge_options.add_argument("--headless")
            # self.edge_options.add_argument("--disable-gpu")
            self.edge_options.add_argument("--window-size=1920,1080")
            self.edge_options.add_argument("--start-maximized")
            self.edge_options.add_argument("--no-sandbox")
            self.edge_options.add_argument("--disable-dev-shm-usage")
            
            print("[LOG] Creating Edge WebDriver service")
            try:
                edge_service = Service(self.driver_path)
                self.driver = webdriver.Edge(service=edge_service, options=self.edge_options)
                print("[LOG] WebDriver initialized with Service class")
            except Exception as e1:
                print(f"[LOG] Error with Service approach: {e1}")
                try:
                    # Try alternative initialization for older selenium versions
                    from selenium.webdriver.edge.service import Service as EdgeService
                    edge_service = EdgeService(self.driver_path)
                    self.driver = webdriver.Edge(service=edge_service, options=self.edge_options)
                    print("[LOG] WebDriver initialized with EdgeService")
                except Exception as e2:
                    print(f"[LOG] Error with EdgeService approach: {e2}")
                    raise
            
            self.log("Edge WebDriver initialized successfully", self.script_logs)
            return True
            
        except Exception as e:
            detailed_error = f"Error initializing Edge WebDriver: {str(e)}"
            self.log(detailed_error, self.script_logs)
            print(f"[LOG] {detailed_error}")
            self.show_error_message(detailed_error)
            return False

    def open_browser(self, link, element, condition_function):
        #check if driver exists
        self.check_driver()

        #open browser link
        self.driver.get(link)
        element = self.deal_with_element(element, condition_function)

        #log result
        if element:
            message = "Browser opened."
        else:
            message = "Browser failed to open."
        self.log(message, self.script_logs)

    #Search Presence of Elements
    def deal_with_element(self, element, condition_function):
        #checks if driver exists
        self.check_driver()
        try:
            # Use WebDriverWait to wait for the condition to be met
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
                return self.deal_with_element(element, condition_function)  # Retry for stale element
            
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
                message = f"Element '{element}' clicked."
                found_element.click()
                self.log(message, self.script_logs)
                print(message)
        except (StaleElementReferenceException) as e:
            log_message = EXCEPTION_MESSAGES[type(e)].format(element=element)
            print(log_message)
            self.log(log_message, self.script_logs)
            if isinstance(e, StaleElementReferenceException):
                return self.click_element(element, function)  # Retry for stale element
            
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
            # Locate the span element using the given locator strategy and value
            span_element = self.driver.find_element(locator_strategy, locator_value)
            
            # Retrieve the text content of the span element
            span_text = span_element.text
            
            return span_text
        except Exception as e:
            print(f"An error occurred: {e}")
            return None