from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from dispatcher.variables.Elements import webdriver_timeout


#Webdriver search for presence
def presence_tag(element):
    return EC.presence_of_element_located((By.TAG_NAME, element))

def presence_class_name(element):
    return EC.presence_of_element_located((By.CLASS_NAME, element))

def presence_ID(element):
    return EC.presence_of_element_located((By.ID, element))

def presence_XPath(element):
    return EC.presence_of_element_located((By.XPATH, element))

def presence_CSS_Selector(element):
    return EC.presence_of_element_located((By.CSS_SELECTOR, element))


#Webdriver check if element is clickable
def clickable_XPath(element):
    return EC.element_to_be_clickable((By.XPATH, element))

def clickable_tag(element):
    return EC.element_to_be_clickable((By.TAG_NAME, element))
    
def clickable_CSS_Selector(element):
    return EC.element_to_be_clickable((By.CSS_SELECTOR, element))
    
def clickable_ID(element):
    return EC.element_to_be_clickable((By.ID, element))
    
def clickable_class_name(element):
    return EC.element_to_be_clickable((By.CLASS_NAME, element))

def clickable_link_text(element):
    return EC.element_to_be_clickable((By.LINK_TEXT, element))

#Switches Frame
def switch_frame_id(id):
    return EC.frame_to_be_available_and_switch_to_it((By.ID, id))

#Window Functions
def detect_new_window(window_handles):
    return EC.new_window_is_opened(window_handles)