import re

def extract_and_log_dispatch_info(self, ticket_element, dispatcher_element, approver_element):
    try:
        ticket_text = ticket_element.text.strip() if ticket_element and ticket_element.text else "Unknown Ticket"
        ticket_match = re.search(r'(\d{10})', ticket_text)
        ticket_number = ticket_match.group(1) if ticket_match else ticket_text
    
        dispatcher_name = dispatcher_element.text.strip() if dispatcher_element and dispatcher_element.text else "Unknown"

        approver_name = approver_element.text.strip() if approver_element and approver_element.text else "Unknown"
        
        # Log information using dispatcher_log instead of possibly calling auto_dispatch
        self.browser.dispatcher_log(ticket_number, dispatcher_name, approver_name)
        
        return ticket_number, dispatcher_name, approver_name
    except Exception as e:
        self.browser.log(f"Error extracting dispatch information: {str(e)}", self.browser.script_logs)
        return "Unknown Ticket", "Unknown", "Unknown"