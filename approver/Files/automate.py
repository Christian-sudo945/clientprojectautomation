import re


def extract_and_log_ticket_number(self, ticket_element):
        try:
            ticket_text = ticket_element.text.strip() if ticket_element and ticket_element.text else "Unknown Ticket"
            
            # Extract ticket number using regex pattern
            ticket_match = re.search(r'(\d{10})', ticket_text)
            ticket_number = ticket_match.group(1) if ticket_match else ticket_text
            
            # Log the ticket number using the updated ticket_log method
            self.browser.ticket_log(ticket_number, self.browser.approve_logs)
            
            return ticket_number
        except Exception as e:
            self.browser.log(f"Error extracting ticket number: {str(e)}", self.browser.script_logs)
            return "Unknown Ticket"