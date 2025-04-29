webdriver_timeout = 20

#Tags
html_body = "body"
iframe_tag = "iframe"

#IDs
ticket_window = "content_frame"
ticket_id = "WD0109"
submit_approve = "WD027B"
submit_window = "content_frame"

#CSS Selectors

#XPaths
scroll_handle = '//*[@acf="Hndl"]'
scroll_track = '//*[@acf="PNext"]'
approver_ticket = '//*[contains(text(), "Approval required for access role request")]/ancestor::tr[@rr="{rr_num}"]//span[contains(text(), "Approval required for access role request") and not(@id="WD01AF")]'
submit_button = '//div[@title="Submit" and contains(@class, "lsButton") and contains(@class, "lsButton--design-emphasized")]'
ticket_title_xpath = '//*[@class="lsPageHeaderTitle--text-overflow"]'
work_inbox = '//a[.//span[text()="Work Inbox"] and contains(@class, "lsLink")]'
ticket_table = '//tbody[contains(@id, "-contentTBody")]/tr[@rr]'
