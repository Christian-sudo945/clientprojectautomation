from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

EXCEPTION_MESSAGES = {
    TimeoutException: "Timeout occurred while waiting for {element} to load.",
    NoSuchElementException: "No {element} element found.",
    StaleElementReferenceException: "The element {element} is stale."
}

success_message = "Function executed successfully for element {element}."