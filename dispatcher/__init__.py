# Initialize the dispatcher package
from .auto_dispatch import EdgeAutomation, auto_dispatch
from .helpers import safe_dispatch

__all__ = ['EdgeAutomation', 'auto_dispatch', 'safe_dispatch']
