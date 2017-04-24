"""General functionality for the win32 api."""
import win32com.client as w32
from ctypes import windll
import time


def run(name):
    app = w32.Dispatch(name)
    app.Visible = True
    return app


def copy(obj):
    windll.user32.EmptyClipboard()
    obj.Activate()
    obj.Copy()


def execute_commandbar(element, command):
    element.Application.CommandBars.ExecuteMso(command)
    time.sleep(0.1)
