from officereports.win32 import copy, execute_commandbar
from officereports import utils, excel, powerpoint
import time

def resources(slide, resources):
    for resource in resources:
        resource.merge(slide)
