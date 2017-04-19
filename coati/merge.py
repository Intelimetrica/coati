from coati.win32 import copy, execute_commandbar
from coati import utils, excel, powerpoint
import time

def resources(slide, resources):
    for resource in resources:
        resource.merge(slide)
