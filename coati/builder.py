"""Functionality for handling the logic
of traversing and working with a project.
"""

import os
import imp
import codecs
from copy import deepcopy
from officereports import excel, merge, powerpoint
from officereports.resources import factory


def loadcode(path, name):
    module = imp.new_module(name)
    code = codecs.open(path, 'r', 'utf-8').read()
    exec code in module.__dict__
    return module


def filllabels(resources, funmodule):
    new_resources = deepcopy(resources)

    def callfn(label):
        return getattr(funmodule, label)()

    new_resources['labels'] = [(label, callfn(label)) for label
                               in resources.get('labels', [])]

    return new_resources


class SlideBuilder(object):
    """Manage the logic for extracting
    information for the current configuration
    of a given slide and build it into a document instance"""

    def __init__(self, path):
        self.path = path

    @property
    def pptx(self):
        if not hasattr(self, '_pptx'):
            self._pptx = self.template.pptx_from_template()
        return self._pptx

    @property
    def template(self):
        if not hasattr(self, '_template'):
            self._template = powerpoint.SlideTemplate(
                os.path.join(self.path, 'slide.pptx'))
        return self._template

    def loadtemplate(self):
        if not hasattr(self, '_pptx'):
            self._pptx = self.template.pptx_from_template()

    def _attemptload(self, fname):
        path = os.path.join(self.path, fname)
        if not os.path.isfile(path):
            return None

        module_name = fname.split('.')[0]
        return loadcode(path, module_name)

    def resourcedict(self):
        resources_module = self._attemptload('resources.py')
        functions_module = self._attemptload('functions.py')

        if not resources_module:
            return None

        resource_dict = resources_module.resources

        if functions_module:
            return filllabels(resource_dict, functions_module)
        else:
            return resource_dict

    def fillexcel(self):
        xlsx_src = os.path.join(self.path, 'slide.xlsx')

        if not os.path.isfile(xlsx_src):
            return None

        self.excelapp = excel.runexcel()
        self.xlsx = excel.open_xlsx(self.excelapp, xlsx_src)

        fillmodule = self._attemptload('fill.py')

        if fillmodule:
            fillmodule.fill(self.xlsx)

    def closexcel(self):
        if hasattr(self, 'xlsx'):
            self.excelapp.CutCopyMode = False
            self.xlsx.Close()

    def mergeresources(self, resources):
        if not hasattr(self, 'xlsx'):
            return None

        slide = powerpoint.slide(self.pptx, 0)

        if len(resources) > 0:
            sheet = excel.sheet(self.xlsx, 0)
            objs = factory(resources, sheet)
            merge.resources(slide, objs)

    def build(self):
        self.fillexcel()
        resources = self.resourcedict() or dict()
        self.loadtemplate()
        self.mergeresources(resources)
        self.closexcel()
