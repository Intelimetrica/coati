"""Functionality for handling the logic
of traversing and working with a project.
"""

import os
import imp
import codecs
import logging
from copy import deepcopy
from coati import excel, merge, powerpoint
from coati.resources import factory
from coati.utils import flatten

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

    def __init__(self, template_pptx, path):
        self.template_pptx = template_pptx
        self.path = path

    @property
    def pptx(self):
        if not hasattr(self, '_pptx'):
            self._pptx = self.template.pptx_from_template()
        return self._pptx

    @property
    def template(self):
        if not hasattr(self, '_template'):
            self._template = powerpoint.SlideTemplate(self.template_pptx)
        return self._template

    def loadtemplate(self):
        if not hasattr(self, '_pptx'):
            self._pptx = self.template.pptx_from_template()

    def loadconfig(self, fname):
        config = self._attemptload(fname)
        if not config or not config.slides or type(config.slides) is not list:
            self.slidesdata = []
            logging.warning("config file no valid")
            return None
        self.slidesdata = flatten(config.slides)
        return config

    def _attemptload(self, fname):
        path = os.path.join(self.path, fname)
        if not os.path.isfile(path):
            return None

        module_name = fname.split('.')[0]
        return loadcode(path, module_name)

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

    def loadresources(self):
        return ((number, self.prepareresource(params)) for number, params in self.slidesdata)

    def prepareresource(self, params):
        if type(params) is not list:
            params = [params]
        return (factory(key, dict_[key]) for dict_ in params for key in dict_)

    def build(self, presentationsrc):
        def merge(number, resource):
            resource.merge(self._pptx.Slides(number))

        return [merge(number, resource)
                for number, slidesrc in presentationsrc
                for resource in slidesrc if resource]
