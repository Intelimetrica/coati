"""Functionality for handling the logic
of traversing and working with a project.
"""

import os
import imp
import codecs
import logging
from copy import deepcopy
from coati import excel, merge, powerpoint
from coati.resources import Factory
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

    def __init__(self, template_pptx, path, output):
        self.template_pptx = template_pptx
        self.path = path
        self.factory = Factory()
        self.slidesdata = []
        self.output = output

    @property
    def target(self):
        if not hasattr(self, '_target'):
            output = os.path.join(self.path, self.output)
            self._target = powerpoint.new(self.template.app,
                                          output)
        return self._target

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
            self._pptx = self.template.open_pptx()

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

    def loadresources(self):
        self.factory.close()
        return ((number, self.prepareresource(params)) for number, params in self.slidesdata)

    def prepareresource(self, params):
        if type(params) is not list:
            params = [params]
        return (self.factory.prepare(key, dict_[key]) for dict_ in params for key in dict_)

    def build(self, presentationsrc):
        def insert(templateidx, targetidx):
            template = self._pptx.Slides(templateidx)
            template.Copy()
            self.target.Slides.Paste(targetidx)
            slide = self.target.Slides(targetidx)
            slide.Design = template.Design
            return slide

        index = 1
        for number, slidesrc in presentationsrc:
            slide = insert(number, index)
            print index, number
            index += 1
            for resource in slidesrc:
                if resource:
                    resource.merge(slide)

