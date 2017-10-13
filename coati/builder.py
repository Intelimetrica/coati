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
import time

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

    def __init__(self, path, template_pptx ,output):
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
            filename = os.path.join(self.path, self.template_pptx)
            self._template = powerpoint.SlideTemplate(filename)
        return self._template

    def loadtemplate(self):
        if not hasattr(self, '_pptx'):
            self._pptx = self.template.open_pptx()

    def loadconfig(self, fname):
        config = self._attemptload(fname)
        if not config or not config.slides or type(config.slides) is not list:
            self.slidesdata = []
            logging.warning("config file no valid")
            if not config:
                logging.warning("config.py not found. Are you on the right directory?")
            return None
        self.slidesdata = flatten(config.slides)
        return config

    def _attemptload(self, fname):
        path = os.path.join(self.path, fname)
        if not os.path.isfile(path):
            return None

        module_name = fname.split('.')[0]
        return loadcode(path, module_name)

    def _getslidesdata(self, slide):
        if slide > 0:
            return [(number, params)
                    for number, params in self.slidesdata
                    if number == slide]
        else:
            return self.slidesdata

    def loadresources(self, slide=0):
        slidesdata = self._getslidesdata(slide)
        return flatten([self.prepareresource(number, params) for number, params in slidesdata])

    def prepareresource(self, number, params):
        if type(params) is not list:
            params = [params]
        def prepare(k, v):
            return self.factory.prepare(k, v[k])
        list_ = []
        for param in params:
            resources = [prepare(key, param) for key in param]
            list_.append((number, resources))
        return list_
        #return [(number, [prepare(key, param[key]) for key in param]) for param in params]

    def build(self):
        def insert(templateidx, targetidx):
            template = self._pptx.Slides(templateidx)
            template.Copy()
            self.target.Slides.Paste(targetidx)
            slide = self.target.Slides(targetidx)
            slide.Design = template.Design
            time.sleep(0.1)
            return slide
        
        def SaveCloseWorkbook(workbook):
            workbook.Save()
            workbook.Close()

        index = 1
        indexWorkbook=0        
        for number, slidedata in self.slidesdata:
            presentationsrc = self.prepareresource(number, slidedata)           
            for number, slidesrc in presentationsrc:            
                slide = insert(number, index)                                   
                for resource in slidesrc:
                    if resource:                    
                        resource.merge(slide)                                             
                if self.factory._workbooks:                                                     
                    [SaveCloseWorkbook(workbook) for (_doc, workbook) in self.factory._workbooks]                                                              
                self.factory._workbooks[:] = []   
            index += 1

    def buildSingleSlide(self, presentationsrc):
        def insert(templateidx, targetidx):
            template = self._pptx.Slides(templateidx)
            template.Copy()
            self.target.Slides.Paste(targetidx)
            slide = self.target.Slides(targetidx)
            slide.Design = template.Design
            time.sleep(0.1)
            return slide

        index = 1
        indexWorkbook=0
        for number, slidesrc in presentationsrc:            
            slide = insert(number, index)                                 
            index += 1            
            for resource in slidesrc:
                if resource:                    
                    resource.merge(slide)
            for resource in slidesrc:                                      
                if hasattr(resource, 'sheet'):                                         
                    self.factory._workbooks[indexWorkbook][1].Save()
                    self.factory._workbooks[indexWorkbook][1].Close()
                    indexWorkbook += 1
                    break   

    def finish(self, close= False):
        self.target.Save()
        if close:
            self.factory.close()
            self._pptx.Close()
            self.target.Close()


