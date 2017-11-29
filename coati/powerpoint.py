"""Abstractions for working with powerpoint files"""
import tempfile
import shutil
import re
import os
from win32 import run

runpowerpoint = lambda: run('PowerPoint.Application')

def new(app, filename):
    presentation = app.Presentations.Add()
    presentation.SaveAs(filename)
    return presentation

def addslide(pptx, idx, style):
    return pptx.Slides.AddSlide(idx, style)

def open_pptx(app, path):
    return app.Presentations.Open(path)


def slide(pptx, index):
    return pptx.Slides(index + 1)


class SlideTemplate(object):
    """Handle the logic for working with pptx files
    as single-slide templates, so it can be copied
    to a final slideshow.
    """

    def __init__(self, template_path):
        self.template_path = template_path

    def pptx_from_template(self):
        tmp_fd, tmp_path = tempfile.mkstemp(prefix='pptxtmp')
        shutil.copyfile(self.template_path, tmp_path)
        self.pptx = open_pptx(self.app, tmp_path)
        return self.pptx

    def open_pptx(self):
        self.pptx = open_pptx(self.app, self.template_path)
        return self.pptx

    @property
    def app(self):
        if not hasattr(self, '_app'):
            self._app = runpowerpoint()
        return self._app

    def append_to(self, destination_pptx):
        # Copy the only slide
        source_slide = slide(self.pptx, 0)
        source_slide.Copy()

        # Source style
        design = source_slide.Design

        # Append copied slide to target presentation with original style
        destination_pptx.Slides.Paste()
        slide(destination_pptx,
              destination_pptx.Slides.Count - 1).Design = design

    def clean(self):
        self.app.DisplayAlerts = False
        self.pptx.Saved = True
        self.pptx.Close()
    
    def close_runpowerpoint(self):
        self._app.Quit()
        self._app = None

class SlideshowJoiner(object):
    """Manages the logic for joining a set of SlideTemplate
    objects into a final slideshow."""

    def __init__(self, builders):
        """Intialize the joiner object. The sources
        are instances of the SlideBuilder class
        and should be provided in the same order
        of the final slideshow.
        """
        self.sources = builders

    def create_document(self):
        self.app.Presentations.Add(1)
        self.pptx = self.app.ActivePresentation
        fd, path = tempfile.mkstemp()
        self.pptx.SaveAs(path)

    def setup(self):
        self.pptx.PageSetup.SlideHeight = self.sources[
            0].pptx.PageSetup.SlideHeight
        self.pptx.PageSetup.SlideWidth = self.sources[
            0].pptx.PageSetup.SlideWidth

    @property
    def app(self):
        if not hasattr(self, '_app'):
            self._app = runpowerpoint()
        return self._app

    def save(self, path):
        self.pptx.SaveAs(path)

    def build(self):
        self.create_document()

        for index, source in enumerate(self.sources):
            if index == 0:
                self.setup()
            source.build()
            source.template.append_to(self.pptx)
            source.template.clean()

    def quit(self):
        self.app.DisplayAlerts = False
        self.app.Quit()


class SlideSourceOrdering(object):
    """Encapsulate logic for ordering slide order
    relative to their names"""

    def __init__(self, filename_list=None):
        if filename_list is not None:
            self.strategy = self.list_ordering
            self.filename_list = filename_list
        else:
            self.strategy = self.numeric_ordering

    def numeric_ordering(self, paths):
        regex = re.compile(r'\d+$')
        fn = lambda st: int(regex.search(st).group(0))
        return sorted((path for path in paths if os.path.isdir(path)),
                      key=fn)

    def list_ordering(self, paths):
        paths_list = list(paths)
        base_dir = os.path.dirname(paths_list[0])
        full_ordered_paths = [os.path.join(
            base_dir, path) for path in self.filename_list]

        if set(paths_list) != set(full_ordered_paths):
            raise ValueError('Mismatch with order specification')
        else:
            return full_ordered_paths

    def __call__(self, paths):
        return self.strategy(paths)
