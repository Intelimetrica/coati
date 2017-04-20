from coati.win32 import copy, execute_commandbar
from coati import utils, excel, powerpoint
import time
import logging

class Chart(object):

    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet

    def merge(self, slide):
        xlsx_chart = self.sheet.ChartObjects(self.name)
        copy(xlsx_chart)

        ppt_chart = utils.grab_shape(slide, self.name)
        chart_styles = utils.grab_styles(ppt_chart)

        ppt_chart.Delete()

        slide.Shapes.Paste()
        time.sleep(0.1)

        new_chart = slide.Shapes(self.name)
        utils.apply_styles(new_chart, chart_styles)


class Table(object):

    def __init__(self, name, sheet, table_range):
        self.name = name
        self.sheet = sheet
        self.table_range = table_range

    def merge(self, slide):
        table_shape = utils.grab_shape(slide, self.name)
        styles = utils.grab_styles(table_shape)
        table_shape.Delete()

        previous_names = set(shape.Name for shape in slide.Shapes)
        copy(self.sheet.Range(self.table_range))

        execute_commandbar(slide, "PasteSourceFormatting")

        new_names = set(shape.Name for shape in slide.Shapes)
        new_table = utils.grab_shape(
            slide, list(new_names - previous_names)[0])
        new_table.Name = self.name

        utils.apply_styles(new_table, styles)


class Label(object):

    def __init__(self, name, content):
        self.name = name
        self.content = content

    def merge(self, slide):
        ppt_label = utils.grab_shape(slide, self.name)
        utils.replace_text(ppt_label, self.content)

class Picture(object):

    def __init__(self, name, path):
        self.name = name
        self.path = path

    def merge(self, slide):
        width, height = (1, 1)
        picture = slide.Shapes.AddPicture(self.path, 1, 1, width, height)
        placeholder = utils.grab_shape(slide, self.name)
        utils.transfer_properties(placeholder, picture)

def _processlabel(shapename, slidetuple):
    return Label(shapename, slidetuple[1])

def _processtable(shapename, slidetuple):
    _stype, sheet, srange = slidetuple
    return Table(shapename, sheet, srange)

def _processimage(shapename, slidetuple):
    return Picture(shapename, slidetuple[1])

def _processchart(shapename, slidetuple):
    _stype, sheet, _chartname = slidetuple
    return Chart(shapename, sheet)

_processfunctions = {'label': _processlabel,
                     'table': _processtable,
                     'image': _processimage,
                     'chart': _processchart}

def _process(slidetype ,shapename, slidetuple):
    return _processfunctions[slidetype](shapename, slidetuple )

def _checktype(slidetype):
    return slidetype in _processfunctions

def factory(shapename, slidetuple):
    slidetype = slidetuple[0]
    if _checktype(slidetype):
        return _process(slidetype, shapename, slidetuple)
    else:
        logging.warning("Type '%s' is not a valid type", slidetype)

