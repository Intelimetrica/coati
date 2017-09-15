from coati.win32 import copy, execute_commandbar
from coati import utils, excel, powerpoint
import time
import logging
import re
import abc


class AbstractResource():
    __metaclass__ = abc.ABCMeta

    @abc.abstractmethod
    def merge(self, slide):
        pass

class Chart(AbstractResource):

    def __init__(self, name, sheet, chartname):
        self.chartname = chartname
        self.name = name
        self.sheet = sheet

    def merge(self, slide):
        slide.Select()
        xlsx_chart = self.sheet.ChartObjects(self.chartname)
        copy(xlsx_chart)

        ppt_chart = utils.grab_shape(slide, self.name)
        chart_styles = utils.grab_styles(ppt_chart)

        ppt_chart.Delete()

        previous_names = set(shape.Name for shape in slide.Shapes)

        slide.Shapes.Paste()

        for second in range(200):  # Wait up to 20 seconds, checking every 0.1 seconds
            new_names = set(shape.Name for shape in slide.Shapes)
            if new_names != previous_names:
                break
            time.sleep(0.1)

        new_chart = utils.grab_shape(
            slide, list(new_names - previous_names)[0])
        new_chart.Name = self.name

        utils.apply_styles(new_chart, chart_styles)


class Table(AbstractResource):

    def __init__(self, name, sheet, table_range):
        self.name = name
        self.sheet = sheet
        self.table_range = table_range

    def merge(self, slide):
        slide.Select()
        table_shape = utils.grab_shape(slide, self.name)
        styles = utils.grab_styles(table_shape)
        table_shape.Delete()

        self.sheet.Select()
        previous_names = set(shape.Name for shape in slide.Shapes)
        copy(self.sheet.Range(self.table_range))

        execute_commandbar(slide, "PasteSourceFormatting")

        for second in range(200):  # Wait up to 20 seconds, checking every 0.1 seconds
            new_names = set(shape.Name for shape in slide.Shapes)
            if new_names != previous_names:
                break
            time.sleep(0.1)

        new_table = utils.grab_shape(
            slide, list(new_names - previous_names)[0])
        new_table.Name = self.name

        utils.apply_styles(new_table, styles)


class Label(AbstractResource):

    def __init__(self, name, content):
        self.name = name
        self.content = content

    def merge(self, slide):
        slide.Select()
        ppt_label = utils.grab_shape(slide, self.name)
        utils.replace_text(ppt_label, self.content)

class Picture(AbstractResource):

    def __init__(self, name, path):
        self.name = name
        self.path = path

    def merge(self, slide):
        slide.Select()
        width, height = (1, 1)
        try:
            picture = slide.Shapes.AddPicture(self.path, 1, 1, width, height)
        except:
            picture = slide.Shapes.AddPicture(self.path.replace("/", "\\\\"), 1, 1, width, height)
        placeholder = utils.grab_shape(slide, self.name)
        utils.transfer_properties(placeholder, picture)

class Factory(object):
    def __init__(self):
        self._workbooks = []
        self._excel = excel.runexcel()
        self._processfunctions = {'label': self._processlabel,
                                  'table': self._processtable,
                                  'image': self._processimage,
                                  'chart': self._processchart}

    def _getworkbook(self, sheetpath):
        workbook = self._findworkbooks(sheetpath)
        if not workbook:
            workbook = excel.open_xlsx(self._excel, sheetpath)
            self._workbooks.append((sheetpath, workbook))
        return workbook

    def _processlabel(self, shapename, slidetuple):
        return Label(shapename, slidetuple[1])

    def _processtable(self, shapename, slidetuple):
        _stype, sheetpath, srange = slidetuple
        workbook = self._getworkbook(sheetpath)
        name = srange.split('!')[0] if '!' in srange else 1
        sheet = excel.sheet(workbook, name)
        return Table(shapename, sheet, srange)

    def _processimage(self, shapename, slidetuple):
        return Picture(shapename, slidetuple[1])

    def _processchart(self, shapename, slidetuple):
        _stype, sheetpath, chartname = slidetuple
        workbook = self._getworkbook(sheetpath)
        sheet_name = chartname.split('!')[0] if '!' in chartname else 1
        chartname = chartname.split('!')[1] if '!' in chartname else chartname
        sheet = excel.sheet(workbook, sheet_name)
        return Chart(shapename, sheet, chartname)

    def _process(self, slidetype ,shapename, slidetuple):
        return self._processfunctions[slidetype](shapename, slidetuple)

    def _checktype(self, slidetype):
        return slidetype in self._processfunctions

    def prepare(self, shapename, slidetuple):
        slidetype = slidetuple[0]
        if self._checktype(slidetype):
            return self._process(slidetype, shapename, slidetuple)
        else:
            logging.warning("Type '%s' is not a valid type", slidetype)

    def close(self):
        return [workbook.Close() for (_doc, workbook) in self._workbooks]

    def _findworkbooks(self, fname):
        l = [workbook for doc, workbook in self._workbooks if doc is fname]
        return l[0] if len(l) > 0 else None
