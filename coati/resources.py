from coati.win32 import copy, execute_commandbar
from coati import utils, excel, powerpoint
import time
import logging

class Chart(object):

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

        slide.Shapes.Paste()
        time.sleep(0.1)

        new_chart = slide.Shapes(self.chartname)
        utils.apply_styles(new_chart, chart_styles)


class Table(object):

    def __init__(self, name, sheet, table_range):
        self.name = name
        self.sheet = sheet
        self.table_range = table_range

    def merge(self, slide):
        slide.Select()
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
        slide.Select()
        ppt_label = utils.grab_shape(slide, self.name)
        utils.replace_text(ppt_label, self.content)

class Picture(object):

    def __init__(self, name, path):
        self.name = name
        self.path = path

    def merge(self, slide):
        slide.Select()
        width, height = (1, 1)
        picture = slide.Shapes.AddPicture(self.path, 1, 1, width, height)
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
        sheet = excel.sheet(workbook, 0)
        return Table(shapename, sheet, srange)

    def _processimage(self, shapename, slidetuple):
        return Picture(shapename, slidetuple[1])

    def _processchart(self, shapename, slidetuple):
        _stype, sheetpath, chartname = slidetuple
        workbook = self._getworkbook(sheetpath)
        sheet = excel.sheet(workbook, 0)
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

