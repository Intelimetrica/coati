"""Abstactions for working with excel files."""
from win32 import run
import constants as xlcharts

runexcel = lambda: run('Excel.Application')


def open_xlsx(app, path):
    return app.Workbooks.Open(path)


def sheet(xlsx, index):
    return xlsx.Sheets(index + 1)


def chart(sheet, reference):
    return sheet.ChartObjects(reference)


class RangeAccessor(object):

    def __init__(self, sheet):
        self.sheet = sheet

    def __setitem__(self, key, value):
        self.sheet.Range(key).Value = value

    def __getitem__(self, key):
        rng = self.sheet.Range(key)
        add_cell_iteration(rng)
        return rng


def add_cell_iteration(cells):
    def iteration(self):
        return iter(self.Value)

    setattr(cells.__class__, '__iter__', iteration)


class Chart(object):

    def __init__(self, excel, sheet, chart_type):
        self.excel = excel
        self.sheet = sheet
        self.chart_type = chart_type

    def __call__(self, cell_range):
        self.sheet.Shapes.AddChart2(-1, xlcharts.XL_XY_SCATTER).Select()
        self.sheet.ChartObjects(self.sheet.Shapes[0].name).Activate()
        self.excel.ActiveChart.SetSourceData(cell_range)
