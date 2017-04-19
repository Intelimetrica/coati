from __future__ import print_function
import tempfile
import os
from officereports import excel, powerpoint
import officereports.constants as xlcharts
from officereports.settings import load

def transpose(lst):
    return [[item] for item in lst]

def main():
    # Load configuration from envfile
    load()

    # Excel tests
    xlapp = excel.runexcel()

    # Create temporary file for excel document
    fd, path = tempfile.mkstemp(prefix='tmp')
    xlsx = excel.open_xlsx(xlapp, path)

    # Create cell accessor helper (optional)
    cells = excel.RangeAccessor(excel.sheet(xlsx, 0))
    
    # Create data (column-wise)
    cells['A1:A10'] = transpose(range(10))
    cells['B1:B10'] = transpose(x ** 2 for x in range(10))

    # Add scatter chart
    chart_type = xlcharts.XL_XY_SCATTER
    scatter = excel.Chart(xlapp, cells.sheet, chart_type)
    scatter(cells['$A$1:$B$10'])
    
    for a, b in cells['A1:B10']:
        print(a + b)

    # PowerPoint tests
    pptApp = powerpoint.runpowerpoint()

    # Use loaded configuration for opening target file
    pth = os.path.abspath(os.environ['SLIDES_FILE'])
    pptx = powerpoint.open_pptx(pptApp, pth)

    xlapp.DisplayAlerts = False
    xlapp.Quit()
    pptApp.Quit()

if __name__ == '__main__':
    main()
