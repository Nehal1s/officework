from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
import datetime

x = datetime.datetime.now()
name = "VEEVA_IF_RUN_STATUS " + x.strftime('%d-%b-%Y') + '.xlsx'

wb = Workbook()

ws = wb.active
ws.title = 'VEEVA_IF_RUN_STATUS'

redFill = PatternFill(start_color='BDD7EE',
                   end_color='BDD7EE',
                   fill_type='solid')



_header = ['Batch Name', 'Start time', 'End Time', 'Duration', 'Status', 'Failure reason', 'Error table count',	'Comments']
for col, val in enumerate(_header):
    _cell = ws.cell(1, col+1)
    _cell.value = val
    _cell.fill = redFill



wb.save(filename = name)