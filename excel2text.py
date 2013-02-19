from xlrd import open_workbook, xldate_as_tuple
from xlutils.display import cell_display
from datetime import datetime, date
import sys

# CSV separator
SEP_CHAR = '|'

def convert(wb_name, sheet_name, start_row):
    text = ''
    wb = open_workbook(wb_name)
    
    for s in wb.sheets():
        if s.name == sheet_name:
            for row in range(s.nrows):
                if row >= start_row - 1:
                    values = []
                    for col in range(s.ncols):
                        data_type = cell_display(s.cell(row, col)).split(" ")[0]
                        if data_type == 'date':
                            # Convert Excel date (looks like a float)
                            the_date = xldate_as_tuple(s.cell(row,col).value, wb.datemode)
                            values.append(date.strftime(date(*the_date[:3]), "%m/%d/%Y"))
                        elif data_type == 'number':
                            values.append(str(int(s.cell(row,col).value)))
                        elif data_type == 'logical':
                            # 0 is false, 1 is true
                            if int(s.cell(row,col).value) == 0:
                                values.append('False')
                            else:
                                values.append('True')
                        else:
                            # Replace CR/LF within cells with a space
                            values.append(str(s.cell(row,col).value).replace("\n", " "))
                    text += SEP_CHAR.join(values) + '\n'
                
    sys.stdout.write(text)

if __name__ == '__main__':

    if len(sys.argv) <> 4:
        print 'Usage: ' + sys.argv[0] + ' file_name.xls sheet_name start_row'
        sys.exit(1)
        
    wb_name = sys.argv[1]
    sheet_name = sys.argv[2]
    start_row = int(sys.argv[3])
    
    convert(wb_name, sheet_name, start_row)
