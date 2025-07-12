from openpyxl.workbook import workbook
from openpyxl.worksheet import worksheet

import pandas as pd
import os
# import multiprocessing
# from openpyxl.workbook import workbook
# from openpyxl.worksheet import worksheet

xls_file = 'prac.xlsx'
# read excel
df = pd.read_excel(xls_file)
e = []
vn = []

def only_alnum(f):
    str_f = ''.join(e for e in f if e.isalnum())
    str_f += ' '
    new = ''.join(n for n in f if n in str_f)
    return new

ent = df['Entity'].values.tolist()
val = df['Values'].values.tolist()

for idx, n in enumerate(ent):
    try:
        # Split Large Values into structered List
        if ':;' in val[idx]:
            v = val[idx].split(':;')
        else:
            v = [val[idx],]
        print('\n')
        print(n)
        e.append(ent[idx])
        # print('\n')
        print(v)
        vn.append(v)
        # print(v[0])

        # If Further Split is required
        # if ',' in v:
        #     vl = v.split(',')
        # else:
        #     vl = v
        # print(vl)

    except KeyboardInterrupt:
        exit()

    except TypeError:
        print('error\n------------------')

my_dic = {
    'Entities': e,
    'Values': vn
    }
df1 = pd.DataFrame(data=my_dic)
excel_path = xls_file

# Get the absolute path of the Excel file
absolute_path = os.path.abspath(excel_path)

# Extract the directory portion of the absolute path
directory_path = os.path.dirname(absolute_path)
# Specify the file name for the Excel file
file_name = 'Results.xlsx'
# Combine the folder path and file name to get the full path
file_path = os.path.join(directory_path, file_name)
# Save the DataFrame to the specified location (full path)
df1.to_excel(file_path, sheet_name='error', index=False)
