#! /usr/bin/env python3

import xlwings as xw
import os
from rich import print

'''
This code is directed to open last month crew time sheet and picks up the time sheet from 26 to 31 of last month and
past it in 4 sheets fom the ops stat excel sheet last it opens the current month excel and copy time sheet from 1st to
25th of the current month and then paste it to the ops stat sheet in the correct sheet and position
The code is very dirty and repetitive and is not factored but it do the task faster than I normally do
- I download the files from OnDrive and rename it to lastMonth and currentMonth and keep a copy of the opsstat sheet
in the same directory.
The code will use the same template sheet opsStat.xlsx so no need to save it or even its better to save as different file
This code is the same one used with openpyxl but using xlwings which is much faster due to the use of copy and paste option which is
more faster then iterating over every single line
Created on : 22 Oct 2021

'''

package_dir = os.path.dirname(os.path.abspath(__file__))
try:
    lastMonth = os.path.join(package_dir, 'lastMonth.xlsx')
    currentMonth = os.path.join(package_dir, 'currentMonth.xlsx')
    ops_file = os.path.join(package_dir, 'opsStat.xlsx')
except FileNotFoundError as error:
    print(error)

##########################
# initializing the xlwings books and getting each sheet with its name
ops_book = xw.Book(ops_file)
ops_book.app.visible = False
try:
    ops_sht_cons = ops_book.sheets['Consultants']
    ops_sht_swt = ops_book.sheets['Emp.SWT']
    ops_sht_sls = ops_book.sheets['Emp.SLS']
    ops_sht_over = ops_book.sheets['WTC OverHead']
except Exception as error:
    print(error)

current_book = xw.Book(currentMonth)
current_book.app.visible = False
try:
    current_sht_cons = current_book.sheets['Consultnat']
    current_sht_wtc = current_book.sheets['WTC']
    current_sht_over = current_book.sheets['Timesheet']
except Exception as error:
    print(error)

last_book = xw.Book(lastMonth)
current_book.app.visible = False
try:
    last_sht_cons = last_book.sheets['Consultnat']
    last_sht_wtc = last_book.sheets['WTC']
    last_sht_over = last_book.sheets['Timesheet']
except Exception as error:
    print(error)

##########################
# Clear the ops stat from any data
ops_sht_cons.range('A13:AM140').clear()
ops_sht_swt.range('A13:AM60').clear()
ops_sht_sls.range('A13:AM40').clear()
ops_sht_over.range('A13:AM40').clear()

##########################
# update the current month sheet
##########################
# No 1 - Consultants details:
con_names = current_sht_cons.range('I3:I60').options(ndim=2).value
con_id = current_sht_cons.range('A3:A60').options(ndim=2).value
con_curr_val = current_sht_cons.range('AD3:BB60').value

# Moving the data to the ops stat sheet
ops_sht_cons.range('O13').value = con_curr_val
ops_sht_cons.range('A13').value = con_id
ops_sht_cons.range('D13').value = con_names

# No 2 - SWT & SLS details:
swt_names = current_sht_wtc.range('I3:I28').options(ndim=2).value
swt_id = current_sht_wtc.range('A3:B28').value
swt_curr_val = current_sht_wtc.range('AD3:BB28').value
sls_names = current_sht_wtc.range('I29:I40').options(ndim=2).value
sls_id = current_sht_wtc.range('A29:B40').value
sls_curr_val = current_sht_wtc.range('AD29:BB40').value

ops_sht_swt.range('A13').value = swt_id
ops_sht_swt.range('D13').value = swt_names
ops_sht_swt.range('O13').value = swt_curr_val
ops_sht_sls.range('A13').value = sls_id
ops_sht_sls.range('D13').value = sls_names
ops_sht_sls.range('O13').value = sls_curr_val

# No 3 - overhead details:
over_names = current_sht_over.range('B14:B20').options(ndim=2).value
over_id = current_sht_over.range('A14:A19').options(ndim=2).value
over_curr_val = current_sht_over.range('E14:AC20').value

ops_sht_over.range('A13').value = over_id
ops_sht_over.range('D13').value = over_names
ops_sht_over.range('O13').value = over_curr_val

##########################
# update the last month sheet
##########################
# No 1 - last Consultants details:
con_last_val = last_sht_cons.range('BC3:BH60').value
ops_sht_cons.range('I13').value = con_last_val

# No 2 - last SWT & SLS details:
swt_last_val = last_sht_wtc.range('BC3:BH28').value
sls_last_val = last_sht_wtc.range('BC29:BH40').value

ops_sht_swt.range('I13').value = swt_last_val
ops_sht_sls.range('I13').value = sls_last_val

# No 3 - last overhead details:
over_last_val = last_sht_over.range('AD14:AI20').value
ops_sht_over.range('I13').value = over_last_val

month_date = current_sht_wtc.range('BB1').value
con_ops_month = ops_sht_cons.range('AU139').value
swt_ops_month = ops_sht_swt.range('AU63').value
sls_ops_month = ops_sht_sls.range('AU42').value
over_ops_month = ops_sht_over.range('AU42').value

# Printing the values to console
print(f'The Month ops : {month_date}')
print(f'The Month ops stat for SWT is: {swt_ops_month}')
print(f'The Month ops stat for SLS is: {sls_ops_month}')
print(f'The Month ops stat for consultants is: {con_ops_month}')
print(f'The Month ops stat for over head is: {over_ops_month}')
