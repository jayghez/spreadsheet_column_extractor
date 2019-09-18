
import csv
import pandas as pd
import xlrd
import os
import sys
import datetime




# function for changing microsoft dates to readable 
def xldate_to_datetime(xldate):
    temp = datetime.datetime(1899, 12, 30)
    delta = datetime.timedelta(days=xldate)
    c= (temp+delta)
    return c.strftime("%m/%d/%Y")

lis = []

#file shakers checking for formatting
folder_1 'some_folder'
for file in os.listdir('some_folder/some_directory'):
    if file.endswith(".xlsx"):
        lis.append(file)


#make a dataframe for extracted info 
df3=pd.DataFrame(columns = ['some_col1','some_col2','Date_Field'])

for x in lis[:5]:
    try:
        row_num = []
        row_f = []
        row_e = []
        workbook = xlrd.open_workbook(os.path.join(folder_1,x))
        sheet = workbook.sheet_by_index(0)
        x = x[:-5]
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                if sheet.cell_value(row, col) == 'Date_Field':   
                    for i, cell in enumerate(sheet.col_slice(col)):
                        if cell.ctype == 3:
                            row_num.append(i)
                            cell2 = xldate_to_datetime(cell.value)
                            row_f.append(cell2)
        for rows in range(sheet.nrows):
            for cols in range(sheet.ncols):
                if sheet.cell_value(rows, cols) == 'some_col1':
                    for ze in row_num:
                        if sheet.cell_type(ze,cols)==1:
                            row_e.append(sheet.cell(ze,cols).value)
        for rowss in range(sheet.nrows):
            for colss in range(sheet.ncols):
                if sheet.cell_value(rowss, colss) == 'some_col2':
                    for zee in row_num:
                        if sheet.cell_type(zee,colss)==1:
                            row_e.append(sheet.cell(zee,colss).value)
        max_len = len(row_f)
        row_e_max = len(row_e)
        x_lee = []
        for i in xrange(len(row_f)):
            x_lee.append(x)
        if not max_len == row_e_max :
            row_e.extend('N'* (max_len - row_e_max))
       
        for i in xrange(max_len):
            df3 =df3.append({'some_col1':x_lee[i], 'Date':row_f[i], 'some_col2':row_e[i]}, ignore_index = True)
    except:
        "NA"



# packed into tuples for Loading 
tuples = [tuple(x) for x in df3.values]
tuples

# make data connection

import requests
import json
import urllib

import psycopg2
from  sqlalchemy import create_engine

# Create Connection to database 
try:
    conn = psycopg2.connect(dbname='dbname', user='user' , host='host' ,password='password')
except:
    print "READ MORE"

# make cursor
cur = conn.cursor()


#how and what to load
ex = """INSERT into some_table (some_id, some_dates, some_values) VALUES (%s, %s, %s);"""


#execute 
for i in range(len(tuples)):
    cur.execute(ex,tuples[i])
    
#commit   
conn.commit() 
        