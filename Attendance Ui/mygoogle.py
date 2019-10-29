import sqlite3
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def my_google(my_date):
    my_date=my_date
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('SpreadSheetExample-c253846fa35c.json', scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    sheet = client.open("date").sheet1
    conn=sqlite3.connect('attendance.db')
    c=conn.cursor()
    c.execute(""" SELECT fname,time,date from my_student WHERE date =?""",(my_date,))
    today=c.fetchall()
    for data in range(1,len(today)):
        sheet.update_cell(data,1,today[data][0])
        sheet.update_cell(data,2,today[data][1])
        sheet.update_cell(data,3,today[data][2])


