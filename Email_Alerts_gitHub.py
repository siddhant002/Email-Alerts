# -*- coding: utf-8 -*-
"""
Created on Mon Aug 31 11:21:45 2020

@author: djrzyz
"""


import pandas as pd
import win32com.client as win32
import time
from os import listdir
from datetime import datetime
from datetime import datetime, timedelta
import io
from os import path


fileLocation , filename,  sheetname = '', '', ''
expiring =[]



def read_excel(fileLocation,filename,sheetname):
    has_content=False
    df_allert_data=pd.read_excel(fileLocation+"\\"+filename, sheet_name=sheetname) # read the files that has cotent expiry dates.
    df_allert_data.dropna(axis=0,subset=['Expire Date'])
    df_allert_data['Expire Date']=pd.to_datetime(df_allert_data['Expire Date'])
    df_allert_data.dropna(axis=0,subset=['Expire Date'],inplace=True)
    for index_number in df_allert_data.index:
        if datetime.today() < df_allert_data['Expire Date'].iloc[index_number]:
           df_allert_data.drop([index_number], inplace=True)
           df_allert_data.sort_values(by=['Expire Date'],ascending=False ,inplace=True, ignore_index=True )
           df_allert_data.index = df_allert_data.index + 1
    return df_allert_data

def send_email(expiring, df_allert_data ):
    html = """\
<html>
  <head></head>
  <body>
    {0}
  </body>
</html>
""".format(df_allert_data[['Description', 'Station', 'Expire Date']].to_html())
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '' # receiving person's email ID
    mail.Subject = 'Q-Alert expired reminder: '+'('+datetime.today().strftime('%m-%d-%Y')+')'
    mail.Body = 'Hi all, \n The following list items have reached expiration date.'
    mail.HTMLBody = html #this field is optional
    
    # To attach a file to the email (optional):
    # attachment  = "Path to the attachment"
    # mail.Attachments.Add(attachment)

    mail.Send()
    
df_allert_data=read_excel(fileLocation, filename,sheetname)
send_email(expiring, df_allert_data )


