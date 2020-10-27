import numpy as np
import pandas as pd
import smtplib 
import yagmail
from flask_mail import Mail, Message
from flask import Flask
import os
import sys


def xcel_compare(file1, file2):
    res = []
    msg=' '
    mail_flag=False
    err_flag=True
    '''
    abspath = os.path.abspath(sys.argv[0])
    dname = os.path.dirname(abspath)
    os.chdir(dname)
    '''
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
    except Exception as e:
        error = (f'Encountered Exception: {e} '
              f'while reading files to DataFrame, '
              f'file1: {file1}, file2: {file2}')
        return error
    
    try:
        df1 = df1.fillna('')
        df2 = df2.fillna('')
    except Exception as e:
        error = (f'failed to strip nan from dataframes with exception : {e}')
        return error
    
    if df1.equals(df2):
        msg = 'success'
    else:
        error = (f'{file1} and {file2} do not contain same number of columns and rows')
        msg= error
        mail_flag=True

    try:
        comparison_values = df1.values == df2.values
    except Exception as e:
        error = (f'Failed to compare datafarames with exception: {e}')
        return error

    try:
        rows, cols = np.where(comparison_values==False)
    except Exception as e:
        error = (f'Failed to generate Rows and Columns from comparison_values '
              f'with Exception: {e}')
        return error

    try:
        columns = df1.columns
        for item in zip(rows, cols):
            res.append((item[0], item[1], str(df1.iloc[item[0], item[1]]), str(df2.iloc[item[0], item[1]])))
            print(res)
    except Exception as e:
        error = (f'Failed to get differences with exception: {e}')
        return error

    try:
        columns = df1.columns
        for item in zip(rows, cols):
            df1.iloc[item[0], item[1]] = \
                f'{df1.iloc[item[0], item[1]]} --> ' \
                f'{df2.iloc[item[0], item[1]]}'
            res.append((item[0], item[1], str(df1.iloc[item[0], item[1]]), str(df2.iloc[item[0], item[1]])))
    except Exception as e:
        error = (f'Failed to get differences with exception: {e}')
        return error

    try:
          # TODO: better output file (not ./)
        output_file = './Excel_diff.xlsx'
        df1.to_excel(output_file, index=False, header=True)
        print(f'Output resulting diff to: {output_file}')
    except Exception as e:
         # TODO: add output filename to exception logging
        error =(f'Failed to write to Excel file with exception: {e}')
        return error

    try:
        if mail_flag:
            app = Flask(__name__)
            app.debug = True
            app.config['MAIL_SERVER'] = 'smtp.gmail.com'
            app.config['MAIL_PORT'] = 465
            app.config['MAIL_USE_TLS'] = False
            app.config['MAIL_USE_SSL'] = True
            app.config['MAIL_USERNAME'] = '*******@gmail.com'  # enter your email here
            app.config['MAIL_DEFAULT_SENDER'] = '******@gmail.com' # enter your email here
            app.config['MAIL_PASSWORD'] = '******' # enter your password here
            mail = Mail(app)
        
            with app.app_context():
                msg = Message(subject="Hello",
                      sender=app.config.get("MAIL_USERNAME"),
                      recipients=["********@gmail.com"], # replace with your email for testing
                      body=" Hi Team, \n\n There are mismatches between the excels {} and {}" .format(file1,file2)+ "\n Please check the attached excel for more details \n\n Regards,")
                mail.send(msg)
        
    except Exception as e:
        error = (f'Failed to send the mail: {e}')
        return error
    
d=xcel_compare('1.xlsx','2.xlsx')
print(d)
