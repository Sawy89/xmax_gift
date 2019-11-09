# -*- coding: utf-8 -*-
"""
Created on Sat Oct 26 17:27:09 2019

Support function for the xmas project

@author: ddeen
"""


import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


# %% Support function

def importFile():
    '''
    Function that:
        - make you select an xlsx file
        - read it
        - interprete it and perform some adjustment
    The scope is to read who want to partecipate to xmas project and which
    friends he/she doesn't want to make a gift
    
    Excel file should contain:
        - "who_give" column with the name of the actor who want to do a gift
        - "mail" column with the mail to send the extraction
        - "exclusion" column with the list of name (exact name) of people to
           be excluded from the extraction for the actor separated by comma
    '''
    # Select file
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    
    # Read file
    df = pd.read_excel(filename)
    df.loc[pd.isnull(df['exclusion']), 'exclusion'] = ''
    df['exclusion'] = df['exclusion'].apply(lambda x: x.replace(' ','').split(',') if len(x)>0 else [])
    
    # index come colonna
    df.reset_index(inplace=True)  
    
    # exclusion
    df['index_ex'] = df['index'].apply(lambda x: [x])   # exclude the same person!
    for id_person in df.index:
        # For every person, cycle on exclusion if present
        if len(df.loc[id_person, 'exclusion']) > 0:
            for exc_name in df.loc[id_person, 'exclusion']:
                # find exclusion id from name
                tmp = df.loc[df['who_give']==exc_name,'index']
                if len(tmp) == 1:
                    # esclusione trovata!
                    df.loc[id_person, 'index_ex'].append(tmp.iloc[0])
                else:
                    # Error!
                    print(df.loc[df['who_give']==exc_name])
                    raise Exception('Exclusion ('+exc_name+') di '+df.loc[id_person, 'who_give']+' not found or too many found!! Verify the imported file!')
    
    return df


def checkExtraction(who_give, who_receive, exclusion):
    '''
    Function for veryfing if the result of the extraction is correct!
    INPUT:
        who_give = string with the name of the person who give the gift
        who_receive = string with the name of the person who receive the gift
        exclusion = list of strings of names of person who can't be "who_receive"
    '''
    # Lower case (to avoid error)
    who_give = who_give.lower()
    who_receive = who_receive.lower()
    exclusion = [i.lower() for i in exclusion]
    
    # Verify
    assert who_give != who_receive
    assert who_receive not in exclusion





# %% Mail

# Read mail pass
df_conf = pd.read_csv('xmas_conf.csv', sep=';', header=None)
MAIL_INVIO = df_conf.loc[df_conf[0]=='mail',1].iloc[0]
MAIL_PASS = df_conf.loc[df_conf[0]=='pass',1].iloc[0]


def sendMail(send_to, subject, text, files=None, server="smtp.gmail.com", mailType=None,
                  send_to_cc=None, send_to_bcc=None):
    '''
    function for sending mail
    INPUT: send_to = list of all receivers mail
           subject = subject of the mail 
           text = body of the mail
           files = list of path+name of all attached files
           mailType = 'text' (DEFAULT) of 'HTML' mail format
           send_to_cc = list of all CC mail (copia)
           send_to_bcc = list of all BCC mail (copia nascosta)
    '''
    # if no to is specified, the mail is sento the the sending mail!
    if send_to == None:
        send_to = [MAIL_INVIO]
        
    assert isinstance(send_to, list)
    send_to_all = send_to.copy()        # copy to avoid modification on global list
        
    # Info mail (shown in the header)
    msg = MIMEMultipart()
    msg['From'] = MAIL_INVIO
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    
    # Add CC
    if send_to_cc != None:
        assert isinstance(send_to_cc, list)
        msg['Cc'] = COMMASPACE.join(send_to_cc) # header della mail
        send_to_all += send_to_cc           # send to
    
    # Add BCC (nascosto) --> only in the list of sending (not showing)
    if send_to_bcc != None:
        assert isinstance(send_to_bcc, list)
        send_to_all += send_to_bcc          # send to
    
    # Different mail type (HTML or text DEFAULT)
    if mailType=='HTML':
        text = text.replace('\n\n','<br><br>')
        msg.attach(MIMEText(text,'HTML'))
    else:
        msg.attach(MIMEText(text))

    # Attached documents
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    # SEND MAIL
    smtp = smtplib.SMTP(server)
    smtp.starttls()
    smtp.login(MAIL_INVIO, MAIL_PASS)
    smtp.sendmail(MAIL_INVIO, send_to_all, msg.as_string())
    smtp.quit()

