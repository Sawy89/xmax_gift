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

# Read mail pass
df_conf = pd.read_csv('xmas_conf.csv', sep=';', header=None)
MAIL_INVIO = df_conf.loc[df_conf[0]=='mail_invio',1].iloc[0]
MAIL_PASS = df_conf.loc[df_conf[0]=='mail_pass',1].iloc[0]





def importFile():
    '''
    Function that:
        - make you select an xlsx file
        - read it
        - interprete it and perform some adjustment
    The scope is to read who want to partecipate to xmas project and which
    friends he/she doesn't want to make a gift
    
    Excel file should contain:
        - "Chi" column with the name of the actor who want to do a gift
        - "Mail" column with the mail to send the extraction
        - "Esclusioni" column with the list of name (exact name) of people to
           be excluded from the extraction for the actor separated by comma
    '''
    # Select file
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    
    # Read file
    df = pd.read_excel(filename)
    df.loc[pd.isnull(df['Esclusioni']), 'Esclusioni'] = ''
    df['Esclusioni'] = df['Esclusioni'].apply(lambda x: x.replace(' ','').split(',') if len(x)>0 else [])
    
    # index come colonna
    df.reset_index(inplace=True)  
    
    # Esclusioni
    df['index_ex'] = df['index'].apply(lambda x: [x])  # escludo se stesso
    for id_persona in df.index:
        # Per ogni persona, ciclo sulle esclusioni (se ci sono)
        if len(df.loc[id_persona, 'Esclusioni']) > 0:
            for nome_ex in df.loc[id_persona, 'Esclusioni']:
                # cerco l'esclusione per trasformarla in index
                tmp = df.loc[df['Chi']==nome_ex,'index']
                if len(tmp) == 1:
                    # esclusione trovata!
                    df.loc[id_persona, 'index_ex'].append(tmp.iloc[0])
                else:
                    # Erorre!
                    print(df.loc[df['Chi']==nome_ex])
                    raise Exception('Esclusione ('+nome_ex+') di '+df.loc[id_persona, 'Chi']+' non trovata o trovate più di una!!')
    
    return df


def checkExtraction(chi_fa, chi_riceve, esclusioni):
    '''
    Funzione che verifica se l'estrazione è stata effettuata correttamente!
    INPUT:
        - il chi_fa è una stringa con il nome di chi effettua il regalo
        - il chi_riceve è una stringa con il nome di chi riceve il regalo
        - esclusioni è una lista contenente i nomi di chi non possono ricevere il regalo dal regalante
    '''
    # tutto minuscolo, per evitare errori
    chi_fa = chi_fa.lower()
    chi_riceve = chi_riceve.lower()
    esclusioni = [i.lower() for i in esclusioni]
    
    # verifico
    assert chi_fa != chi_riceve
    assert chi_riceve not in esclusioni



def sendMail(send_to, subject, text, files=None, server="smtp.gmail.com", mailType=None,
                  send_to_cc=None, send_to_bcc=None):
    '''
    funzione per inviare MAIL
    INPUT: send_to come lista delle mail dei ricevitori
           subject = oggetto, text = corpo della mail
           files come lista degli allegati (se presente)
           mailType = indica se in formato testo (default) o 'HTML'
           send_to_cc e send_to_bcc sono liste con gli indirizzi mail da inserire
               (se presenti) in CC e BCC (copia nascosta)           
    '''
   
    assert isinstance(send_to, list)
    send_to_all = send_to.copy() # copia per evitare di modificare la lista globale
        
    # Info mail (quelle mostrate nell'header)
    msg = MIMEMultipart()
    msg['From'] = MAIL_INVIO
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    
    # Add CC
    if send_to_cc != None:
        assert isinstance(send_to_cc, list)
        msg['Cc'] = COMMASPACE.join(send_to_cc) # header della mail
        send_to_all += send_to_cc # a chi la invia
    
    # Add BCC (nascosto) --> solo nella lista degli invii!
    if send_to_bcc != None:
        assert isinstance(send_to_bcc, list)
        send_to_all += send_to_bcc # a chi la invia
    
    # Differenze per tipo mail
    if mailType=='HTML':
        text = text.replace('\n\n','<br><br>')
        msg.attach(MIMEText(text,'HTML'))
    else:
        msg.attach(MIMEText(text))

    # Allegati
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



if __name__ == '__main__':
    df = importFile()