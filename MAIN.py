# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 11:39:27 2019
MAIN function for extracting the gift!

@author: ddeen
"""

from xmas_perm import xmas_perm
from xmas_support import importFile, checkExtraction, sendMail
import os


# %% Import file
print("Select and import file with participant list and data")
df = importFile()


# %% Permutation
print("Permutation")
d = xmas_perm(df['index_ex'].values, df.shape[0])


# %% Convert and check
print('Save and convert the extraction (controlling the result)')
df['index_who_receive'] = d
df['who_receive'] = ''
for id_who_give in df.index:
    # convert in name
    df.loc[id_who_give, 'who_receive'] = df.loc[df.loc[id_who_give, 'index_who_receive'], 'who_give']
    # salvo
    who_give = df.loc[id_who_give, 'who_give']
    who_receive = df.loc[id_who_give, 'who_receive']
    exclusion = df.loc[id_who_give, 'exclusion']
    checkExtraction(who_give, who_receive, exclusion)


# %% Send mail
print('Send mail')
oggetto = 'Regalo di Natale!'
testo = 'Ciao #Chi#! \n'+ \
        'Mancano pochi giorni a Natale, ma soprattutto manca poco tempo per preparare il tuo fantastico regalo per... me! Scheerzo, per #to#! \n\n' + \
        "Con discrezione scopri i suoi interessi, ma l'importante è il pensiero! \n" + \
        'e mi raccomando, ci vediamo la Notte di Natale! \n\n' +\
        'Buona giornata \nElena \n\nP.S.(se funziona, altrimenti è colpa di Denny)'

for id_who_give in df.index:
    send_to = [df.loc[id_who_give,'mail']]
    subject = oggetto
    text = testo.replace('#Chi#',df.loc[id_who_give,'who_give']).replace('#to#',df.loc[id_who_give,'who_receive'])
    sendMail(send_to, subject, text)


# %% END: send recap to master
print('Save result in excel file (backup) and send to master')
res_name = 'XmasGift_result.xlsx'
df.to_excel(res_name)
sendMail(None, 'Risultati regalo natale', df.to_html(), mailType='HTML', files=[os.getcwd()+'\\'+res_name])
