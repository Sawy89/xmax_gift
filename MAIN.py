# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 11:39:27 2019

@author: ddeen
"""

from xmas_perm import xmas_perm
from xmas_support import importFile, checkExtraction, sendMail
import os


print("Seleziono il file con l'elenco dei partecipanti")
df = importFile()

print("Faccio le permutazioni")
d = xmas_perm(df['index_ex'].values, df.shape[0])

print('Salvo e converto in nome (controllando)')
df['index_to'] = d
df['to'] = ''
for id_chi in df.index:
    # converto in nome
    df.loc[id_chi, 'to'] = df.loc[df.loc[id_chi, 'index_to'], 'Chi']
    # salvo
    chi_fa = df.loc[id_chi, 'Chi']
    chi_riceve = df.loc[id_chi, 'to']
    esclusioni = df.loc[id_chi, 'Esclusioni']
    checkExtraction(chi_fa, chi_riceve, esclusioni)

print('Invio mail')
oggetto = 'Regalo di Natale!'
testo = 'Ciao #Chi#! \n'+ \
        'Mancano pochi giorni a Natale, ma soprattutto manca poco tempo per preparare il tuo fantastico regalo per... me! Scheerzo, per #to#! \n\n' + \
        "Con discrezione scopri i suoi interessi, ma l'importante è il pensiero! \n" + \
        'e mi raccomando, ci vediamo la Notte di Natale! \n\n' +\
        'Buona giornata \nElena \n\nP.S.(se funziona, altrimenti è colpa di Denny)'

for id_chi in df.index:
    send_to = [df.loc[id_chi,'Mail']]
    subject = oggetto
    text = testo.replace('#Chi#',df.loc[id_chi,'Chi']).replace('#to#',df.loc[id_chi,'to'])
    sendMail(send_to, subject, text)


print('Salvataggio Risultato e invio a Elena Patata')
nome_ris = 'RegaliNatale_risultati.xlsx'
df.to_excel(nome_ris)
sendMail(['l.scotto.es@gmail.com'],'Risultati regalo natale',df.to_html(), mailType='HTML', files=[os.getcwd()+'\\'+nome_ris])
#os.remove(nome_ris)
