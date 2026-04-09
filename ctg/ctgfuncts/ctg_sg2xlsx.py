_all_ = ["sg2xlsx"]

from pathlib import Path
from datetime import datetime
from tkinter import messagebox
import math
import glob
import os
import re
import shutil

import pandas as pd
import numpy as np

import ctg.ctggui.guiglobals as gg

def sg2xlsx():
    global n,last_solde
    f = lambda x: '' if x == '' else float(x.replace(',','.').replace(' ',''))
    
    def numerotation(row):
        global n,last_solde
        if row['Date'] == '':
            return ''
        else:
            n += 1
            return n

    def solde(row):
        global n,last_solde
        if row['Débit'] == '' and row['Crédit'] == '':
            return ''
        elif row['Débit'] == '':
            if isinstance(row['Crédit'],float):
                last_solde = last_solde + row['Crédit']
            else:   
                last_solde = last_solde + float(row['Crédit'].replace(',','.').replace(' ',''))
            return last_solde
        if isinstance(row['Débit'],float):
            last_solde = last_solde + row['Débit']
        else:
            last_solde = last_solde + float(row['Débit'].replace(',','.').replace(' ',''))
        return last_solde

    def get_solde_initial(year):
        
        root = root_finance / Path(str(year-1))  / Path(r"COMPTABILITE-COURANTE")
        file = root / Path(f"CTG-Compta-{str(year-1)}.xlsx")
        df = pd.read_excel(file,
                          sheet_name=f"Ecritures{str(year-1)}")

        solde_list =  [x for x in df["Solde"].tolist() if not math.isnan(x)]
        solde_initial = solde_list[-1]

        return solde_initial
        

    def read_file_finance_ctg(year,solde_initial):
        
        df = pd.read_excel(root / Path(f"CTG-Compta-{str(year)}.xlsx"),
                          sheet_name=f"Ecritures{str(year)}")
        if len(df)>0:
            df = df.fillna('')
            
            n = max([x for x in df["Numéro"] if x != ''])
            date_list = [datetime.strptime(x,'%d/%m/%Y') for x in df["Date SG"].tolist() if x != '']
            last_date = date_list[-1]
    
            credit_list = [x.replace(',','.').replace(' ','') if isinstance(x,str) else x for x in df["Credit"].tolist()]
            credit_list = [float(x) for x in credit_list if x != '']
    
            debit_list = [x.replace(',','.').replace(' ','') if isinstance(x,str) else x for x in df["Debit"].tolist()]
            debit_list = [float(x) for x in debit_list if x != '']
            last_solde = solde_initial + sum(credit_list) + sum(debit_list)
        else:
            last_solde = solde_initial
            n = 0
            last_date = datetime(int(year)-1,1,1)

        return n,last_date, last_solde

    def read_file_sg():
        root_sg = root / Path(r"4_SOCIETE-GENERALE")
        list_of_files = glob.glob(f'{str(root_sg)}/*.csv')
        if len(list_of_files) == 0:
            messagebox.showinfo("showinfo", 'Pas de fichiers SG détectés')
            df = None
            date_solde_sg = None
            solde_sg = None
            return df,date_solde_sg,solde_sg
            
        file_sg = max(list_of_files, key=os.path.getctime)
        file_path = root / Path(file_sg)
        df = pd.read_csv(file_path,skiprows=6,delimiter=';',decimal =',')
        df = df.fillna('')

        with open(file_path, 'r') as file:
            lines = file.readlines()

        pattern = r'\d{2}/\d{2}/\d{4}'
        date_solde = re.findall(pattern, lines[3])
        date_solde_sg = datetime.strptime(date_solde[0],'%d/%m/%Y') 
        
        pattern = r'[\d\s,]{2,14}'
        solde = re.findall(pattern, lines[4])
        solde_sg = float(solde[0].replace(' ','').replace(',','.'))
        return df,date_solde_sg,solde_sg
        
    def move_file_sg(year,root_finance):
        path_telechargement = Path.home() / Path('Downloads')

        files = [x for x in os.listdir(path_telechargement) if re.search('\d{6}-\d{16}\.csv',x)]
        root = root_finance / Path(str(year))  / Path(r"COMPTABILITE-COURANTE\4_SOCIETE-GENERALE")

        for file in files:
           shutil.move(path_telechargement / Path(file), root / Path(file))
        
    global n,last_solde
    now = datetime.now()
    year = now.year
    month = now.month
    if month == 11 or month == 12:
        year = year + 1
    root_finance = Path.home() / Path(gg.nextcloud) / Path(r"BASE_FINANCES_CTG")
    root = root_finance / Path(str(year))  / Path(r"COMPTABILITE-COURANTE")
    
    move_file_sg(year,root_finance)

    solde_initial = get_solde_initial(year)
    n,last_date, last_solde = read_file_finance_ctg(year,solde_initial)
    df,date_solde_sg,solde_sg = read_file_sg()
    date_sg_max = max([datetime.strptime(x,'%d/%m/%Y') for x in df["Date"].tolist() if x != '' ])
    df['Date_']= df['Date'].apply(lambda x: datetime.strptime(x,'%d/%m/%Y') if x != '' else date_sg_max)
    
    df = df.query('Date_>@last_date')
    if len(df)==0:
        messagebox.showinfo("showinfo", 'pas de nouveaux mouvements bancaires détectés')
        return
    df['Code CTG']=[""]*len(df)
    df['Code CTG affiné']=[""]*len(df)
    df['Solde'] = df.apply(solde,axis=1)
    df['Objet'] = [""]*len(df)
    df['Numero'] = df.apply(numerotation,axis=1)
    df['pj'] = ''
    
    
    
    numero_dep = min([x for x in df['Numero'] if x !=''])
    indice_dep = np.where(df["Numero"] == numero_dep)[0][0]
    df = df[indice_dep:]

    m_list = []
    for idx, date in enumerate(df['Date']):
        if idx == 0: 
            m_list.append(int(date.split('/')[1]))
        elif date == '':
            m_list.append(m_list[idx-1])
        else:
            m_list.append(int(date.split('/')[1]))
    df['m'] = m_list 
    
    df = df[['Numero',
            'Code CTG',
            'Code CTG affiné',
            'Date',
            "Nature de l'opération",
            'Débit',
            'Crédit',
            'pj',
            'Solde',
            'Objet',
            'm']]
    file = root / Path("SG.xlsx")
    df.to_excel(file,index=None)
    erreur_solde = round(solde_sg - [x for x in df['Solde'].tolist() if x != ''][-1],2)
    
    txt = f"votre fichier est disponible sous : {str(file)}\n"
    txt = txt + f"l'erreur sur lesolde est de : {erreur_solde} €"
    messagebox.showinfo("showinfo", txt)
    