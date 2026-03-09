

_all_ = ["sg2xlsx"]

from pathlib import Path
from datetime import datetime
import glob
import os

import pandas as pd
import string


def sg_csv2xlsx(file):
    global n
    #n = 48
    f = lambda x: '' if x == '' else float(x.replace(',','.').replace(' ',''))
    
    def numerotation(row):
        global n
        if row['Date'] == '':
            return ''
        else:
            n += 1
            return n
    
    file_path = root / Path(file)
    df = pd.read_csv(file_path,
                    skiprows=6,
                    delimiter=';',
                    decimal =',')
    
    
    
    df = df.fillna('')
    df.to_excel(r"c:\users\franc\Temp\bidon.xlsx")
    df['Débit'] = df['Débit'].map(f)
    df['Crédit'] = df['Crédit'].map(f)
    df['Numero']=[""]*len(df)
    df['Code CTG']=[""]*len(df)
    df['Code CTG affiné']=[""]*len(df)
    df['Solde']=[""]*len(df)
    df['Objet']=[""]*len(df)
    df['Numero'] = df.apply(numerotation,axis=1)
    df = df[['Numero',
            'Code CTG',
            'Code CTG affiné',
            'Date',
            "Nature de l'opération",
            'Débit',
            'Crédit',
            'Solde',
            'Objet']]
    df.to_excel(root / Path("SG.xlsx"),index=None)

def sg2xlsx():
    global n
    now = datetime.now()
    year = now.year
    root = Path.home() / Path(r"Nextcloud2\BASE_FINANCES_CTG") / Path(str(year)) 
    root = root / Path(r"COMPATBILITE-COURANTE FBE")
    df = pd.read_excel(root / Path(f"CTG-Compta-{str(year)}.xlsx"),
                      sheet_name=f"Ecritures{str(year)}")
    n = max(df["Numéro"])
    root_sg = root / Path(r"SOCIETE-GENERALE")
    list_of_files = glob.glob(f'{str(root_sg)}/*.csv')
    file_sg = max(list_of_files, key=os.path.getctime)
    sg_csv2xlsx(file_sg)