_all_ = ["finance_ffct","mouvement_comptable_ffvelo"]

from pathlib import Path
import re
import os
import datetime
from tkinter import messagebox

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

import ctg.ctggui.guiglobals as gg

def verif(root,erreur_txt):
    from math import isnan, nan
    import pandas as pd
    
    df = pd.read_excel(root / Path('budget_adhesion_vs_ffct.xlsx'))
    change_nan = lambda x: 0 if isnan(x) else x
    
    ffct_df = df.groupby(['famille'])
    s = 0
    for x in ffct_df:
        dh = x[1]
        dh = dh.fillna(0)
        montant_cheque = sum((dh['MONTANT_CHEQUE']).tolist())
        total = sum((dh['TOTAL']).tolist())
        entree_club = sum(dh['COTISATION_CLUB '].tolist())
        revue = sum((dh['revue_ffct']).tolist())
        ecart =  montant_cheque - entree_club + revue - total
        ecart =  montant_cheque - entree_club  - total
        if ecart !=0 :
            s += ecart
            print(dh['Nom1'].tolist(),
                  montant_cheque,
                  total,
                  entree_club,
                  revue,
                  ecart)
    erreur_txt = erreur_txt + f'Erreur totale : {s} €'
    messagebox.showinfo("showinfo", erreur_txt)

def plot_pie_synthese(year,finance_dic,n_adherants,root):

    '''Plot from the EXCEL file `synthese.xlsx` the pie plot of 
    the number of participation to the evenments'''

    def func(pct, allvalues):
        absolute = pct / 100.*np.sum(allvalues)
        return "{:.1f}%\n({:.0f}€)".format(pct, absolute)
        #label = f"{absolute} €\n{round(pct,1)} %"
        #return label



    explode_dic = {'Licence FFCT':0.0,
                   'Assurance':0.1,
                   'Revue':0.0,
                   'Cotisation CTG':0.0,
                   }

    data = list(finance_dic.values())
    sorties = list(finance_dic.keys())

    explode = [explode_dic[typ] for typ in sorties]

    _, _, autotexts = plt.pie(data,
                              labels = sorties,
                              autopct = lambda pct: func(pct, data),
                              explode = explode,
                              textprops={'fontsize': 10})
                              
    title = f"Budget inscription de l'année {year} : {sum(data)} €) \n budget par adhérent {round(sum(data)/n_adherants,2)} €"
    plt.title(title, pad=20, fontsize=12)

    _ = plt.setp(autotexts, **{'color':'k', 'weight':'bold', 'fontsize':10})

    plt.tight_layout()

    fig_file = 'BUDGET_INSCRIPTION.png'
    plt.savefig(root / Path(fig_file),bbox_inches='tight')
    plt.show()

def borderau(row):
    
    if isinstance(row['Banque'],str):
        nom = row['Nom1'].replace('_','')
        return nom.split()[0] + " - " + row['Banque']
    return ''

def read_ffct_file(root):
    file = 'finances_ffct.xlsx'
    
    if not os.path.isfile(root / Path(file)):
        messagebox.showinfo("showinfo", f"le fichier:\n{root / Path(file)}\est inexistant!")
        return None
    
    df = pd.read_excel(root / Path(file))
    
    df.rename(columns={"Libellé de l'écriture": "ecriture", }, inplace=True)
    
    df["Date de l'écriture"] = pd.to_datetime(df["Date de l'écriture"],
                                              format="%d/%m/%Y %H:%M:%S")
    df_list = []
    
    pattern = "Prise de licence"
    mask = df['ecriture'].astype(str).str.contains(pattern, case=False, na=False, regex=False)
    df_list.append(df[mask])
    
    pattern = "Prise d'option Assurance"
    mask = df['ecriture'].astype(str).str.contains(pattern, case=False, na=False, regex=False)
    df_list.append(df[mask])
    
    pattern = "Prise d'option Abonnement revue fédérale"
    mask = df['ecriture'].astype(str).str.contains(pattern, case=False, na=False, regex=False)
    df_list.append(df[mask])

    pattern = "Remboursement"
    mask = df['ecriture'].astype(str).str.contains(pattern, case=False, na=False, regex=False)
    df_list.append(df[mask])

    ffct_df = pd.concat(df_list, axis=0)
    ffct_df['Nom'] = ffct_df['ecriture'].apply(lambda x: re.split(r'\sM\s|\sMe\s', x)[-1].strip())
    return ffct_df

def erreur_cheque(df):
    erreur = df['COTISATION_CLUB '].sum() + df['licence FFCT'].sum()
    erreur = erreur+df['assurance_x'].sum()+df['revue_x'].sum()
    erreur = erreur-df['MONTANT_CHEQUE'].sum()
    erreur_cheque_txt = f"erreur entre le montant des chèques le montant des options : {erreur}€\n\n"
    return erreur_cheque_txt

def rapprochement(ffct_df,root):
    dic = {}
    for x in ffct_df.groupby(['Nom']):
        
        
        # list option [licence,assurance,revue]
        list_options = [i for i in x[1]['Débit'].tolist() if i != 0]
        list_remboursement = [i for i in x[1]['Crédit'].tolist() if i !=0 ]

        if len(list_options) >1:
            for y in list_remboursement:
                if y != 0 :list_options.remove(y)
            if len(list_options)==2 : list_options.append(0)
            if len(list_options)!=3:
                list_options = [list_options[0]]+[list_options[2]]+[0]
        else:
            list_options = [0,0,0]
        
        list_options.append(list_options[0]+list_options[1]+list_options[2])
        
        list_options.append(x[1]["Date de l'écriture"].tolist()[0])
        
        nom = x[0][0]   # extraction du nom de l'adhérent
        list_options.append(nom)
        
        txt = '; '.join(x[1]['ecriture'])  # extraction du N° de licence
        l = re.findall(":\s\d{6}\s", txt)
        list_options.append(int(l[0].split(':')[1].strip()))
        dic[nom] = list_options
    dh = pd.DataFrame.from_dict(dic).T
    dh.columns = ["licence","assurance","revue","total","Date","Nom","N°"]
    dh = dh[["N°","Nom","Date","licence","assurance","revue","total"]]
    dh.to_excel(r"c:\users\franc\Temp\spy.xlsx")
   
    df_ctg  = pd.read_excel(root / Path('budget_adhesion.xlsx'),sheet_name='data')
    df_ctg['Nom'] = df_ctg['Nom1'].apply(lambda text: re.split('1er adulte|2e adulte', text)[0].strip())
    
    df = df_ctg.merge(dh, left_on='N°', right_on='N°',how='outer')
    df['bordereau'] = df.apply(borderau, axis=1)
    df = df[['N°', 'Nom1', 'Nom2','famille', 'Date', 'licence FFCT',
           'assurance_x', 'revue_x', 'TOTAL', 'COTISATION_CLUB ', 'Montant intermédiaire',
           'MONTANT_CHEQUE', 'Banque','bordereau', 'LFS', 'LOT',
           'licence', 'assurance_y', 'revue_y', 'total']]
    
    
    erreur_cheque_txt = erreur_cheque(df)

    finance_dic = {}
    finance_dic['Licence FFCT'] = df['licence FFCT'].sum()
    finance_dic['Assurance'] = df['assurance_x'].sum()
    finance_dic['Revue'] = df['revue_x'].sum()
    finance_dic['Cotisation CTG'] = df['COTISATION_CLUB '].sum()
   
    df['erreur'] = df['TOTAL'] - df['assurance_y'] - df['revue_y'] - df['licence']
    erreur_cheque_ffct = f"erreur entre le montant FFCT le montant option : {df['erreur'].sum()}€\n\n"
    erreur_txt = erreur_cheque_txt + erreur_cheque_ffct
    df.rename(columns={'revue_x': 'revue_ctg',
                       'assurance_x': 'assurance_ctg',
                       'licence FFCT' : 'licence_ffct',
                       'assurance_y': 'assurance_ffct',
                        'revue_y' : 'revue_ffct'},
                        inplace=True)
    df.to_excel(root / Path('budget_adhesion_vs_ffct.xlsx'),index=None)
    return finance_dic, dh, erreur_txt
    
def finance_ffct():
    now = datetime.datetime.now()
    year = now.year
    month = now.month
    if month == 11 or month == 12:
        year = year + 1
    root = Path.home() / Path(gg.nextcloud) / Path('BASE_FINANCES_CTG')/ Path(str(year)) / Path('COMPTABILITE-COURANTE')
    root = root / Path('3_COMPTABILITE PRISE DE LICENCE')
    ffct_df = read_ffct_file(root)
    if ffct_df is None:
        return
    finance_dic, dh, erreur_txt = rapprochement(ffct_df,root)
    verif(root, erreur_txt)
    plot_pie_synthese(year,finance_dic,len(dh),root)
    

def mouvement_comptable_ffvelo():
    ctg_path_finance =Path.home() / Path(gg.nextcloud) 
    ctg_path_finance = ctg_path_finance / Path('BASE_FINANCES_CTG') / Path(str(year)) / Path('COMPTABILITE-COURANTE')
    ctg_path_finance = ctg_path_finance / Path('3_COMPTABILITE PRISE DE LICENCE') 
    file = ctg_path_finance / Path("MouvementComptable.xlsx")
    df = pd.read_excel(file,skiprows=3,skipfooter=3)
    solde_initial = float((df.iloc[0]["Solde"]).replace(" E",''))
    df = df.drop([0])
    df = df[df.columns[1:]]
    df["Débit"] = df["Débit"].apply(lambda x: float(x.replace(" E",'')))
    df["Crédit"] = df["Crédit"].apply(lambda x: float(x.replace(" E",'')))
    df["Solde"] = df["Solde"].apply(lambda x: float(x.replace(" E",'')))
    dg = df.groupby(["Libellé"]).sum()
    dh = dg["Nombre"]
    ffct_dic = {}
    
    ffct_dic["Prise de Licence"] = [dg["Nombre"]["Prise de Licence"]-dg["Nombre"]["Annulation de Licence"],
                                    dg["Débit"]["Prise de Licence"]-dg["Crédit"]["Annulation de Licence"]]
    ffct_dic["Souscription Assurance"] = [dg["Nombre"]["Souscription Assurance"]-dg["Nombre"]["Annulation assurance"],
                                          dg["Débit"]["Souscription Assurance"]-dg["Crédit"]["Annulation assurance"]]
    ffct_dic["Souscription Revue fédérale"] = [dg["Nombre"]["Souscription Revue fédérale"]-dg["Nombre"]["Annulation revue federale"],
                                          dg["Débit"]["Souscription Revue fédérale"]-dg["Crédit"]["Annulation revue federale"]]
    ffct_dic["Débit Divers"] = [dg["Nombre"]["Débit Divers"],
                                dg["Débit"]["Débit Divers"]]
    ffct_dic["Réaffiliation Club"] = [dg["Nombre"]["Réaffiliation Club"],
                                dg["Débit"]["Réaffiliation Club"]]
    ffct_dic["Souscription d'Assurance Option A"] = [dg["Nombre"]["Souscription d'Assurance Option A"],
                                dg["Débit"]["Souscription d'Assurance Option A"]]
    ffct_dic["Souscription d'Assurance Option B"] = [dg["Nombre"]["Souscription d'Assurance Option B"],
                                dg["Débit"]["Souscription d'Assurance Option B"]]
    ffct_dic["Prélèvement"] = [dg["Nombre"]["Prélèvement"],
                                dg["Crédit"]["Prélèvement"]]
    ffct_dic["Solde initial"] = ['',
                                solde_initial]
    l = sum([x[1] for x in list(ffct_dic.values())[:-2]])
    print(l)
    print(ffct_dic["Prélèvement"][1])
    print(solde_initial)

    print(solde_initial - (l-ffct_dic["Prélèvement"][1]))
   
    ffct_dic["Solde final"] = ['',
                                solde_initial - (l-ffct_dic["Prélèvement"][1])]
    ffct_df = pd.DataFrame.from_dict(ffct_dic).T
    ffct_df.columns = ["Nombre","Somme (€)"]
    ffct_df.to_excel(ctg_path_finance / Path("MouvementComptable_resume.xlsx"))   