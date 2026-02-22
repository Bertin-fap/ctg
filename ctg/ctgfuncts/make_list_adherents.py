_all_ = ["make_list_adherents"]

from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

from ctg.ctgfuncts.ctg_classes import EffectifCtg
import ctg.ctgfuncts as ctg

year = 2025

def make_list_adherents(ctg_path):

    current_year = datetime.now().year
    eff = ctg.EffectifCtg(current_year,ctg_path)
    df = eff.effectif
    df = df.sort_values(by=['Nom','Prénom'])
    df.rename(columns={'N° Licencié': 'Licence',
                       'Nom': 'Nom',
                       'Prénom':'Prénom',
                       'Date de naissance':'D de N',
                       'Sexe':'S',}, inplace=True)
    df['D de N'] =  df['D de N'].astype(str)
    result_path = Path(ctg_path).parent.parent / Path(r'1_FONCTIONNEMENT_CTG/1-1_BASE_ADHERENTS_CTG') / Path(str(current_year))
    output_file = result_path / f'liste_adherents_CTG_{current_year}.docx'
    template_path_docx = Path(__file__).parent.parent / 'ctgfuncts' / 'CTG_RefFiles'
    
    frameworks = []
    for idx,row in df.iterrows():
        frameworks.append(dict(id=row['Licence'],
                                  surname=row['Nom'],
                                  name=row['Prénom'],
                                  ddn=row['D de N'],
                                  s=row['S'] ))
            
    
    context ={'year': current_year,
              'frameworks': frameworks}
    
    template_docx = template_path_docx / 'template_Liste_CTG.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    convert(output_file)
    messagebox.showinfo("CTG_METER", f"Le fichier {output_file} a été créé")

def make_list_emargement(ctg_path,day,month,year):

    """
    Elle se compose de tous les membres de l'association, de plus de 16 ans, à jour de leur cotisation. 
    L'assemblée générale de l'association se réunit une fois par an, moins de six mois après la clôture de l'exercice comptable. 
    La convocation est adressée à tous les membres par écrit au moins quinze jours avant la date fixée.
    Elle comprendra obligatoirement l'ordre du jour établi par le comité directeur.  
    """

    date_ag = f"{day}-{month}-{year}"
    eff = ctg.EffectifCtg(year,ctg_path)
    df = eff.effectif
    df = df[['N° Licencié','Nom','Prénom','Date de naissance','Sexe']]
    # add column Age compute at the 30 september of year
    df['D de N'] = df['Date de naissance']
    df['Date de naissance'] = pd.to_datetime(df['Date de naissance'],
                                                 format="%d/%m/%Y")
    
    df = df.sort_values(by=['Nom','Prénom'])
    df.rename(columns={'N° Licencié': 'Licence',
                     'Nom': 'Nom',
                     'Prénom':'Prénom',
                     'Sexe':'S'}, inplace=True)
    
    df['Age']  = df['Date de naissance'].apply(lambda x :
                                                  (pd.Timestamp(int(year), 12, 7)-x).days/365)
    df = df.query('Age>16')
    quorum = int(len(df)/3)
    df['D de N'] =  df['D de N'].astype(str)
    result_path = Path(ctg_path).parent.parent / Path(r'1_FONCTIONNEMENT_CTG/1-1_BASE_ADHERENTS_CTG') / Path(str(year))
    output_file = result_path / f'liste_emargement_CTG_{year}.docx'
    template_path_docx = Path(__file__).parent.parent / 'ctgfuncts' / 'CTG_RefFiles'
    
    l = []
    for idx,row in df.iterrows():
        l.append(dict(id=row['Licence'],
                     surname=row['Nom'],
                     name=row['Prénom'],
                     ddn=row['D de N'],
                     s="M." if row['S']=="M" else "Me."))
            
    long = 22
    frameworks = []
    for i_dep in range(0,len(l),long):
        if i_dep+long < len(l):
            frameworks.append([l[index] for index in range(i_dep,i_dep+long)])
        else:
            frameworks.append([l[index] for index in range(i_dep,len(l))])      
    context ={'year': year,
             'date_ag': date_ag,
             'quorum':quorum,
             'n_adherents':len(df),
             'frameworks': frameworks}

    template_docx = template_path_docx / 'template_Liste_emargement_CTG.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    convert(output_file)
    messagebox.showinfo("CTG_METER", f"Le fichier {output_file} a été créé")
