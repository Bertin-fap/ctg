__all__ = ['inscrit_sejour','search_adherent']     

from pathlib import Path
import datetime
import pathlib
import difflib
import functools
import os
import re
import unicodedata
from tkinter import messagebox

import pandas as pd
import matplotlib.pyplot as plt

import ctg.ctggui.guiglobals as gg

import ctg.ctggui 
import ctg.ctgfuncts
from pathlib import Path

from ctg.ctgfuncts.ctg_tools import read_sortie_csv
from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctgfuncts.ctg_tools import normalize_num_tel

# used to supress no ascii characters suh as accent cedilla,...
nfc = functools.partial(unicodedata.normalize,'NFD')
convert_to_ascii = lambda text : nfc(text). \
                                     encode('ascii', 'ignore'). \
                                     decode('utf-8').\
                                     strip()  

def correct_name(nom):
    if not isinstance(nom,str):
        return 'baba au rhum'
    nom = re.sub(r'\s+', ' ', nom).strip()
    nom = nom.upper()
    nom = convert_to_ascii(nom)
    dic_nom = {"DANIE": "DANIELLE PUECH",
               "JPRDUBONNET":"JEAN-PIERRE ROULET-DUBONNET",
               "JPGUIGA":"JEAN-PIERRE GUIGA"}
    for k,v in dic_nom.items():
        if nom == k : nom=v
    return nom
    

def search_name(nom_list1,nom_list2,id_list,nom):
    
    nom1 = difflib.get_close_matches(nom, nom_list1, n=1)
    nom2 = difflib.get_close_matches(nom, nom_list2, n=1)
    if nom1:
        seq_match = difflib.SequenceMatcher(None, nom, nom1[0])
        ratio1 = seq_match.ratio()
        nom1 = nom1[0]
        id1 =id_list[nom_list1.index(nom1)]
    else:
        ratio1 = 0
        id1 = 0
        
    if nom2:
        seq_match = difflib.SequenceMatcher(None, nom, nom2[0])
        ratio2 = seq_match.ratio()
        nom2 = nom2[0]
        id2 = id_list[nom_list2.index(nom2)]
    else:
        ratio2 = 0
        id2 = 0
    
    
    if ratio2>ratio1:
        nomc = nom2.split()[1]+' '+nom2.split()[0]
        id = id2
    else:
        id = id1
    return id

def inscrit_sejour(file:pathlib.WindowsPath,no_match:list,deffectif,nbr_jours=None,type=None,cout_sejour=None,nom_parcours=None):

    '''builds the DataFrame dg for one event using the csv file of this event.
    The DataFrame dg has 5 columns named :'N° Licencié','Nom','Prénom','Sexe','sejour'
    And EXCEL file is stored in the corresponding EXCEL directory.
    '''

    nom_list1 = (deffectif['Nom']+' '+deffectif['Prénom']).tolist()
    nom_list2 = (deffectif['Prénom']+' '+deffectif['Nom']).tolist()
    id_list = deffectif['N° Licencié'].tolist()
    
    sejour = os.path.splitext(os.path.basename(file))[0]
    
    col = ['N° Licencié','Nom','Prénom','Sexe','Pratique VAE','Nom_brut']
    df_list = []
    list_nom_brut = []
    dg = read_sortie_csv(file)
    
    if dg is not None:
        list_nom =  dg[0].tolist()
        for nom in list_nom:
            nom_brut = nom
            nom = correct_name(nom_brut)
            id = search_name(nom_list1,nom_list2,id_list,nom)
            if id == 0:
                no_match.append((file,nom_brut))
            else:
                #nom_= nomc.split()[0]
                #prenom = nomc.split()[1]
                df_list.append(deffectif.query('`N° Licencié`==@id'))
                list_nom_brut.append(nom_brut)
        dg = pd.concat(df_list)
    else:
        dg = pd.DataFrame([[None,None,None,None,None,None,sejour,]], columns=col+['sejour'])
        list_nom_brut.append('')
    dg['sejour'] = sejour
    if nbr_jours is not None : dg['nbr_jours'] = nbr_jours
    if type is not None :dg['Type'] = type
    if cout_sejour is not None : dg['cout_sejour'] = cout_sejour
    if nom_parcours is not None : dg['nom_parcours'] = nom_parcours
    return dg
    
def search_adherent(year,nom, ctg_path):
    
    #ctg_path  = r"C:\Users\franc\Nextcloud2\BASE_DOCUMENTS_CTG\2_ACTIVITES_CTG\2-2_STATS_DES_SORTIES_ANNEES"
    effectif = ctg.ctgfuncts.EffectifCtg(year,Path(ctg_path))
    
    effectif = effectif.effectif
    deffectif = effectif[['N° Licencié','Nom','Prénom','Sexe','Pratique VAE','N° Tél', 'N° Portable','Age','Adresse','E-mail']]
    
    nom_list1 = (deffectif['Nom']+' '+deffectif['Prénom']).tolist()
    nom_list2 = (deffectif['Prénom']+' '+deffectif['Nom']).tolist()
    id_list = deffectif['N° Licencié'].tolist()
    
    id = search_name(nom_list1, nom_list2, id_list, correct_name(nom))
   
    if id != 0:
        dh = deffectif.query('`N° Licencié`==@id')
        dh = dh.fillna('')
        dh['N° Tél']= dh['N° Tél'].apply(normalize_num_tel)
        dh['N° Portable']= dh['N° Portable'].apply(normalize_num_tel)
        
        txt = '\n'.join([f'{k} : {dh.iloc[0][k]}  ' for k in dh.columns])
        
        file = Path(ctg_path) / Path(str(year)) / Path('STATISTIQUES/EXCEL/synthese_adherent.xlsx')
        if os.path.isfile(file):
            df = pd.read_excel(file)
            df.rename(columns={"Unnamed: 0": "ID",}, inplace=True) 
            id = dh.iloc[0]['N° Licencié']
            
            dh = df.query('ID == @id')[['SORTIE DU DIMANCHE CLUB',
                   'SORTIE DU SAMEDI CLUB', 'SORTIE DU JEUDI CLUB', 'RANDONNEE',
                   'Nbr_SEJOURS']]
          
            if len(dh)>0:
                txt = txt+'\n'+'\n'.join([f'{k} : {dh.iloc[0][k]}  ' for k in dh.columns])
        txt = txt + '\n' + duree_presence_club(ctg_path,id)
        messagebox.showinfo('INFO',txt)
    else:
        messagebox.showinfo('INFO',f'Le nom {nom} est inconnu')

def duree_presence_club(ctg_path,id):
    currentDateTime = datetime.datetime.now()
    date = currentDateTime.date()
    current_year = int(date.strftime("%Y"))
    
    current_year = datetime.datetime.now().year
    
    
    path = Path(ctg_path).parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
    path_history = path / Path(str(current_year)) / Path(r"STATISTIQUES\effectif_history.xlsx")
    if os.path.isfile(path_history):
        df = pd.read_excel(path_history )
        df.rename(columns={'N° Licencié': 'id',}, inplace=True)
        df = df.fillna(0)
        
        duree = []
        if id not in df['id']:
            return ''
        f_dict = {}
        f_dict_ = df.query('id == @id').to_dict(orient='records')[0]
            
        for k,v in f_dict_.items():
            if re.findall(r'\d{4}',str(k)):
                f_dict[int(k)] = v
            else:
                f_dict[k] = v
            
            
        years = [int(x) for x in df.columns if re.findall(r'\d{4}',str(x))]
        x = int(min([f_dict[int(x)] for x in years if f_dict[x] != 0]))
        if x == 1999:
            txt = f'Adhérent(e) au CTG depuis au moins {current_year-x+1} ans'
        else :
            txt = f'Adhérent(e) au CTG  depuis {current_year-x+1} ans'
    else:
        txt= ''
    
    return txt