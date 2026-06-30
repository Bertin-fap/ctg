__all__ = ['inscrit_sejour',]     

from pathlib import Path
import pathlib
import difflib
import functools
import os
import unicodedata

import pandas as pd
import matplotlib.pyplot as plt

import ctg.ctggui.guiglobals as gg

import ctg.ctggui 
import ctg.ctgfuncts
from pathlib import Path

from ctg.ctgfuncts.ctg_tools import read_sortie_csv

# used to supress no ascii characters suh as accent cedilla,...
nfc = functools.partial(unicodedata.normalize,'NFD')
convert_to_ascii = lambda text : nfc(text). \
                                     encode('ascii', 'ignore'). \
                                     decode('utf-8').\
                                     strip()  

def correct_name(nom):
    
    nom = nom.upper()
    nom = convert_to_ascii(nom)
    dic_nom = {"DANIE": "DANIELLE PUECH",
               "JPRDUBONNET":"JEAN-PIERRE ROULET-DUBONNET",
               "JPGUIGA":"JEAN-PIERRE GUIGA"}
    for k,v in dic_nom.items():
        if nom == k : nom=v
    return nom
    

def search_name(nom_list1,nom_list2,nom):
    
    nom1 = difflib.get_close_matches(nom, nom_list1, n=1)
    nom2 = difflib.get_close_matches(nom, nom_list2, n=1)
    
    if nom1:
        seq_match = difflib.SequenceMatcher(None, nom, nom1[0])
        ratio1 = seq_match.ratio()
        nom1 = nom1[0]
    else:
        ratio1 = 0
        nom1 = ' '
        
    if nom2:
        seq_match = difflib.SequenceMatcher(None, nom, nom2[0])
        ratio2 = seq_match.ratio()
        nom2 = nom2[0]
    else:
        ratio2 = 0
        nom2 = ' '
    
    
    if ratio2>ratio1:
        nomc = nom2.split()[1]+' '+nom2.split()[0]
    else:
        nomc = nom1
    return nomc

def inscrit_sejour(file:pathlib.WindowsPath,no_match:list,deffectif,nbr_jours=None,type=None,cout_sejour=None,nom_parcours=None):

    '''builds the DataFrame dg for one event using the csv file of this event.
    The DataFrame dg has 5 columns named :'N° Licencié','Nom','Prénom','Sexe','sejour'
    And EXCEL file is stored in the corresponding EXCEL directory.
    '''

    nom_list1 = (deffectif['Nom']+' '+deffectif['Prénom']).tolist()
    nom_list2 = deffectif['Prénom']+' '+deffectif['Nom']
    
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
            nomc = search_name(nom_list1,nom_list2,nom)
            if nomc == ' ':
                no_match.append((file,nom_brut))
            else:
                nom_= nomc.split()[0]
                prenom = nomc.split()[1]
                df_list.append(deffectif.query('Nom==@nom_ and Prénom==@prenom'))
                list_nom_brut.append(nom_brut)
        dg = pd.concat(df_list)
    else:
        dg = pd.DataFrame([[None,None,None,None,None,None,sejour,]], columns=col+['sejour'])
        return dg
    
    dg['Nom_brut'] = list_nom_brut
    dg['sejour'] = sejour
    if nbr_jours is not None : dg['nbr_jours'] = nbr_jours
    if type is not None :dg['Type'] = type
    if cout_sejour is not None : dg['cout_sejour'] = cout_sejour
    if nom_parcours is not None : dg['nom_parcours'] = nom_parcours
    return dg
