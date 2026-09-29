from datetime import datetime
from pathlib import Path
import random
import pandas as pd

def make_sortie_jeudi(current_year):
    
    root_calendar = Path.home() / Path(r"Nextcloud2\BASE_DOCUMENTS_CTG")
    root_calendar = root_calendar / Path(r"2_ACTIVITES_CTG\2-3_ELABORATION_CALENDRIER\PUBLIC")
    file = root_calendar / Path("sorties_jeudi.xlsx")
    
    file = pd.ExcelFile(file)
    sheet_list = file.sheet_names
    year_list = [x for x in sheet_list if '20' in x]
    year_list = [x for x in year_list  if int(x) < current_year]
    
    current_year_df =  pd.read_excel(file,sheet_name=str(current_year))
    current_year_semaine = current_year_df['semaine'].tolist()
    
    dic_sortie = {year : pd.read_excel(file,sheet_name=str(year)) for year in year_list}
    
    dg_list = []
    spy = []
    for semaine in current_year_semaine:
        for i in [1,2,3,4]:
            year = random.choice(year_list)
            sortie = dic_sortie[year].query('semaine==@semaine')[['semaine','Nom','GP','MP','PP','Départ']][0:1]
            if len(sortie['Nom'])>0:
                dg_list.append(sortie)
                break
                
                
    dg = pd.concat(dg_list)


    dg['semaine'] = current_year_semaine
    
    dg = current_year_df[['semaine','mois','N° mois','Date','jour']].merge(dg, left_on="semaine", right_on="semaine")
    
    
    file = root_calendar / Path(str(current_year)) / Path(f"sorties_jeudi_{current_year}.xlsx")
    
    dg.to_excel(file,index=None)
    print(f"Fichier {file} créé")
    
    return dg