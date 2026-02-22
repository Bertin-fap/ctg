# Standard library imports
import os
import pathlib
import re
from datetime import datetime
from pathlib import Path
from math import asin, cos, radians, sin, sqrt
from math import nan, isnan
from tkinter import messagebox

# Third party imports
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Internal imports
from ctg.ctgfuncts.ctg_tools import built_lat_long
from ctg.ctgfuncts.ctg_cd import plot_cd_evolution

class EffectifCtg():


    def __init__(self,year:str,ctg_path:pathlib.WindowsPath,rebond=True):

        self.year = year
        self.ctg_path = Path(ctg_path)
        
        # get effectif of the year year
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_root = path_effectif  / Path(str(self.year)) 
        df = pd.read_excel(path_root / Path(str(self.year)+'.xlsx'))
        if 'N° Licencié' not in df.columns:
            df = df.rename(columns={'N°': 'N° Licencié'})
        if 'Ville' not in df.columns:
            df['Ville'] = df['Adresse'].apply(lambda row: re.split(r'\s+\d{5,6}\s+', row)[-1])
        if 'Nom, Prénom'  in df.columns:
            df['Nom'] = df['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[1])
            df['Prénom'] = df['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[2])
            df['Sexe'] = df['Sexe'].apply(lambda row:row[0])
            df = df.rename(columns={'N°': 'N° Licencié',})
            vae_df = pd.read_excel(path_root / "VAE.xlsx")
            vae_dic = dict(zip(vae_df["N° Licencié"], (vae_df["Pratique VAE"])))
            df["Pratique VAE"] = df["N° Licencié"].map(vae_dic)
        # add column Age compute at the 30 september of year
        df['Date de naissance'] = pd.to_datetime(df['Date de naissance'],
                                                 format="%d/%m/%Y")
        df['Age']  = df['Date de naissance'].apply(lambda x :
                                                  (pd.Timestamp(int(year), 9, 30)-x).days/365)
        # add column distance from Grenoble
        dh = built_lat_long(df)
        df['distance'] = df.apply(lambda row: self.distance_(row, dh),axis=1)
        
        self.effectif = df      # effectif year
        self.effectif = self.correction_effectif()

        # get effectif of the year year-1
        year_1 = int(year)-1
        path_root = path_effectif / Path(str(year_1))
        path_file = path_root /Path(str(year_1)+'.xlsx')
        if path_file.is_file():
            df_1 = pd.read_excel(path_file)
            
            if 'Nom, Prénom'  in df_1.columns:
                df_1['Nom'] = df_1['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[1])
                df_1['Prénom'] = df_1['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[2])
                df_1['Sexe'] = df_1['Sexe'].apply(lambda row:row[0])
                df_1 = df_1.rename(columns={'N°': 'N° Licencié',})
                vae_df = pd.read_excel(path_root / "VAE.xlsx")
                vae_dic = dict(zip(vae_df["N° Licencié"], (vae_df["Pratique VAE"])))
                df_1["Pratique VAE"] = df_1["N° Licencié"].map(vae_dic)
            
            
            df_1['Date de naissance'] = pd.to_datetime(df_1['Date de naissance'],
                                                       format="%d/%m/%Y")
            
            
            df_1['Age']  = df_1['Date de naissance'].apply(lambda x :
                                                          (pd.Timestamp(int(year), 9, 30)-x).days/365)
            
            self.effectif_1 = df_1  # effectif year moins un
        else:
            self.effectif_1 = None

        # deal with nouveaux entrants and sortants
        (self.moy_age_entrants,
        self.nbr_nouveaux_membres,
        self.nouveaux_membres_noms) = self.nouveaux_entrants()
        self.moy_age_sortants, self.nbr_sortants, self.sortants_noms = self.sortants()
        if rebond : self.rebond = self.get_rebond()

        self.cotisation_licence,self.cotisation_totale, self.cotisation_ctg = self.cotisation()

        self.membres_sympathisants, self.nbr_membres_sympathisants = self.membres_sympathisants()
        self.effectif_tot = self.add_membres_sympathisants()

    @staticmethod
    def distance_(row,dh):

        phi1, lon1 = dh.query("Ville=='GRENOBLE'")[['long','lat']].values.flatten()
        phi1, lon1 = radians(phi1), radians(lon1)
        phi2, lon2 = radians(row['long']), radians(row['lat'])
        rad = 6371
        dist = 2 * rad * asin(sqrt(sin((phi2 - phi1) / 2) ** 2
                                 + cos(phi1) * cos(phi2) * sin((lon2 - lon1) / 2) ** 2))
        return np.round(dist,1)

    def stat(self):

        da = self.effectif.groupby('Sexe')['Age'].agg(['count','median','max','min'])
        res = self.effectif['Age'].agg(['count','median','max','min']).tolist()
        stat = []
        stat.append(f"Année : {self.year}")
        stat.append(" ")
        nbr_membres = round(res[0],0)
        stat.append(f"Nombre d'adhérents : {nbr_membres}")
        nbr_femmes = da.loc['F','count']
        stat.append(f"Nombre de femmes : {nbr_femmes} ({round(100*nbr_femmes/nbr_membres,1)} %)")
        nbr_hommes = da.loc['M','count']
        stat.append(f"Nombre d'hommes : {nbr_hommes} ({round(100*nbr_hommes/nbr_membres,1)} %)")
        stat.append(' ')
        stat.append(f"Age médian total : {round(res[1],1)} ans")
        stat.append(f"Age maximum : {round(res[2],1)} ans")
        stat.append(f"Age minimum : {round(res[3],1)} ans")
        stat.append(f"Age médian des femmes : {round(da.loc['F','median'],1)} ans")
        stat.append(f"Age médian des hommes : {round(da.loc['M','median'],1)} ans")
        stat.append(f"Age maximum des femmes : {round(da.loc['F','max'],1)} ans")
        stat.append(f"Age maximum des hommes : {round(da.loc['M','max'],1)} ans")
        stat.append(f"Age minimum des femmes : {round(da.loc['F','min'],1)} ans")
        stat.append(f"Age minimum des hommes : {round(da.loc['M','min'],1)} ans")
        stat.append(' ')


        stat.append(f'Nombre de membres sympatisants : {self.nbr_membres_sympathisants}')
        stat.append(f'Membres sympatisants : {self.membres_sympathisants}')
        stat.append(' ')
        
        if self.nbr_nouveaux_membres is not None:
            long_str = (f"{self.nbr_nouveaux_membres} "
                         "nouveaux membres de moyenne d'âge de "
                        f"{round(self.moy_age_entrants,1)} ans")
            stat.append(long_str)
            stat.append(f"Liste des nouveaux :\n{self.nouveaux_membres_noms}")
            long_str = (f"{self.nbr_sortants} "
                         "licences non renouvellées de moyenne d'âge de "
                        f"{round(self.moy_age_sortants,1)} ans")
            stat.append(long_str)
            stat.append(f"Liste des sortants :\n{self.sortants_noms}")
        
        if self.rebond:
            stat.append(f'Nombre de rebonds : {len(self.rebond)}')
            rebond_str = ' ; '.join(self.rebond)
            stat.append(f'Membres rebonds : {rebond_str}')
        stat.append(' ')
        
        stat.append("Composition du commité directeur")
        dg,cd_dict =  plot_cd_evolution(self.ctg_path,plot=False)
        dict_cd_year = cd_dict[int(self.year)]
        for typ in dict_cd_year.keys():
            stat.append(f'{typ} : {" ; ".join(dict_cd_year[typ])}')
        stat.append(' ')
        
        mask = self.effectif['Formation Educateur'].str.contains("Animateur Club") 
        mask = [False  if isnan(x) else x for x in mask]
        dg = self.effectif[mask]
        stat.append('Animateur Club')
        stat.append(', '.join(dg['Nom, Prénom'].tolist()))
        stat.append(' ')

        mask = self.effectif['Formation Educateur'].str.contains("Initiateur fédéral") 
        mask = [False  if isnan(x) else x for x in mask]
        dg = self.effectif[mask]
        stat.append('Initiateur fédéral')
        stat.append(', '.join(dg['Nom, Prénom'].tolist()))
        stat.append(' ')

        mask = self.effectif['Formation Educateur'].str.contains("Moniteur Fédéral") 
        mask = [False  if isnan(x) else x for x in mask]
        dg = self.effectif[mask]
        stat.append('Moniteur Fédéral')
        stat.append(', '.join(dg['Nom, Prénom'].tolist()))
        stat.append(' ')

        if 'Pratique VAE' in self.effectif.columns:
            da = self.effectif.groupby(['Sexe','Pratique VAE'])['Nom'].agg(['count'])
            nbr_vae_femme = nbr_femmes-da.loc['F','Non']['count']
            nbr_vae_homme = nbr_hommes-da.loc['M','Non']['count']
            nbr_vae_tot = nbr_vae_femme + nbr_vae_homme
            long_string = ("Nombre de membres équippées de VAE : "
                          f"{nbr_vae_tot} ({round(100*nbr_vae_tot/nbr_membres,1)} %)")
            stat.append(long_string)
            long_string = ("Nombre de femmes équippées de VAE : "
                          f"{nbr_vae_femme} ({round(100*nbr_vae_femme/nbr_femmes,1)} %)")
            stat.append(long_string)
            long_string = ("Nombre d'hommes équippés de VAE: "
                          f"{nbr_vae_homme} ({round(100*nbr_vae_homme/nbr_hommes)} %)")
            stat.append(long_string)

            stat.append(' ')
        if 'Discipline' in self.effectif.columns:            
            da = self.effectif.groupby(['Discipline'])['Nom'].agg('count')
            for pratique in da.index:
                stat.append(f'{pratique} : {da[pratique]}')
            stat.append(' ')

        self.effectif = self.effectif.rename(columns={'\n\t\t\t\tAbonnements':'Abonnements',
                                                  'Revue':'Abonnements'})
        nbr_abonnements = self.effectif.Abonnements.str.contains(r'Revue').sum()
        long_string = (f"Nombre d'abonnés à la revue FFCT : {nbr_abonnements} "
                       f"({round(100*nbr_abonnements/nbr_membres)} %)")
        stat.append(long_string)
        #stat.append(f"\ncotisation licence ffct : {self.cotisation_licence} €")
        #stat.append(f"cotisation totale : {self.cotisation_totale} €")
        #stat.append(f"cotisation CTG : {self.cotisation_ctg} €")
        
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_root = path_effectif  / Path(str(self.year)) 
        path_root = path_root / Path(f'info_effectif_{self.year}.txt')
        with open(path_root,'w') as f:
            f.write('\n'.join(stat))

        stat.append(("\n Ces informations sont disponibles dans le fichier : "
                    f"\n{path_root}"))
        stat ='\n'.join(stat)
        messagebox.showinfo(f'Statistique {self.year}',stat)


        

    def nouveaux_entrants(self):
    
        if self.effectif_1 is not None:
            nouveaux_membres_id = set(self.effectif["N° Licencié"])- \
                                    set(self.effectif_1["N° Licencié"])
            
            
            dg = self.effectif[self.effectif['N° Licencié'].isin(nouveaux_membres_id)]
            moy_age_entrants = dg['Age'].mean() + 1
            
            nouveaux_membres_list = []
            for idx,row in dg.iterrows():
                nouveaux_membres_list.append(f"{row['Prénom'][0]}. {row['Nom']}")
            nouveaux_membres_noms = ' ; '.join(nouveaux_membres_list)
            nbr_nouveaux_membres = len(dg)
            
            return moy_age_entrants, nbr_nouveaux_membres, nouveaux_membres_noms
        else:
            return None, None, None

    def membres_sympathisants(self):
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_root = path_effectif  / Path(str(self.year)) 
        file_path = path_root / Path('membres_sympatisants.xlsx')
        if os.path.isfile(file_path):
            membres_sympathisants_df = pd.read_excel(file_path)
            membres_sympathisants_df['Nom_Prenom'] = membres_sympathisants_df['Nom']+\
                                                     " "+\
                                                     membres_sympathisants_df['Prénom'].str[0]
            membres_sympathisants = ' , '.join(membres_sympathisants_df['Nom_Prenom'].tolist())
            nbr_membres_sympathisants = len(membres_sympathisants_df)
        else:
            nbr_membres_sympathisants = 'inconnu'
            membres_sympathisants = 'inconnu'

        return membres_sympathisants, nbr_membres_sympathisants
        
    def correction_effectif(self):
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_root = path_effectif  / Path(str(self.year)) 
        correction_effectif = pd.read_excel(path_root/Path('correction_effectif.xlsx'))
        correction_effectif.index = correction_effectif['N° Licencié']
        for num_licence in correction_effectif.index:
            idx = self.effectif[self.effectif["N° Licencié"]==num_licence].index
            self.effectif.loc[idx,'Prénom'] = correction_effectif.loc[num_licence,'Prénom']
            self.effectif.loc[idx,'Nom'] = correction_effectif.loc[num_licence,'Nom']
        return self.effectif

    def add_membres_sympathisants(self):
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_root = path_effectif  / Path(str(self.year)) 
        file_path = path_root / Path('membres_sympatisants.xlsx')
        
        if os.path.isfile(file_path):
            membres_sympathisants_df = pd.read_excel(file_path)
            membres_sympathisants_df = membres_sympathisants_df[['N° Licencié',
                                                                 'Nom',
                                                                 'Prénom',
                                                                 'Sexe',
                                                                 'Pratique VAE',
                                                                 'E-mail']]
            
            effectif_tot = pd.concat([self.effectif, membres_sympathisants_df], ignore_index=True, axis=0)
            effectif_tot['Prénom'] = effectif_tot['Prénom'].str.replace(' ','-')

            return effectif_tot
            
        else:
            return self.effectif

    def sortants(self):
        if self.effectif_1 is not None:
            sortants_id = set(self.effectif_1["N° Licencié"]) - set(self.effectif["N° Licencié"])
            
            dg = self.effectif_1[self.effectif_1['N° Licencié'].isin(sortants_id)]
            
            moy_age_sortants = dg['Age'].mean() + 1
            
            sortants_list = []
            for idx,row in dg.iterrows():
                sortants_list.append(f"{row['Prénom'][0]}. {row['Nom']}")
            sortants_noms = ' ; '.join(sortants_list)
            nbr_sortants = len(dg)
            
            return moy_age_sortants, nbr_sortants, sortants_noms
        else:
            return None, None, None

    def cotisation(self):
        cotisation_licence = 'inconnue'
        if 'Cotisation Licence' in self.effectif.columns:
            cotisation_licence = sum(self.effectif['Cotisation Licence'])

        cotisation_totale = 'inconnue'
        if 'Cotisation Totale' in self.effectif.columns:
            cotisation_totale = sum(self.effectif['Cotisation Totale'])

        cotisation_ctg = 15*len(self.effectif)
        return cotisation_licence,cotisation_totale, cotisation_ctg

    def get_rebond(self):
        current_year = datetime.now().year
        path_effectif = self.ctg_path.parent.parent / Path(r"1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
        path_history = path_effectif / Path(str(current_year))
        path_history = path_history / Path('effectif_history.xlsx')
        df = pd.read_excel(path_history)
        df = df.fillna(0)

        year_int = int(self.year)
        list_rebond = []
        if year_int-2 in df.columns:
            for idx,row in df.iterrows():
                if (row[year_int-2] == 0 and row[year_int-1] == year_int-1 and row[year_int] ==0):
                    list_rebond.append('-'.join([row['Nom'],row['Prénom']]))
        return list_rebond

    def plot_histo(self):

        fig, ax = plt.subplots(figsize=(10,4))
        self.effectif['age group'] = pd.cut(self.effectif.Age, bins=range(0, 95, 5), right=False)
        result_hist = self.effectif.groupby('Sexe')['age group']
        result_hist = result_hist.value_counts().unstack().T
        result_hist = result_hist.plot.bar(width=1, stacked=False,ax=ax)

        plt.tick_params(axis='x', labelsize=20)
        plt.tick_params(axis='y', labelsize=20)
        plt.title(self.year,fontsize=20)
        plt.legend()
        plt.tight_layout()
        plt.show()
