__all__ = ['anciennete_au_club',
           'builds_excel_presence_au_club',
           'count_participation',
           'evolution_age_median',
           'evolution_effectif',
           'inscrit_sejour',
           'plot_rebond',
           'read_effectif',
           'read_effectif_corrected',
           'statistique_vae',
           ]

# Standard library imports
import datetime
import functools
import os
import pathlib
import re
import unicodedata
from collections import Counter
from math import asin, cos, radians, sin, sqrt
from pathlib import Path

# 3rd party imports
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Internal imports
from ctg.ctgfuncts.ctg_tools import built_lat_long
from ctg.ctgfuncts.ctg_tools import read_sortie_csv

def read_effectif_corrected(ctg_path:pathlib.WindowsPath, year=None):

    '''Lecture du fichier effectif et correction
    '''


    if year is not None:
        file_effectif = str(year) + '.xlsx'
    else:
        currentDateTime = datetime.datetime.now()
        date = currentDateTime.date()
        year = date.strftime("%Y")
        file_effectif = year + '.xlsx'

    effectif_df = pd.read_excel(ctg_path / Path(str(year))/Path('DATA')/ Path(file_effectif))
    effectif_df = effectif_df[['N° Licencié', 'Nom','Prénom','Sexe','Pratique VAE']]
    path_root = ctg_path / Path(str(year))/Path('DATA')
    correction_effectif = pd.read_excel(path_root/Path('correction_effectif.xlsx'))
    correction_effectif.index = correction_effectif['N° Licencié']
    root_path = ctg_path / Path(str(year))/Path('DATA')
    membres_sympathisants_df = pd.read_excel(root_path/Path('membres_sympatisants.xlsx'))
    membres_sympathisants_df = membres_sympathisants_df[['N° Licencié',
                                                         'Nom',
                                                         'Prénom',
                                                         'Sexe',
                                                         'Pratique VAE']]

    for num_licence in correction_effectif.index:
        idx = effectif_df[effectif_df["N° Licencié"]==num_licence].index
        effectif_df.loc[idx,'Prénom'] = correction_effectif.loc[num_licence,'Prénom']
        effectif_df.loc[idx,'Nom'] = correction_effectif.loc[num_licence,'Nom']

    effectif_df = pd.concat([effectif_df, membres_sympathisants_df], ignore_index=True, axis=0)
    effectif_df['Prénom1'] = effectif_df['Prénom'].str[0]
    effectif_df['Prénom'] = effectif_df['Prénom'].str.replace(' ','-')

    return effectif_df

def inscrit_sejour(file:pathlib.WindowsPath,no_match:list,df_effectif):

    '''builds the DataFrame dg for one event using the csv file of this event.
    The DataFrame dg has 5 columns named :'N° Licencié','Nom','Prénom','Sexe','sejour'
    And EXCEL file is stored in the corresponding EXCEL directory.
    '''

    # used to supress no ascii characters suh as accent cedilla,...
    nfc = functools.partial(unicodedata.normalize,'NFD')
    convert_to_ascii = lambda text : nfc(text). \
                                     encode('ascii', 'ignore'). \
                                     decode('utf-8').\
                                     strip()

    # read the csv file of the event with full path file
    df = read_sortie_csv(file)

    # extract the event type (SEJOUR, SORTIES DU JEUDI,...)
    sejour = os.path.splitext(os.path.basename(file))[0]

    col = ['N° Licencié','Nom','Prénom','Sexe','Pratique VAE','sejour',]
    if df is not None: # file a valid non-empty file
        dg = df[0].str.upper()
        dg = dg.dropna()
        dg = dg.str.replace(' \t?','',regex=False)
        dg = dg.str.replace('.',' ',regex=False)
        dg = dg.str.replace(' - ','-',regex=False)
        dg = dg.apply(convert_to_ascii)
        dg = dg.drop_duplicates()
        dg = dg.str.split('\s{1,10}')

        dg = dg.apply(lambda row : row+[None] if len(row)==2 else row)

        split_dg = pd.DataFrame(dg.tolist(), columns=['name1', 'name2', 'name3'])

        dic = {}
        for idx,row in split_dg.iterrows():
            if (row.name3 is None) and ( row.name2 is not None):
                if len(row.name1)==1:
                    dr = df_effectif.query('Prénom1==@row.name1[0] and Nom==@row.name2')
                    if len(dr):
                        dic[idx] =dr.iloc[0].tolist()[:-1]+[sejour]
                    else:
                        print(f'no match,{row.name2},{row.name1} dans {sejour}')
                        no_match.append((file,row.name2,row.name1))
                elif len(row.name2)==1:
                    dr = df_effectif.query('Prénom1==@row.name2 and Nom==@row.name1')
                    if len(dr):
                         dic[idx] =dr.iloc[0].tolist()[:-1]+[sejour]
                    else:
                        print(f'no match,{row.name2},{row.name1} dans {sejour}')
                        no_match.append((file,row.name2,row.name1))
                else:
                    if len((dr:=df_effectif.query('Prénom==@row.name2 and Nom==@row.name1'))):
                         dic[idx] =dr.iloc[0].tolist()[:-1]+[sejour]
                    elif len((dr:=df_effectif.query('Prénom==@row.name1 and Nom==@row.name2'))):
                         dic[idx] =dr.iloc[0].tolist()[:-1]+[sejour]
                    else:
                        print((f'{sejour} : no match, '
                               f'prénom/prénom:{row.name2}, '
                               f'nom/prénom: {row.name1}'))
                        no_match.append((file,row.name2,row.name1))
            else:
                print((f'WARNING: incorrect name {row.name1}, '
                       f'{row.name2}, {row.name3} in sejour {sejour}'))

        dg = pd.DataFrame.from_dict(dic).T
        if len(dg) !=0:
            dg.columns = col
        else:
            dg = pd.DataFrame([[None,None,None,None,None,sejour,]], columns=col)

    else:
        dg = pd.DataFrame([[None,None,None,None,None,sejour,]], columns=col)

    return dg


def count_participation(path:pathlib.WindowsPath,
                        ctg_path:pathlib.WindowsPath,
                        year:str,
                        info_rando:pd.DataFrame,
                        ):

    '''Creates the DataFrame df_total with 11 columns:
    'N° Licencié', 'Nom', 'Prénom', 'Sexe', 'Pratique VAE', 'sejour',
       'nbr_jours', 'Type', 'Prénom1', 'sexe', 'VAE'
    'sejour' is set to 'aucun' if the member has participate to no event.
    '''

    flag_sejour = False
    if os.path.split(path)[-1] == 'SEJOUR' :
        flag_sejour = True

    type_sortie_default = os.path.basename(path)
    type_sortie_default = type_sortie_default.split('.', 1)[0]

    df_effectif = read_effectif_corrected(ctg_path,year)

    no_match = []
    df_list = []
    info_sejours = []
    sejours = [x for x in os.listdir( path / Path('CSV')) if x.endswith('.csv')]

    sejours_xlsx = [x for x in os.listdir( path / Path('EXCEL')) if x.endswith('.xlsx')]
    for sejours_xlsx in sejours_xlsx:
        path_file_xlsx = path / Path('EXCEL') / Path(sejours_xlsx)
        os.remove(path_file_xlsx)

    nbr_moyen_participants = 0
    counter = 1
    for sejour in sejours:
        dg = inscrit_sejour( path / Path('CSV') /Path(sejour),no_match,df_effectif)

        if re.findall(r'^\d{4}[-_]',sejour):
            date = sejour[2:10].replace('_','-')
        else:
            date = str(year)[2:4] + '-' +sejour[0:5].replace('_','-')

        dg['nbr_jours'] = 0
        if date in info_rando['date'].tolist():
            dh = info_rando.query('type=="randonnee" and date==@date')
            if not flag_sejour and len(dh) != 0 :
                dg['Type'] = "RANDONNEE"
            else:
                dg['Type'] = type_sortie_default

            dg['nbr_jours'] = 1
            dh = info_rando.query('type=="sejour" and date==@date')
            dg['cout_sejour'] = 0
            if flag_sejour and len(dh) != 0:
                dg['nbr_jours'] = dh['nbr_jours'].tolist()[0]
                dg['cout_sejour'] = 0
                if type_sortie_default == 'SEJOUR':
                    dg['cout_sejour'] = dh['Cout'].tolist()[0]


        if dg.reset_index().loc[0,'Nom'] is not None:
            nbr_inscrits = len(dg)
            if nbr_inscrits != 0:
                nbr_moyen_participants = nbr_moyen_participants +\
                                        (nbr_inscrits - nbr_moyen_participants)/counter
                counter += 1
                long_string = (f'{os.path.split(path)[-1]} :{sejour}, '
                               f"Nombre d'inscrits : {nbr_inscrits}")
                info_sejours.append(long_string)

        df_list.append(dg)

        # Store ae an EXCEL file
        file_store = os.path.splitext(sejour)[0]+'.xlsx'
        dg.to_excel(path / Path('EXCEL') / Path(file_store))

    long_string = ("Nombre d'évènenements : "
                  f"{counter-1}. Nombre moyen de participants : {nbr_moyen_participants}")
    info_sejours.append(long_string)
    info_sejours = '\n'.join(info_sejours)
    info_sejours_path = ctg_path / Path(str(year))
    info_sejours_path = info_sejours_path / Path('STATISTIQUES') / Path('TEXT')
    info_sejours_path = info_sejours_path / Path(type_sortie_default+'.txt')
    with open(info_sejours_path,'w') as f:
        f.write(info_sejours)

    nbr_evenement = counter-1
    if counter-1 > 0 :  #
        df_total = pd.concat(df_list,ignore_index=True)
    else:
        return (None, None, None)


    liste_licence = df_effectif['N° Licencié']
    liste_licence_sejour = df_total['N° Licencié']
    index = list(set(liste_licence)-set(liste_licence_sejour))

    # take care of the member with no participation to the events
    df_non_inscrits = df_effectif.copy()
    df_non_inscrits = df_non_inscrits[df_non_inscrits['N° Licencié'].isin(index)]
    df_non_inscrits['sejour'] = 'aucun'
    df_total = pd.concat([df_total,df_non_inscrits],ignore_index=True)


    return(no_match,df_total,index)


def read_effectif(ctg_path:pathlib.WindowsPath,year:str)->pd.DataFrame:

    def distance_(row)->float:

        phi1, lon1 = dh.query("Ville=='GRENOBLE'")[['long','lat']].values.flatten()
        phi1, lon1 = radians(phi1), radians(lon1)
        phi2, lon2 = radians(row['long']), radians(row['lat'])
        rad = 6371
        dist = 2 * rad * asin(
                                sqrt(
                                    sin((phi2 - phi1) / 2) ** 2
                                    + cos(phi1) * cos(phi2) * sin((lon2 - lon1) / 2) ** 2
                                ))
        return np.round(dist,1)


    df = pd.read_excel(ctg_path / Path(str(year))/ Path('DATA')/Path(str(year)+'.xlsx'))

    df['Date de naissance'] = pd.to_datetime(df['Date de naissance'], format="%d/%m/%Y")

    df['Age']  = df['Date de naissance'].apply(lambda x : (pd.Timestamp(year, 9, 30)-x).days/365)

    dh = built_lat_long(df)

    df['distance'] = df.apply(distance_,axis=1)

    return df

def evolution_effectif(ctg_path:pathlib.WindowsPath):

    # Evolution des effectifs hommes et femme de 2016 à 2022

    def addlabels(x,y,z=None):
        for i in range(len(x)):
            if z is None:
                plt.text(x[i]-0.4,y[i]+5,y[i],size=10)
            else:
                plt.text(x[i]-0.4,y[i]+5,z[i],size=10,rotation='vertical')


    def _evolution_effectif(years):
        nbr_hommes = []
        nbr_femmes = []
        nbr_total = []
        ratio_femmes = []
        for year in years:
            if year == 2006:
                nbr_femmes.append(41)
                nbr_hommes.append(100)
                nbr_total.append(141)
                ratio_femmes.append(str(29)+'%')
            else:
                df_effectif = read_effectif(ctg_path,year)
                nh = len(df_effectif.query('Sexe =="M"'))
                nbr_hommes.append(nh)
                nf = len(df_effectif.query('Sexe =="F"'))
                nbr_femmes.append(nf)
                nbr_total.append(nh + nf)
                ratio_femmes.append(str(int(100*nf/(nh+nf)))+'%')
        return nbr_hommes, nbr_femmes,nbr_total,ratio_femmes

    years = [2006] + [int(x) for x in os.listdir(ctg_path) if re.findall('^\d{4}$',x)]
    nbr_hommes, nbr_femmes,nbr_total,ratio_femmes = _evolution_effectif(years)

    title='Evolution du nombre de membres CTG'

    ax = plt.axes()
    plt.bar(years, nbr_femmes,label= 'Femme')
    addlabels(years, nbr_femmes)
    plt.bar(years, nbr_hommes,bottom=nbr_femmes,label= 'Hommes')
    plt.legend()
    addlabels(years, nbr_total)
    ax.set_xticks(years)
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.show()

def evolution_age_median(ctg_path:pathlib.WindowsPath):

    def addlabels(x,y,z=None):
        for i in range(len(x)):
            if z is None:
                plt.text(x[i]-0.3,y[i]+0.1,y[i],size=10)
            else:
                plt.text(x[i]-0.2,y[i]+0.1,z[i],size=10,rotation='vertical')

    xt = range(55,90)
    years = [int(x) for x in os.listdir(ctg_path) if re.findall('^\d{4}$',x)]
    nbr_femmes = []
    ratio_femmes = []
    nbr_total = []
    age_mean = []
    age_naturel = []
    for idx,year in enumerate(years):
        file = ctg_path / Path(str(year)) / Path('DATA') / Path(str(year)+'.xlsx')
        df_effectif = pd.read_excel(file)
        df_effectif['Date de naissance'] = pd.to_datetime(df_effectif['Date de naissance'],
                                                                      format="%d/%m/%Y")
        df_effectif['Age']  = df_effectif['Date de naissance'].apply(lambda x :
                                                        (pd.Timestamp(year, 9, 30)-x).days/365)
        age_median = df_effectif['Age'].median()
        age_mean.append(age_median)
        if idx == 0:
           age_median_0 = age_median
           age_naturel.append(age_median_0)
        else:
           age_naturel.append(age_median_0+idx)

    fig, ax = plt.subplots()
    plt.bar(years,age_mean)
    plt.plot(years,age_naturel,"--r")
    plt.ylabel('Age moyen')
    plt.ylim(50,1.1*max(age_naturel))
    addlabels(years,[round(x,1) for x in age_mean])
    linear_model = np.polyfit(years,age_mean,1)
    linear_model_fn = np.poly1d(linear_model)
    plt.xticks(years,years)
    plt.tick_params(axis='x', labelsize=20,rotation=90)
    plt.tick_params(axis='y', labelsize=20)
    plt.ylabel('Age median',size=20)
    x_s =[years[0]-1] + years + [years[-1]+1]
    plt.plot(x_s,linear_model_fn(x_s),"--g")
    plt.title(f'Pente vieillissement : {round(linear_model[0]*12,1)} mois par ans')
    plt.tight_layout()
    plt.show()

def builds_excel_presence_au_club(ctg_path):

    list_date = [int(x) for x in os.listdir(ctg_path) if re.findall('^\d{4}$',x)]
    list_df = []

    for date in list_date:
        df = pd.read_excel(ctg_path /Path(str(date)) / Path('DATA') / Path(str(date)+'.xlsx'))
        df['date'] = date
        list_df.append(df[['N° Licencié','Nom','Prénom','date']])

    df = pd.concat(list_df) #.to_excel(file / Path('effectif_total.xlsx'),index=False)

    dic = {}
    list_num_licence = []
    list_nom = []
    list_prenom = []
    list_date_ = []
    for licence in df.groupby('N° Licencié'):
        list_c = [x if x in licence[1]['date'].to_list() else None for x in list_date]
        #if len(list_c) - list_c.count(None) == 1: singleton
        list_num_licence.append(licence[0])
        list_nom.append(licence[1]['Nom'].unique()[0])
        list_prenom.append(licence[1]['Prénom'].unique()[0])
        list_date_.append(list_c)


    dic['N° Licencié'] = list_num_licence
    dic['Nom'] = list_nom
    dic['Prénom'] = list_prenom
    dic['date'] = list_date_

    df = pd.DataFrame.from_dict(dic)
    split_df = pd.DataFrame(df['date'].tolist(), columns=list_date)

    df = pd.concat([df, split_df], axis=1)
    df = df.drop('date',axis=1)
    out_path = ctg_path / Path(str(list_date[-1]))
    out_path = out_path / Path('STATISTIQUES') / Path('EXCEL') / Path('effectif_history.xlsx')
    df.to_excel(out_path)
    return out_path

def statistique_vae(ctg_path):

    # function to add value labels
    def addlabels(x,y):
        for i in range(1,len(x)):
            plt.text(x[i]-0.2,y[i]+0.5,y[i],size=10)

    current_year = datetime.datetime.now().year

    last_year = 2018
    years =[]
    nb_vae_m = []
    nb_vae_f = []
    nb_vae_tot = []
    for year in range(last_year,current_year+1):
        df_N1 = pd.read_excel(ctg_path / Path(str(year)) / Path('DATA') /Path(str(year)+'.xlsx'))
        years.append(year)
        nb_vae_m.append(sum((df_N1['Pratique VAE'] == 'Oui') & (df_N1['Sexe'] == 'M')))
        nb_vae_f.append(sum((df_N1['Pratique VAE'] == 'Oui') & (df_N1['Sexe'] == 'F')))
        nb_vae_tot.append(sum((df_N1['Pratique VAE'] == 'Oui') & (df_N1['Sexe'] == 'F')) +
                          sum((df_N1['Pratique VAE'] == 'Oui') & (df_N1['Sexe'] == 'M')))
    plt.bar(years, nb_vae_m,label= 'Homme')
    plt.bar(years, nb_vae_f,bottom=nb_vae_m,label= 'Femme')
    plt.tick_params(axis='x', labelsize=15)
    plt.tick_params(axis='y', labelsize=15)
    plt.ylabel('nombre de VAEs',size=15)
    plt.xticks(rotation=90)
    plt.legend()
    plt.ylabel('nombre de VAE')
    addlabels(years, nb_vae_m)
    addlabels(years, [x+y for x,y in zip(nb_vae_m,nb_vae_f)])
    plt.tight_layout()
    plt.show()

def anciennete_au_club(ctg_path):

    def addlabels(x,y,offset):
        for i in range(len(x)):
            if y[i] != 0:
                plt.text(x[i]-0.2,y[i]+offset,round(y[i],1),size=15)

    currentDateTime = datetime.datetime.now()
    date = currentDateTime.date()
    current_year = int(date.strftime("%Y"))

    in_path = ctg_path / Path(str(current_year))
    in_path = in_path / Path('STATISTIQUES') /Path('EXCEL') / Path('effectif_history.xlsx')
    df = pd.read_excel(in_path)
    eff = []

    years = list(range(2012,current_year+1))
    for year in years:
        dg = df.dropna(subset=[year,current_year])
        eff.append(len(dg))
    eff = [eff[0]] + list(np.diff(eff))

    # creating the bar plot
    fig = plt.figure(figsize = (10, 5))
    plt.bar(years, eff, color ='maroon',
            width = 0.4)

    plt.xlabel("")
    plt.ylabel("# adhérents")
    plt.title("Ancienneté au CTG")
    addlabels(list(range(2012,2025)), eff,0)
    plt.tight_layout()
    plt.show()


def plot_rebond(ctg_path:pathlib.WindowsPath):

    '''
    '''
    
    def addlabels(x,y):
        for i in range(len(x)):
            ax[0].text(i-0.2,-23,y[i],size=10)


    current_year = int(datetime.datetime.now().year)
    file_path = ctg_path / Path(str(current_year)) / Path('STATISTIQUES') / Path('EXCEL')
    file_path = file_path / Path('effectif_history.xlsx')
    if not os.path.isfile(file_path):
        builds_excel_presence_au_club(ctg_path)

    years_list = [int(x) for x in os.listdir(ctg_path) if re.findall('^\d{4}$',x)]
    year_dep = min(years_list)+2
        
    df = pd.read_excel(file_path)
    df = df.fillna(0)
    dic = {}
    years = range(year_dep,current_year+1)
    for year in years:
        list_rebond = []
        list_entrant = []
        list_sortant = []
        for idx,row in df.iterrows():
            if (row[year-2] == 0 and row[year-1] == year-1 and row[year] ==0):
                #print(year,row['Nom'],row['Prénom'],row['N° Licencié'])
                list_rebond.append('-'.join([row['Nom'],row['Prénom']]))
            if row[year-1] == 0 and row[year] == year:
                list_entrant.append('-'.join([row['Nom'],row['Prénom']]))
            if row[year] == 0 and row[year-1] == year-1:
                list_sortant.append('-'.join([row['Nom'],row['Prénom']]))
        dic[year]=[len(list_rebond),'; '.join(list_rebond),
                   len(list_entrant),'; '.join(list_entrant),
                   -len(list_sortant),'; '.join(list_sortant)]
        
    dg = pd.DataFrame.from_dict(dic).T
    dg.columns=['# rebonds',
                'Nom rebonds',
                '# entrants',
                'Nom entrants',
                '# sortants',
                'Nom sortants']
    rebond_pourcent = [None] + [round(100*x[0]/x[1],1) for x in 
                       zip(dg['# rebonds'].tolist()[1:], dg['# entrants'].tolist()[:-1])]
    dg['% rebond'] = rebond_pourcent
    file_rebond = ctg_path / Path(str(current_year)) / Path('STATISTIQUES')
    file_rebond = file_rebond / Path('EXCEL') / Path('rebond.xlsx')
    dg.to_excel(file_rebond)
    year_pourcent = list(years[1:])
    
    fig, ax = plt.subplots(nrows=1, ncols=2)
    dg[['# entrants','# sortants']].plot.bar(stacked=True,ax=ax[0])
    addlabels(dg.index.tolist(),(dg['# entrants']+dg['# sortants']).tolist())
    ax[0].yaxis.grid()
    ax[0].set_ylabel('# membres')
    ax[0].set_ylim((-25,40))
    ax[0].legend(loc='upper center',ncol=1)
    
    dg[['# rebonds']].plot.bar(ax=ax[1])
    ax[1].yaxis.grid()
    plt.show()