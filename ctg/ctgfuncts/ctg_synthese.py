__all__ =['evolution_sorties',
          'nbr_sejours_adherent',
          'plot_pie_synthese',
          'stat_cout_sejour',
          'synthese',
          'synthese_adherent',
          'synthese_randonnee',
          ]

# Standard library imports
import datetime
import pathlib
import os
import os.path
from collections import Counter
from collections import namedtuple
from pathlib import Path
from tkinter import messagebox
from typing import Optional

# Third party imports
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import yaml

# Internal imports
from ctg.ctggui.guiglobals import ACTIVITE_LIST
from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctgfuncts.ctg_tools import get_sejour_info
from ctg.ctgfuncts.ctg_tools import get_cout_total
from ctg.ctgfuncts.ctg_tools import normalize_tag

def synthese(year:str,ctg_path:pathlib.WindowsPath)->None:

    path_dir_list = [ctg_path / Path(year) / Path('SORTIES DU SAMEDI') / Path('EXCEL'),
                     ctg_path / Path(year) / Path('SORTIES DU DIMANCHE') / Path('EXCEL'),
                     ctg_path / Path(year) / Path('SORTIES DU JEUDI') / Path('EXCEL'),
                     ctg_path / Path(year) / Path('SORTIES HIVER') / Path('EXCEL'),
                     ctg_path / Path(year) / Path('SORTIES DU LUNDI')/ Path('EXCEL'),
                     ctg_path / Path(year) / Path('SEJOUR') / Path('EXCEL')]

    list_df = []
    for path_dir in path_dir_list:
        if os.path.isdir(path_dir):
            files = [x for x in os.listdir(path_dir) if x.endswith('.xlsx') and "~$" not in x]
            list_df.extend([pd.read_excel(path_dir / Path(file), engine='openpyxl')
                           for file in files])

    df_total = pd.concat(list_df, ignore_index=True)

    df_total['Pratique VAE'].fillna('Non',inplace=True)

    # nombre moyen de participant par activité
    for x in df_total.groupby(['Type']):
        print(x[0],len(x[1]),len(set(x[1]['sejour'])),len(x[1])/len(set(x[1]['sejour'])))

    file = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL') / Path('synthese.xlsx')
    df_total.to_excel(file)

def plot_pie_synthese(year:str,ctg_path:pathlib.WindowsPath,mode: Optional[bool] = False)->None:

    '''Plot from the EXCEL file `synthese.xlsx` the pie plot of 
    the number of participation to the evenments'''

    def func(pct, allvalues):
        absolute = round(pct / 100.*np.sum(allvalues),0)
        #return "{:.1f}%\n({:d})".format(pct, absolute)
        label = f"{int(round(absolute,1))}\n{round(pct,1)} %"
        return label
    
    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL') / Path('synthese.xlsx')
    df_total = pd.read_excel(file_in)
    df_total = df_total.dropna(subset=['Type'])
    df_total = df_total.dropna(subset=['Nom'])
    
    if mode: # un sejour vaut pour un jour 
        df_total['nbr_jours'] = 1

    dagg = df_total.groupby('Type')['nbr_jours'].agg('sum')



    explode_dic = {'RANDONNEE':0.1,
                   'SEJOUR':0.0,
                   'SORTIES DU DIMANCHE':0.2,
                   'SORTIES DU JEUDI':0.2,
                   'SORTIES DU LUNDI':0.2,
                   'SORTIES DU SAMEDI':0.2,
                   'SORTIES HIVER':0.2}

    data = []
    sorties = []
    for type_sejour, nbr in zip(dagg.index,dagg.tolist()):
        if nbr!=0:
            data.append(nbr)
            sorties.append(type_sejour)

    explode = [explode_dic[typ] for typ in sorties]


    # Creating plot
    #fig = plt.figure(figsize =(10, 7))
    _, _, autotexts = plt.pie(data,
                              labels = sorties,
                              autopct = lambda pct: func(pct, data),
                              explode = explode,
                              textprops={'fontsize': 10})
                              
    title = f'{year} (Nombre total de participations: {sum(data)})'
    plt.title(title, pad=20, fontsize=12)

    _ = plt.setp(autotexts, **{'color':'k', 'weight':'bold', 'fontsize':10})

    plt.tight_layout()

    fig_file = 'SORTIES_PIE.png'
    plt.savefig(ctg_path / Path(year) / Path('STATISTIQUES/IMAGE') / Path(fig_file),bbox_inches='tight')
    plt.show()


def synthese_adherent(year:str,ctg_path:pathlib.WindowsPath):

    '''generates the EXCEL file 'synthese_adherent.xlsx' with 11 columns :'Nom','Prénom',
    'SORTIE DU DIMANCHE CLUB','SORTIE DU SAMEDI CLUB','SORTIE DU JEUDI CLUB',
    'RANDONNEE','SEJOUR-JOUR','Nbr_SEJOURS', 'SORTIE HIVER','TOTAL','COUT_SEJOUR',
    '''

    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL') / Path('synthese.xlsx')
    df_total = pd.read_excel(file_in)
    df_total = df_total.dropna(subset=['Type'])
    nbre = {}

    for id_licence, dg in df_total.groupby('N° Licencié'):

        nbre[id_licence]=[dg['Nom'].unique()[0],dg['Prénom'].unique()[0]]

        nb_sortie_dimanche = len(dg.query("Type.str.contains('SORTIES DU DIMANCHE')"))
        nbre[id_licence] = nbre[id_licence] + [nb_sortie_dimanche]

        nb_sortie_samedi = len(dg.query("Type.str.contains('SORTIES DU SAMEDI')"))
        nbre[id_licence] = nbre[id_licence] + [nb_sortie_samedi]

        nb_sortie_jeudi = len(dg.query("Type.str.contains('SORTIES DU JEUDI')"))
        nbre[id_licence] = nbre[id_licence] + [nb_sortie_jeudi]

        nb_rando = len(dg.query("Type.str.contains('RANDONNEE')"))
        nbre[id_licence] = nbre[id_licence] + [nb_rando]

        nb_sejour_jours = sum(dg.query("Type.str.contains('SEJOUR')")['nbr_jours'].tolist())
        nbre[id_licence] = nbre[id_licence] + [nb_sejour_jours]

        nb_sejour = len(dg.query("Type.str.contains('SEJOUR')")['sejour'].unique())
        nbre[id_licence] = nbre[id_licence] + [nb_sejour]

        nb_hiver = len(dg.query("Type.str.contains('SORTIES HIVER')"))
        nbre[id_licence] = nbre[id_licence] + [nb_hiver]

        nbr_evenements = [nb_sortie_dimanche +
                          nb_sortie_samedi +
                          nb_sortie_jeudi +
                          nb_rando +
                          nb_sejour_jours +
                          nb_hiver]
        nbre[id_licence] = nbre[id_licence] + nbr_evenements

        dg['cout_sejour'] = dg['cout_sejour'].fillna(0)
        nbre[id_licence] = nbre[id_licence] + [sum(dg['cout_sejour'])]


    dg = pd.DataFrame.from_dict(nbre).T
    dg.columns = [
                 'Nom',
                 'Prénom',
                 'SORTIE DU DIMANCHE CLUB',
                 'SORTIE DU SAMEDI CLUB',
                 'SORTIE DU JEUDI CLUB',
                 'RANDONNEE',
                 'SEJOUR-JOUR',
                 'Nbr_SEJOURS',
                 'SORTIE HIVER',
                 'TOTAL',
                 'COUT_SEJOUR',
                 ]

    effectif = EffectifCtg(year,ctg_path)
    df_effectif = effectif.effectif
    df_effectif.index = df_effectif['N° Licencié']
    orphan = set(df_effectif.index) - set(dg.index)
    df_orphan = df_effectif.loc[list(orphan)][['Nom','Prénom']]
    df_orphan[['SORTIE DU DIMANCHE CLUB',
               'SORTIE DU SAMEDI CLUB',
               'SORTIE DU JEUDI CLUB',
               'RANDONNEE',
               'SEJOUR-JOUR',
               'Nbr_SEJOURS',
               'SORTIE HIVER',
               'TOTAL',
               'COUT_SEJOUR',]] = [0,0,0,0,0,0,0,0,0]

    dg = pd.concat([dg, df_orphan], axis=0)

    file_out = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL')
    file_out = file_out / Path('synthese_adherent.xlsx')
    dg.to_excel(file_out)

def synthese_randonnee(year:str,ctg_path:pathlib.WindowsPath,type_sejour:str):

    '''Creates a plot synthetizing the participation at th randonnées.
    '''

    def addlabels(x,y,offset,y_val=None):
        for i in range(len(x)):
            if y[i] != 0:
                if y_val is None:
                    plt.text(x[i]-0.25,y[i]+offset,round(y[i],1),size=10)
                else:
                    plt.text(x[i]-0.25,y[i]+offset,round(y_val[i],1),size=10)


    file_in = Path(ctg_path) / Path(year) / Path('STATISTIQUES')
    file_in = file_in / Path('EXCEL') / Path('synthese.xlsx')
    df_total = pd.read_excel(file_in)
    df_total = df_total.dropna(subset=['Nom'])
    df_total = df_total.query('Type==@type_sejour')
    dg = df_total.groupby('sejour').agg('count')['N° Licencié']
    
    file_info = Path(ctg_path) / Path(year) / Path('DATA') / Path('info_randos.xlsx')
    info_df = pd.read_excel(file_info)
    type_sejour_m = type_sejour.lower()
    info_df=info_df.query('type==@type_sejour_m')
    info_dic = dict(zip(info_df['date'],info_df['name_activite']))
    info_duree = dict(zip(info_df['date'],info_df['nbr_jours']))    
    
    tag_list = [normalize_tag(x,year) for x in dg.index]
    labelx = [f'{k[3:8]} {info_dic[k].strip()}' for k in tag_list]
    duree_sejour = [info_duree[k] for k in tag_list]

    cout_total_rando = get_cout_total(year,type_sejour.lower(),dg,ctg_path)

    fig = plt.figure()
    plt.bar(range(len(dg)),dg.tolist())
    addlabels(list(range(len(dg))),dg.tolist(),0.2)
    plt.xticks(range(len(dg)), labelx, rotation='vertical')
    plt.tick_params(axis='x', rotation=90, labelsize=10)
    plt.tick_params(axis='y',labelsize=10)

    if type_sejour == 'RANDONNEE':
        long_string = (f'Année : {year}\n'
                       f'# randos : {len(dg)} , '
                       f'# participants : '
                       f'{sum(dg)}, Coût : {cout_total_rando} €')
        _ = plt.title(long_string)
    else:
        y = [1] *  len(duree_sejour)
        addlabels(list(range(len(dg))),y,1,duree_sejour)
        long_string = (f'Année : {year}\n'
                       f'# jours : {sum(duree_sejour)} '
                       f'# sejours : {len(dg)} , '
                       f'# participants : {sum(dg)}'
                       f', Coût : {cout_total_rando} €')
        _ = plt.title(long_string)
    plt.tight_layout()
    plt.show()

def nbr_sejours_adherent(year:str, ctg_path:pathlib.WindowsPath):

    '''Generates the histgramm of the number of sejour versus the number of members
    '''

    plt.style.use('ggplot')

    # function to add value labels
    def addlabels(x,y):
        for i in range(len(x)):
            plt.text(x[i]-0.2,y[i]+1,y[i],size=label_size)
    label_size = 15
    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL')
    file_in = file_in / Path('synthese_adherent.xlsx')
    df_total = pd.read_excel(file_in)

    c = Counter()
    c = Counter(df_total['Nbr_SEJOURS'].tolist())
    c = c.most_common()

    x, y = zip(*c)
    x = list(x)
    y = list(y)

    fig, ax = plt.subplots(figsize=(5,5))
    plt.bar([str(x_s) for x_s in x],y)
    plt.xlabel('Nombre de participation à des séjours')
    plt.ylabel('Nombre de licenciers')
    plt.tick_params(axis='x', labelsize=label_size)
    plt.tick_params(axis='y', labelsize=label_size)
    plt.title(year,pad=20, fontsize=20)

    ax.set_xlabel('N séjours', fontsize = label_size)
    ax.set_ylabel('Nombre de CTG ayant \n participé à N séjours', fontsize = label_size)

    x.sort()
    addlabels(x,y)
    plt.tight_layout()
    plt.show()

    fig_file = 'SEJOURS_STAT_PARTICIPATION.png'
    plt.savefig(ctg_path / Path(year) / Path('STATISTIQUES/IMAGE') / Path(fig_file),bbox_inches='tight')

def _read_memory_sorties()->dict:

    '''Reads the default PVcharacterization.yaml config file'''

    parent = Path(__file__).parent.parent
    path_config_file = parent  / Path('ctgfuncts/CTG_RefFiles') / Path('memory_sorties.yml')
    with open(path_config_file) as file:
        memory = yaml.safe_load(file)

    return memory

def evolution_sorties(type:str,ctg_path:pathlib.WindowsPath):

    def add_memory(stat_dic,years):
        
        '''Add years from 2014 to 2021to the dic `stat_year`. These statistics are stored
        in the package and can only be modified by the package owner by pulling a request at
        https://github.com/Bertin-fap/ctgutils.
        '''
        
        memory = _read_memory_sorties()


        for year,v in memory['memory'].items():

            stat_dic[str(year)] = statyear(v['PARTICIPATION_SEJOURS'], # participants sejour
                                           v['Nombre_sorties_sejour'], # jours_sejour
                                           v['SORTIES_CLUB_DIMANCHE'], # sortie_dimanche_club
                                           v['SORTIES_CLUB_SAMEDI'],   # sortie_samedi_club
                                           v['SORTIES_HIVER'],         # sortie_hiver_club
                                           v['SORTIES_CLUB_JEUDI'],    # sortie_jeudi_club
                                           v['RANDONNEES'],            # randonnee
                                           v['Nombre_sejours'],        # nbr_sejours
                                           v['Nombre_jours_sejour'],)  # nbr_jours_sejours

            years.append(str(year))

    def fill_stat_year(year:str):
    
        '''Builds the dict stat_year using the EXCEL file 
        '''

        sejour_info = get_sejour_info(ctg_path,year)

        file_in = ctg_path / Path(str(year)) / Path('STATISTIQUES') / Path('EXCEL')
        file_in = file_in / Path('synthese_adherent.xlsx')

        df = pd.read_excel(file_in)

        stat_year = statyear(df['Nbr_SEJOURS'].sum(),             # nbr_jours_participation_sejours
                             df['SEJOUR-JOUR'].sum(),             # nbr_participations_sejour
                             df['SORTIE DU DIMANCHE CLUB'].sum(), # sortie_dimanche_club
                             df['SORTIE DU SAMEDI CLUB'].sum(),   # sortie_samedi_club
                             df['SORTIE HIVER'].sum(),            # sortie_hiver_club
                             df['SORTIE DU JEUDI CLUB'].sum(),    # sortie_jeudi_club
                             df['RANDONNEE'].sum(),               # randonnee
                             sejour_info.nbr_sejours,             # nbr_sejours
                             sejour_info.nbr_jours)               # nbr_jours_sejours

        return stat_year

    def addlabels(x,y,offset):
        for i in range(len(x)):
            plt.text(x[i],y[i]+offset,round(y[i],1),size=10)

    def plot_stat(years,nb_participants,title,label_y):
        plt.figure(figsize=(8, 6))
        colors = ['#fdaa48']
        size_label = 20
        plt.bar(years,nb_participants,color=colors)
        plt.ylabel(label_y,size=size_label)
        addlabels(years,nb_participants,1)
        plt.xticks(rotation=90)
        plt.tick_params(axis='x', labelsize=size_label)
        plt.tick_params(axis='y', labelsize=size_label)
        plt.title(title,fontsize=18)
        plt.tight_layout()
        plt.show()

    statyear = namedtuple('activite', ACTIVITE_LIST)

    plt.style.use('ggplot')
    years = []
    stat_dic = {}
    add_memory(stat_dic,years)

    today = datetime.datetime.now()
    years_new = [str(year) for year in range(2022,today.year+1)]
    
    for year in years_new:
        stat_dic[year] = fill_stat_year(year)

    years = years + years_new


    if type == 'nbr_jours_participation_sejours':
        plot_stat(years,
                  [stat_dic[year].nbr_jours_participation_sejours for year in years],
                  type,
                  type)
    elif  type == 'sortie_dimanche_club':
        plot_stat(years,
                  [stat_dic[year].sortie_dimanche_club for year in years],
                  type,
                  '#participants')
    elif  type == 'sortie_samedi_club':
        plot_stat(years,
                  [stat_dic[year].sortie_samedi_club for year in years],
                  type,
                  '#participants')
    elif  type == 'sortie_jeudi_club':
        plot_stat(years,
                  [stat_dic[year].sortie_jeudi_club for year in years],
                  type,
                  '#participants')
    elif  type == 'randonnee':
        plot_stat(years,
                  [stat_dic[year].randonnee for year in years],
                  type,
                  '#participants')
    elif  type == 'nbr_participations_sejours':
        plot_stat(years,
                  [stat_dic[year].nbr_participations_sejours for year in years],
                  'Nombre de participations aux séjours',
                  '# participations aux séjours')
    elif  type == 'nbr_sejours':
        plot_stat(years,
                  [stat_dic[year].nbr_sejours for year in years],
                  'Nombre de séjours',
                  '# séjours')
    elif  type == 'nbr_jours_sejours':
        plot_stat(years,
                  [stat_dic[year].nbr_jours_sejours for year in years],
                  'Nombre de jours séjours',
                  '# jours séjour')
    elif type == 'synthese':
        plot_synthese_sortie(stat_dic)
        

def stat_cout_sejour(year:str,ctg_path:pathlib.WindowsPath)->None:

    '''Builds the histogramm of the number m of members whose have take part to n sejours.
    '''
    FONT_SIZE = 12
    LABEL_SIZE = 10
    def plot_histo(col_name,idx_col,idx_row,labelx,labely,unit,bins):
        df.hist(column=col_name,
                bins=bins,
                ax=ax[idx_col,idx_row],
                xlabelsize=LABEL_SIZE,
                ylabelsize=LABEL_SIZE,
                color='aqua',)
        ax[idx_col,idx_row].set_xlabel(labelx,fontsize=FONT_SIZE)
        ax[idx_col,idx_row].set_ylabel(labely,fontsize=FONT_SIZE)
        ax[idx_col,idx_row].set_title('',fontsize=FONT_SIZE)

        col_without_zero = [x for x in df[col_name] if x>0]

        mean_col = round(np.mean(df[col_name]),1)
        mean_col_without_zero = round(np.mean(col_without_zero),1)
        med_col = round(np.median(df[col_name]),1)
        med_col_without_zero = round(np.median(col_without_zero),1)
        long_string = (f'     Total : {sum(df[col_name])} {unit}\n'
                       f'     Moyenne : {mean_col} {unit}\n'
                       f'     Mediane : {med_col} {unit}\n'
                       'Sans prendre en compte la classe 0 :\n'
                       f'     Moyenne : {mean_col_without_zero} {unit}\n'
                       f'     Mediane : {med_col_without_zero} {unit}\n\n')
        return long_string

    def lorentz(col,num_col,num_row,labely):
        arr = np.array([x for x in df[col].tolist() if x>-1])
        arr.sort()
        
        X_lorenz = 100*arr.cumsum() / arr.sum()
        
        ax[num_col,num_row].scatter(np.arange(X_lorenz.size), X_lorenz, 
                   marker='o', color='darkgreen', s=10)
        ax[num_col,num_row].set_ylabel(labely,fontsize=FONT_SIZE)
        ax[num_col,num_row].set_xlabel('# adhérents',fontsize=FONT_SIZE)
        ax[num_col,num_row].grid(True)
        

    fig, ax = plt.subplots(nrows=2, ncols=2, sharey=False)

    file = ctg_path / Path(year) / Path('STATISTIQUES') / Path('EXCEL') 
    file = file / Path('synthese_adherent.xlsx')
    if not os.path.isfile(file):
        synthese_adherent(year,ctg_path)

    df = pd.read_excel(file)

    file = ctg_path / Path(year) / Path('DATA') / Path('info_randos.xlsx')
    info_df = pd.read_excel(file)
    nbr_sejours = info_df.query('type=="sejour"')['date'].count()
    nbr_jours = sum([x for x in info_df["nbr_jours"] if x>1])

    comment = f'Année : {year}\n'
    comment += f'Nombre de séjours : {nbr_sejours}\n'
    comment += 'Analyse du coût des séjours :\n'
    comment_cout = plot_histo('COUT_SEJOUR',0,0, 'Coût séjour €','# membres','€',[0,50,600,1200,1800,2400,3000])
    comment += f'     Cout annuel des séjours : {sum(info_df["Cout"].fillna(0))} €\n'
    comment += comment_cout
    comment += 'Analyse de la durée des séjours :\n'
    comment += f'     Durée totale des séjours : {nbr_jours} jours\n'
    comment_jour = plot_histo('SEJOUR-JOUR',0,1,'# jours séjour','','jours',[0,1,5,10,15,20,25,30,35,40])
    comment += comment_jour

    lorentz("COUT_SEJOUR",1,0,'%')
    lorentz("SEJOUR-JOUR",1,1,'')
    
    fig.suptitle(f'Année : {str(year)}',fontsize=20)
    plt.tight_layout()
    
    fig.suptitle(f'Année : {str(year)}',fontsize=16)
    plt.tight_layout()
    plt.show()
    
    file = ctg_path / Path(year) / Path('STATISTIQUES') / Path('TEXT')
    file = file / Path('synthese_sejour.txt')
    with open(file,'w', encoding='utf-8') as f:
        f.write(comment)

def plot_synthese_sortie(stat_dic:dict):
    
    '''Synthetic plot of the events
    '''
    df = pd.DataFrame.from_dict(stat_dic).T
    df.columns = ['PARTICIPATION_SEJOURS',
                  'Nombre_sorties_sejour',
                  'SORTIES_CLUB_DIMANCHE',
                  'SORTIES_CLUB_SAMEDI',
                  'SORTIES_HIVER',      
                  'SORTIES_CLUB_JEUDI', 
                  'RANDONNEES',         
                  'Nombre_sejours',     
                  'Nombre_jours_sejour']
    df = df.drop(['SORTIES_HIVER','Nombre_sejours',],axis=1)
    df['total'] = (df['RANDONNEES']+
                   df['Nombre_sorties_sejour']+ 
                   df['SORTIES_CLUB_DIMANCHE']+
                   df['SORTIES_CLUB_JEUDI']+
                   df['SORTIES_CLUB_SAMEDI'])
    
    plt.style.use('ggplot')
    fig, (ax1, ax2) = plt.subplots(2, 1, sharex=True,figsize=(15, 5))
    fig.subplots_adjust(hspace=0.05) 
    
    ax2.plot(df.index[0:6],df['RANDONNEES'][0:6], '-*b',label = 'RANDONNEES')
    ax2.plot(df.index[8:12],df['RANDONNEES'][8:],'-*b')
    ax2.plot(df.index[0:6],df['PARTICIPATION_SEJOURS'][0:6],'r-',label='PARTICIPATION_SEJOURS',linewidth=3)
    ax2.plot(df.index[8:12],df['PARTICIPATION_SEJOURS'][8:],'r-',linewidth=3)
    ax1.plot(df.index[0:6],df['Nombre_sorties_sejour'][0:6],'r-',label='Nombre_sorties_sejour',linewidth=3)
    ax1.plot(df.index[8:12],df['Nombre_sorties_sejour'][8:],'r-',linewidth=3)    
    ax2.plot(df.index[0:6],df['SORTIES_CLUB_DIMANCHE'][0:6],'*-.k',label='SORTIES_CLUB_DIMANCHE')
    ax2.plot(df.index[8:12],df['SORTIES_CLUB_DIMANCHE'][8:],'*-.k')
    ax2.plot(df.index[1:6],df['SORTIES_CLUB_JEUDI'][1:6],'-..m',label='SORTIES_CLUB_JEUDI')
    ax2.plot(df.index[8:12],df['SORTIES_CLUB_JEUDI'][8:],'-..m')
    ax2.plot(df.index[0:6],df['SORTIES_CLUB_SAMEDI'][0:6],'-+g',label='SORTIES_CLUB_SAMEDI')
    ax2.plot(df.index[8:12],df['SORTIES_CLUB_SAMEDI'][8:],'-+g')
    
    ax1.plot(df.index[0:6],df['total'][0:6],'-+y',label='total')
    ax1.plot(df.index[8:12],df['total'][8:],'-+y')
    ax2.legend(loc='upper right', bbox_to_anchor=(1.25, 1))
    ax1.legend(loc='upper right', bbox_to_anchor=(1.25, 1))
    plt.xticks(rotation=90)
    ax1.tick_params(axis='both', which='major', labelsize=16)
    ax2.tick_params(axis='both', which='major', labelsize=16)
   
    ax1.set_ylabel('# participations', fontsize=16)
    ax2.set_ylabel('# participations', fontsize=16)
    
    ax1.set_ylim(1250, 3000)
    ax2.set_ylim(50, 650)
    
    ax1.spines.bottom.set_visible(False)
    ax2.spines.top.set_visible(False)
    ax1.xaxis.tick_top()
    ax2.xaxis.tick_bottom()
    
    d = .5  # proportion of vertical to horizontal extent of the slanted line
    kwargs = dict(marker=[(-1, -d), (1, d)], markersize=12,
                  linestyle="none", color='k', mec='k', mew=1, clip_on=False)
    ax1.plot([0, 1], [0, 0], transform=ax1.transAxes, **kwargs)
    ax2.plot([0, 1], [1, 1], transform=ax2.transAxes, **kwargs)
    plt.tight_layout()
    plt.show()