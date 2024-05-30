__all__ =['evolution_sorties',
          'nbr_sejours_adherent',
          'plot_pie_synthese',
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

# Third party imports
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import yaml

# Internal imports
from ctg.ctggui.guiglobals import ACTIVITE_LIST
from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctgfuncts.ctg_tools import get_cout_total
from ctg.ctgfuncts.ctg_effectif import read_effectif_corrected
from ctg.ctgfuncts.ctg_tools import get_sejour_info

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
            list_df.extend([pd.read_excel(path_dir / Path(file), engine='openpyxl') for file in files])

    df_total = pd.concat(list_df, ignore_index=True)

    df_effectif = read_effectif_corrected(ctg_path)

    df_total['Pratique VAE'].fillna('Non',inplace=True)

    # nombre moyen de participant par activité
    for x in df_total.groupby(['Type']):
        print(x[0],len(x[1]),len(set(x[1]['sejour'])),len(x[1])/len(set(x[1]['sejour'])))

    file = ctg_path / Path(year) / Path('STATISTIQUES') / Path('synthese.xlsx')
    df_total.to_excel(file)

def plot_pie_synthese(year:str,ctg_path:pathlib.WindowsPath)->None:

    def func(pct, allvalues):
        absolute = round(pct / 100.*np.sum(allvalues),0)
        #return "{:.1f}%\n({:d})".format(pct, absolute)
        return  f"{int(round(absolute,1))}\n{round(pct,1)} %"

    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('synthese.xlsx')
    df_total = pd.read_excel(file_in)
    df_total = df_total.dropna(subset=['Type'])
    df_total = df_total.dropna(subset=['Nom'])

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
    fig = plt.figure(figsize =(10, 7))
    _, _, autotexts = plt.pie(data,
                              labels = sorties,
                              autopct = lambda pct: func(pct, data),
                              explode = explode,
                              textprops={'fontsize': 18})
    plt.title(year, pad=50, fontsize=20)

    _ = plt.setp(autotexts, **{'color':'k', 'weight':'bold', 'fontsize':14})

    plt.tight_layout()
    plt.show()

    fig_file = 'SORTIES_PIE.png'
    plt.savefig(ctg_path / Path(year) / Path('STATISTIQUES') / Path(fig_file),bbox_inches='tight')



def synthese_adherent(year:str,ctg_path:pathlib.WindowsPath):

    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('synthese.xlsx')
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
               'TOTAL']] = [0,0,0,0,0,0,0,0]

    dg = pd.concat([dg, df_orphan], axis=0)

    file_out = ctg_path / Path(year) / Path('STATISTIQUES') / Path('synthese_adherent.xlsx')
    dg.to_excel(file_out)

def synthese_randonnee(year:str,ctg_path:pathlib.WindowsPath,type_sejour:str):

    def addlabels(x,y,offset):
        for i in range(len(x)):
            if y[i] != 0:
                plt.text(x[i],y[i]+offset,round(y[i],1),size=10)


    file_in = Path(ctg_path) / Path(year) / Path('STATISTIQUES') / Path('synthese.xlsx')
    df_total = pd.read_excel(file_in)
    df_total = df_total.dropna(subset=['Nom'])
    dg = df_total.query('Type==@type_sejour').groupby('sejour').agg('count')['N° Licencié']

    file_info = Path(ctg_path) / Path(year) / Path('DATA') / Path('info_randos.xlsx')
    cout_total_rando = get_cout_total(year,type_sejour.lower(),dg,ctg_path)

    fig = plt.figure()
    plt.bar(range(len(dg)),dg.tolist())
    addlabels(list(range(len(dg))),dg.tolist(),0.2)
    plt.xticks(range(len(dg)), dg.index, rotation='vertical')
    plt.tick_params(axis='x', rotation=90, labelsize=10)
    plt.tick_params(axis='y',labelsize=10)

    if type_sejour == 'RANDONNEE':
        _ = plt.title(f'# randos : {len(dg)} , # participants : {sum(dg)}, Coût : {cout_total_rando} €')
    else:
        _ = plt.title(f'# sejours : {len(dg)} , # participants : {sum(dg)}, Coût : {cout_total_rando} €')
    plt.tight_layout()
    plt.show()

def nbr_sejours_adherent(year:str, ctg_path:pathlib.WindowsPath):

    plt.style.use('ggplot')

    # function to add value labels
    def addlabels(x,y):
        for i in range(len(x)):
            plt.text(x[i]-0.2,y[i]+1,y[i],size=label_size)
    label_size = 15
    file_in = ctg_path / Path(year) / Path('STATISTIQUES') / Path('synthese_adherent.xlsx')
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
    plt.savefig(ctg_path / Path(year) / Path('STATISTIQUES') / Path(fig_file),bbox_inches='tight')

def _read_memory_sorties():

    # Reads the default PVcharacterization.yaml config file
    path_config_file = Path(__file__).parent.parent / Path('ctgfuncts') / Path('CTG_RefFiles') / Path('memory_sorties.yml')
    with open(path_config_file) as file:
        memory = yaml.safe_load(file)


    return memory

def evolution_sorties(type:str,ctg_path:pathlib.WindowsPath):

    def add_memory(stat_dic,years):
        memory = _read_memory_sorties()


        for year,v in memory['memory'].items():

            stat_dic[str(year)] = statyear(v['PARTICIPATION_SEJOURS'],                # nbr_sejours
                                           0,                          # jours_sejour
                                           v['SORTIES_CLUB_DIMANCHE'], # sortie_dimanche_club
                                           v['SORTIES_CLUB_SAMEDI'],   # sortie_samedi_club
                                           v['SORTIES_HIVER'],         # sortie_hiver_club
                                           v['SORTIES_CLUB_JEUDI'],    # sortie_jeudi_club
                                           v['RANDONNEES'],            # randonnee
                                           v['Nombre_sejours'],        # nbr_sejours
                                           0,)                         # nbr_jours_sejours

            years.append(str(year))

    def fill_stat_year(year:str):

        sejour_info = get_sejour_info(ctg_path,year)


        file_in = ctg_path / Path(str(year)) / Path('STATISTIQUES') / Path('synthese_adherent.xlsx')

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