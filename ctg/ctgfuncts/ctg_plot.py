__all_ = ["built_lat_long",
          "plot_ctg",
          "stat_sorties_club",]

# Standard library import
import datetime
import os
from collections import Counter
from math import asin, cos, radians, sin, sqrt
from pathlib import Path
from tkinter import messagebox

# 3rd party imports
import folium
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import pandas

# Internal imports
from ctg.ctgfuncts.ctg_effectif import read_effectif
from ctg.ctgfuncts.ctg_effectif import read_effectif_corrected
from ctg.ctgfuncts.ctg_effectif import count_participation
from ctg.ctgfuncts.ctg_tools import parse_date
from ctg.ctgfuncts.ctg_tools import built_lat_long

def plot_ctg(df:pandas.core.frame.DataFrame):

    '''generates an html file of the membrers geographical location using the
    column ville of the DataFrame df
    '''

    trace_radius = True
    dh = built_lat_long(df)
    
    villes_set = set(dh['Ville'])
    dh = dh.dropna()
    villes1_set = set(dh['Ville'])
    if len(villes_set-villes1_set) !=0:
        messagebox.showwarning('Villes non reconnues',';'.join(list(villes_set-villes1_set)))

    group_adjacent = lambda a, k: list(zip(*([iter(a)] * k)))

    dict_cyclo = {}
    for ville,y in df.groupby(['Ville'])['Nom']:
        chunk = []
        for i in range(0,len(y),3):
            chunk.append(','.join(y[i:i+3] ))

        dict_cyclo[ville] = '\n'.join(chunk)
    dict_cyclo = {k[0]:v for k,v in dict_cyclo.items()}
    kol = folium.Map(location=[45.2,5.7], tiles='openstreetmap', zoom_start=12)

    long_genoble, lat_grenoble = dh.query("Ville=='GRENOBLE'")[['long','lat']].values.flatten()
    if trace_radius:
        folium.Circle(
                      location=[lat_grenoble, long_genoble],
                      radius=8466,
                      popup='50 km ',
                      color="black",
                      fill=False,
                      ).add_to(kol)
    for latitude,longitude,size, ville in zip(dh['lat'],dh['long'],dh['number'],dh['Ville']):

        long_ville, lat_ville =dh.query("Ville==@ville")[['long','lat']].values.flatten()
        dist_grenoble_ville = _distance(lat_grenoble, long_genoble,lat_ville, long_ville )
        color='red' if dist_grenoble_ville>19.35 else 'blue'
        if ville == "grenoble":
            folium.Circle(
                location=[latitude, longitude],
                radius=size*50,
                popup=f'{ville} ({size}): {dict_cyclo[ville]} ',
                color="yellow",
                fill=True,
            ).add_to(kol)
        else:
                folium.Circle(
                location=[latitude, longitude],
                radius=size*100,
                popup=f'{ville} ({size}): {dict_cyclo[ville]}',
                color=color,
                fill=True,
            ).add_to(kol)
    return kol

def stat_sorties_club(path_sorties_club, ctg_path, ylim=None, file_label=None,year = None):

    ''' to do
    '''

    def addlabels(x,y):

        for i in range(len(x)):
            d = x[i]
            color = 'g'
            if os.path.split(path_sorties_club)[-1] == "SEJOUR" :
                v = info_rando.query('date==@d and type=="sejour"')['name_activite'].tolist()
            else:
                v = info_rando.query('date==@d and type!="sejour"')['name_activite'].tolist()
                t = info_rando.query('date==@d and type!="sejour"')['type'].tolist()
                if len(t) != 0:
                    color = "k" if t[0] =='randonnee' else "g"

            name = v[0] if len(v) != 0 else ""

            plt.text(i-0.2,y[i]+1,
                     name,
                     size=10,
                     rotation=90,
                     color=color #info_rando['color']
                     )

    if file_label is not None and os.path.isfile(file_label):
        flag_labels = True
        info_rando = pd.read_excel(file_label)
    else:
        flag_labels = False
        print(file_label)

    no_match,df_total,_ = count_participation(path_sorties_club,ctg_path,year,info_rando)
    if no_match is None:
        messagebox.showinfo('WARNING',"Aucun participant n'a participé à ce type de sortie" )
    else:
        text_message = ''
        for tup in no_match:
              text_message += f'Le nom {tup[1]}, {tup[2]} est inconnu dans le fichier : "{os.path.split(tup[0])[-1]}"'
              text_message += '\n'

        if len(text_message) : messagebox.showinfo('WARNING',text_message )

    if year is None:
        currentDateTime = datetime.datetime.now()
        date = currentDateTime.date()
        year = date.strftime("%Y")

    df_effectif = read_effectif_corrected(ctg_path,
                                          year)

    dic_sexe = dict(zip(df_effectif['N° Licencié'], df_effectif['Sexe']))
    dic_sexe[None] = 'irrelevant'
    dic_vae =dict(zip(df_effectif['N° Licencié'],df_effectif['Pratique VAE']))

    df_total['sexe'] = df_total['N° Licencié'].map(dic_sexe)

    df_total = df_total[df_total['sejour']!='aucun' ]
    df_total['sejour'] = df_total['sejour'].apply(lambda s:parse_date(s,str(year)).strftime('%y-%m-%d'))
    df_total['VAE'] = df_total['N° Licencié'].map(dic_vae)
    df_total['VAE'].fillna('Non',inplace=True)
    dic_sexe = dict(M="Homme",F="Femme")
    dic_vae = dict(Oui="VAE",Non="Musculaire")
    df_total = df_total.replace({"sexe": dic_sexe})
    df_total = df_total.replace({"VAE": dic_vae})
    if df_total['Nom'].isna().all():
        return None
    dg = df_total.groupby(['sexe','VAE'])['sejour'].value_counts().unstack().T

    try:
        dg['irrelevant'] = dg['irrelevant'] - 1
    except KeyError as error:
        pass

    fig, ax = plt.subplots(figsize=(15, 5))


    dg[['Femme','Homme']].plot(kind='bar',
                       ax=ax,
                       width=0.5,
                       stacked=True,
                       color = {('Femme', 'Musculaire'): '#1f77b4',
                                ('Femme', 'VAE'): '#ff7f0e',
                                ('Homme', 'Musculaire'): '#2ca02c',
                                ('Homme', 'VAE'): '#d62728',} )

    if flag_labels : addlabels(dg.index,dg.sum(axis=1).astype(int).tolist())

    plt.xlabel('')
    plt.tick_params(axis='x', rotation=90,labelsize=15)
    plt.ylabel('Nombre de licenciers',size=15)
    plt.xlabel('')
    plt.tick_params(axis='x', rotation=90,labelsize=15)
    plt.tick_params(axis='y',labelsize=15)
    type_sortie = os.path.split(path_sorties_club)[-1] + ' ' + str(year)
    plt.title(type_sortie,fontsize=15,pad=50)

    if ylim is not None:
        plt.ylim(ylim)
    else:
        ylim = (0,1.5*max(Counter(df_total['sejour']).values()))
        plt.ylim(ylim)


    plt.legend(bbox_to_anchor =(0.75, 1.15), ncol = 2)
    plt.tight_layout()
    plt.show()
    fig_file = os.path.split(path_sorties_club)[-1].replace(' ','_')+'.png'
    plt.savefig(ctg_path / Path(str(year)) / Path('STATISTIQUES') / Path(fig_file),bbox_inches='tight')

    return df_total

def _distance(ϕ1:float, λ1:float,ϕ2:float, λ2:float)->float:

    '''Computes the distance in kilometers between to points referenced 
    by there longitudes (in decimal degrees) and there latitudes (in decimal degrees)
    '''

    ϕ1, λ1 = radians(ϕ1), radians(λ1)
    ϕ2, λ2 = radians(ϕ2), radians(λ2)
    rad = 6371 # Earth radius [km]
    dist = 2 * rad * asin(
                          sqrt(
                               sin((ϕ2 - ϕ1) / 2) ** 2
                             + cos(ϕ1) * cos(ϕ2) * sin((λ2 - λ1) / 2) ** 2
                            ))
    return dist