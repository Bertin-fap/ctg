import itertools
from calendar import monthrange
import datetime
import os
from pathlib import Path

import pandas as pd

def jours_sortie_club(year,sorties_df,day_list):

    days_dict = {0: 'dimanche',
                 1: 'lundi',
                 2: 'mardi',
                 3: 'mercredi',
                 4: 'jeudi',
                 5: 'vendredi',
                 6: 'samedi'}
    
    inv_days_dict = dict(zip(days_dict.values(),days_dict.keys()))
    month_dict = dict(zip(itertools.islice(itertools.cycle(l:=range(1,13)),2,2+len(l)),l))
    dic = {}

    
            
    date_list = []
    num_week_list = []
    name_sortie = []
    jour_list = []
    for m in range(1,11): #range(1,13):
        m_save=m
        y = year%100
        c = int(year/100)
        m = month_dict[m]
        if m>10 : y = y-1
        _,nb_days  = monthrange(y+200*c, m_save)
        for d in range (1,nb_days+1):
            num_day = (d + int(2.6*m - 0.2) - 2*c + y + int(c/4) + int(y/4))%7
            day_flag = [True if inv_days_dict[x]==num_day else False for x in day_list]
            if any(day_flag):
                try:
                    num_week = datetime.date(year,m_save, d).isocalendar().week
                except:
                    num_week = None
                
                if num_week is not None and num_week>7 and num_week<45:
                    num_week_list.append(num_week)
                    date = f'{str(year)[2:]}-{str(m_save).zfill(2)}-{str(d).zfill(2)}'
                    date_list.append(date)
                    dg = sorties_df.query('date==@date')
                    if len(dg)>0:
                        name_sortie.append(dg['name_activite_long'].tolist()[0])
                    else:
                        name_sortie.append(None)
                    jour_list.append(days_dict[num_day])

                    
    dic['semaine'] = num_week_list 
    dic[str(year)] = date_list 
    dic['jour'] = jour_list
    dic[f'Nom-{str(year)}'] = name_sortie
    df = pd.DataFrame(dic)        
    return df

def build_synthese_rando(year_list,day_list):
    year_df_list = []
    for year in year_list:
        file = Path(r'C:\Users\franc\CTG\SORTIES') / str(year) / 'DATA' / 'info_randos.xlsx'
        year_df_list.append(pd.read_excel(file))
    sorties_df = pd.concat(year_df_list, axis=0)
    
    
    for year in year_list:
        df_list.append(jours_sortie_club(year,sorties_df,day_list))
    
    df_concat = pd.concat(df_list, axis=1)
    df_concat.to_excel(r'c:\users\franc\Temp\sorties_we.xlsx',index=False)

year_list = [2022,2023,2024,2025]
day_list = ['samedi','dimanche']
build_synthese_rando(year_list,day_list)