__all__ = ["make_calendar",]

import calendar
import datetime
import os
from pathlib import Path
from tkinter import messagebox

import pandas as pd

def calcul_date_paques(annee):
    """Calcule la date de Pâques pour une année donnée"""
    a = annee % 19
    b = annee // 100
    c = annee % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois = (h + l - 7 * m + 114) // 31
    jour = 1 + (h + l - 7 * m + 114) % 31
    return datetime.date(annee, mois, jour)

def calcul_date_ascension(annee):
    """Calcule la date de l'Ascension pour une année donnée"""
    date_paques = calcul_date_paques(annee)
    date_ascension = date_paques + datetime.timedelta(days=39)
    return (date_ascension.day,date_ascension.month, date_ascension.year)

def calcul_date_lundi_pacques(annee):
    """Calcule la date de du lundi de Pâques pour une année donnée"""
    date_paques = calcul_date_paques(annee)
    date_lundi_pacques = date_paques + datetime.timedelta(days=1)
    return (date_lundi_pacques.day,date_lundi_pacques.month, date_lundi_pacques.year)

def calcul_date_lundi_pentecote(annee):
    """Calcule la date de du lundi de Pâques pour une année donnée"""
    date_paques = calcul_date_paques(annee)
    date_lundi_pentecote = date_paques + datetime.timedelta(days=50)
    return ( date_lundi_pentecote.day, date_lundi_pentecote.month,  date_lundi_pentecote.year)

def day_of_the_date(day,month,year):
    # Compute the day of the week of the date after "Elementary Number Theory David M. Burton Chap 6.4" [1]
    
    
    days_dict = {0: 'Dimanche',
                 1: 'Lundi',
                 2: 'Mardi',
                 3: 'Mercredi',
                 4: 'Jeudi',
                 5: 'Vendredi',
                 6: 'Samedi'}

    month_dict = {3: 1, 4: 2, 5: 3, 6: 4, 7: 5,
                  8: 6, 9: 7, 10: 8, 11: 9, 12:
                  10, 1: 11, 2: 12} # [1] p. 125
    
    y = year%100
    c = int(year/100)
    m = month_dict[month]
    if m>10 : y = y-1

    return days_dict[(day + int(2.6*m - 0.2) - 2*c + y + int(c/4) + int(y/4))%7] # [1] thm 6.12

def month_idx(month):
    months_inv = {'Janvier':1, 
                  'Févier':2,
                  'Mars':3,
                  'Avril':4, 
                  'Mai':5,
                  'Juin':6,
                  'Juillet':7,
                  'Août':8,
                  'Septembre':9,
                  'Octobre':10,
                  'Novembre':11,
                  'Decembre':12}
    index = months_inv[month]

    return index 


def make_calendar(year,ctg_path):

    """
    """
   
    month_list = ['Févier', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre','Novembre']
    dic_jours_feries = {"Lundi de Pâques" :  calcul_date_lundi_pacques(year),
                      "Fête du Travail" : (1,5, year),
                      "Victoire 1945" : (8, 5, year),
                      "Ascension" : calcul_date_ascension(year),
                      "Lundi de Pentecôte" : calcul_date_lundi_pentecote(year),
                      "Fête nationale" : (14, 7, year),
                      "Assomption" : (15, 8, year),
                      "Toussaint" : (1,11, year),
                      "Armistice 1918" : (11, 11, year),
                      }
    
    dic = {}
    for index, month in enumerate(month_list):
        label_list = []
        for d in range(1,calendar.monthrange(year, month_idx(month))[1]+1):
            day = day_of_the_date(d,month_idx(month),year)
            if day=="Samedi" or day=="Dimanche" or  (d,month_idx(month),year) in dic_jours_feries.values():
                label = d
            else:
                label = day[0]
            label_list.append(label)
            
        for _ in range(31-len(label_list)):
            label_list.append('  ')
            
        dic[str(index)] = label_list
        dic[month] = ["          "]*31
       
    df = pd.DataFrame.from_dict(dic)
    
    path_file = Path(ctg_path).parent.absolute() / "CALENDRIER" / str(year)
    os.makedirs(path_file,exist_ok=True)
    file = path_file / f"calendar_{str(year)}.xlsx"
    
    df.to_excel(file,index=False)
    messagebox.showinfo(title="Calendar", message=f"Le fichier {file} a été crée")