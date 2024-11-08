import datetime
import os
import shutil
import tkinter as tk
import tkinter.messagebox
from pathlib import Path
from tkinter import filedialog


from tkcalendar import Calendar
import pandas as pd

from ctg.ctgfuncts.make_list_adherents import make_list_emargement
from ctg.ctgfuncts.make_list_adherents import make_list_adherents

from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow

def create_listes(self,master, page_name, institute, ctg_path):

    def grad_date():
        d = str(cal.selection_get()).split('-')
        day = d[2]
        month =  d[1]
        year = d[0]
       
        variable_month.set(month)
        variable_jour.set(day)
        variable_year.set(year)

    def _create_liste_emargement():
        make_list_emargement(ctg_path,variable_jour.get(),variable_month.get(),variable_year.get())
    
    def _create_liste_rando():
        make_list_adherents(ctg_path)
        
        
    
           
    variable_year = tk.StringVar(self)
    variable_year.set('')
    variable_jour = tk.StringVar(self)
    variable_jour.set('')
    variable_month = tk.StringVar(self)
    variable_month.set('')
   
    
    ### Gestion du calendrier
    cal = Calendar(self, 
                   selectmode = 'day',
                   year = datetime.datetime.today().year, 
                   month = datetime.datetime.today().month,
                   day = datetime.datetime.today().day)
     
    cal.place(x = 0, y = 100)
    date_button = tk.Button(self,
                            text = "Saisissez la date de l'assemblée générale",
                            command = grad_date)

    place_after(cal,date_button,dx=0,dy=250)
    
    liste_emargement_button = tk.Button(self,
                            text = "Création de la liste d'émargement",
                            command = _create_liste_emargement)

    place_bellow(date_button,liste_emargement_button,dx=30,dy=10)
    
    liste_rando_button = tk.Button(self,
                                text = "Création de la liste randonnée",
                                command = _create_liste_rando)

    place_bellow(liste_emargement_button,liste_rando_button,dx=0,dy=10)