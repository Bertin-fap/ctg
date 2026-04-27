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
from ctg.ctgfuncts.ctg_sg2xlsx import sg2xlsx
from ctg.ctgfuncts.ctg_ffvelo_adhesion import finance_ffct
from ctg.ctgfuncts.ctg_operation import create_justificatif

from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow

def create_compta(self,master, page_name, institute, ctg_path):

    def grad_date():
        d = str(cal.selection_get()).split('-')
        day = d[2]
        month =  d[1]
        year = d[0]
       
        variable_month.set(month)
        variable_jour.set(day)
        variable_year.set(year)

    def _input_new_operation():
        
        if somme_value.get() == '':
            tkinter.messagebox.showwarning("WARNING", "Vous devez saisir la somme")
            return
            
        if nom_operation.get() == '':
            tkinter.messagebox.showwarning("WARNING", "Vous devez le libellé de l'opération")
            return
            
        if variable_year.get() == '':
            tkinter.messagebox.showwarning("WARNING", "Vous devez la date de l'opération")
            return
            
        ymd= f'{variable_year.get()}{str(variable_month.get()).zfill(2)}{str(variable_jour.get()).zfill(2)}'
        somme = somme_value.get()  
        somme = str("%.2f" % (float( somme.replace(',','.')))).replace('.',',')
        create_justificatif(ymd,variable_year.get(),nom_operation.get(),somme)
        
    
    def _create_sg2xlsx():
        pass
        
    def _ffct_finance():
        finance_ffct()    
         
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
                            text = "Saisissez la date de l'opération",
                            command = grad_date)
    place_after(cal,date_button,dx=0,dy=250)
                            
    ## Nature de l'opération
    operation = tk.Label(self,text="Nature de l'opération :")
    place_bellow(cal,operation,dx=350)
    nom_operation = tk.StringVar()
    textbox1 = tk.Entry(self, textvariable=nom_operation)
    place_after(operation, textbox1,dy =0) 

    ## Somme
    somme = tk.Label(self,text='Somme (€) :')
    place_bellow(cal,somme,dx=350,dy=200)
    somme_value = tk.StringVar()
    textbox2 = tk.Entry(self, textvariable=somme_value)
    place_after(somme, textbox2,dy =0)

    
    ## Saisir une opération
    liste_emargement_button = tk.Button(self,
                            text = "Entrer une opération",
                            command = _input_new_operation)

    place_bellow(date_button,liste_emargement_button,dx=30,dy=10)
    
    liste_sg2excel_button = tk.Button(self,
                                  text = "Création d'un fichier xlsx à partir d'un fichier csv de SG",
                                  command = _create_sg2xlsx)

    place_bellow(liste_emargement_button,liste_sg2excel_button,dx=0,dy=10)
    
    
    ffct_finance_button = tk.Button(self,
                                text = "FFCT Finance",
                                command = _ffct_finance)

    place_bellow(liste_sg2excel_button ,ffct_finance_button,dx=0,dy=10)