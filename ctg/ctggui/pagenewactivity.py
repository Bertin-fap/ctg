import datetime
import os
import shutil
import tkinter as tk
import tkinter.messagebox
from pathlib import Path
from tkinter import filedialog


from tkcalendar import Calendar
import pandas as pd

from ctg.ctgfuncts.ctg_tools import day_of_the_date
from ctg.ctgfuncts.ctg_tools import launch_bloc_note
from ctg.ctgfuncts.ctg_tools import read_sortie_csv
from ctg.ctgfuncts.ctg_effectif import inscrit_sejour
from ctg.ctgfuncts.ctg_classes import EffectifCtg
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow

def get_new_activity(self,master, page_name, institute, ctg_path):

    def grad_date():
        d = str(cal.selection_get()).split('-')
        day = d[2]
        month =  d[1]
        year = d[0]
        prefixe = f'{month.zfill(2)}_{day.zfill(2)}'
        jour_semaine = day_of_the_date(int(day),int(month),int(year)).upper()
        variable_prefixe.set(prefixe)
        variable_jour_semaine.set(jour_semaine)
        variable_year.set(year)
        
    def browseFiles():
        filename = filedialog.askopenfilename(initialdir = Path.home() / Path('download'),
                                              title = "Select a File",
                                              filetypes = (("Fichiers .csv",
                                                            "*.csv"),
                                                           ("all files",
                                                            "*.*")))
        variable_file.set(filename)

    def get_modified_time(file):
        try:
            mtime = os.path.getmtime(file)
        except OSError:
            mtime = 0
        last_modified_date = datetime.datetime.fromtimestamp(mtime)
        return last_modified_date
        
    def set_outpath():
    
        type_sortie = variable_sortie.get()
        year = variable_year.get()
        prefixe = variable_prefixe.get()
        jour_semaine = variable_jour_semaine.get()
        
        if type_sortie == 'SEJOUR':
           output_path = ctg_path / Path(year) / Path('SEJOUR/CSV') / Path(prefixe+' sejour.csv')
        elif type_sortie == 'SORTIE DERNIERE MINUTE':
           output_path = ctg_path / Path(year) / Path('SORTIES DE DERNIERE MINUTE/CSV')
           output_path = output_path / Path(prefixe+' derniere minute.csv')
        elif type_sortie == 'VTT':
           output_path = ctg_path / Path(year) / Path('SORTIES VTT/CSV')
           output_path = output_path / Path(prefixe+' vtt.csv')
        elif type_sortie == "SORTIE HIVER":
           output_path = ctg_path / Path(year) / Path('SORTIES HIVER/CSV')
           output_path = output_path / Path(year+'_'+prefixe+' hiver.csv')
        elif type_sortie == 'SORTIE CLUB' or type_sortie == 'RANDONNEE':
            if jour_semaine== 'DIMANCHE':
                file = ' sortie du dimanche.csv' if type_sortie== 'SORTIE CLUB' else ' randonnee.csv'
                output_path = ctg_path / Path(year) / Path('SORTIES DU DIMANCHE/CSV') 
                output_path = output_path / Path(prefixe+file)
            elif jour_semaine== 'SAMEDI':
                file = ' sortie du samedi.csv' if type_sortie== 'SORTIE CLUB' else ' randonnee.csv'
                output_path = ctg_path / Path(year) / Path('SORTIES DU SAMEDI/CSV')
                output_path = output_path / Path(prefixe+file)
            else:
                file = ' sortie du jeudi.csv' if type_sortie== 'SORTIE CLUB' else ' randonnee.csv'
                output_path = ctg_path / Path(year) / Path('SORTIES DU JEUDI/CSV')
                output_path = output_path / Path(prefixe+file)
        return output_path
        
    def put_file_in_db():
        jour_semaine = variable_jour_semaine.get()
        prefixe = variable_prefixe.get()
        year = variable_year.get()
        type_sortie = variable_sortie.get()
        input_file = variable_file.get()
        name_activite = nom_sortie.get()
        
        if input_file == '':
            tkinter.messagebox.showwarning("WARNING", "Vous devez saisir le nom du fichier de la sortie")
            return
            
        if year == '':
            tkinter.messagebox.showwarning("WARNING", "Vous devez saisir la date de la sortie")
            return

 
        output_path = set_outpath()       
        shutil.copyfile(input_file, output_path)
        
        
        # Update info_randos.xlsx
        info_randos_file = ctg_path / Path(year) / Path('DATA/info_randos.xlsx')
        info_randos_df = pd.read_excel(info_randos_file)
        if type_sortie == 'SEJOUR':
            type = 'sejour'
        elif type_sortie == 'RANDONNEE':
            type = 'randonnee'
        else:
            type ='club'
            
        add_indo_df = pd.DataFrame.from_dict({'date':[year[2:]+'-'+prefixe.replace('_','-')],
                                              'jour':[jour_semaine.lower()],
                                              'name_activite':[name_activite],
                                              'name_activite_long':[name_activite],
                                              'type':[type],
                                              'nbr_jours':[int(nbr_sortie.get())],
                                              'Cout':[float(cout_sejour.get())]
                                             })
        info_randos_df = pd.concat([info_randos_df, add_indo_df], axis=0)
        info_randos_df.to_excel(info_randos_file,index=False)
        message = f'1- Le fichier :\n {input_file}\na eté copié dans: \n{output_path}\n\n'
        message = message + f'2- Mise à jour du fichier :\n {info_randos_file}'
        tkinter.messagebox.showinfo('message',message)
    
    def verif_nom():
    
        # Reads the club effectif
        year = variable_year.get()
        eff = EffectifCtg(year,ctg_path)
        df_effectif = eff.effectif_tot
        df_effectif['sejour'] = None
        df_effectif =df_effectif[['N° Licencié',
                                  'Nom',
                                  'Prénom',
                                  'Sexe',
                                  'Pratique VAE',
                                  'sejour']]
        
        # read the csv file of the event with full path file
        file_path = set_outpath()
        no_match = []
        inscrit_sejour(file_path,no_match,df_effectif)
        if no_match:
            message = "Les noms suivants n'ont pas été reconnus:\n"
            message = message + "\n".join(['-  '+tup[1]+' '+tup[2] for tup in no_match])
            message = message + "\n\n1- Un blocnote va s'ouvrir.\n2- Faites vos corrections\n3- Enregistrer le fichier et fermer blocnote"
            tkinter.messagebox.showwarning("Noms non reconnus", message)
            modified_time_init = get_modified_time(file_path)
            launch_bloc_note(file_path)
            while True:
                 modified_time = get_modified_time(file_path)
                 if modified_time != modified_time_init : break
        else:
            tkinter.messagebox.showwarning("Noms non reconnus", "Fichier correct")

    variable_year = tk.StringVar(self)
    variable_year.set('')
    variable_jour_semaine = tk.StringVar(self)
    variable_jour_semaine.set('')
    variable_prefixe = tk.StringVar(self)
    variable_prefixe.set('')
    variable_file = tk.StringVar(self)
    variable_file.set('')
    
    ### Gestion du calendrier
    cal = Calendar(self, 
                   selectmode = 'day',
                   year = datetime.datetime.today().year, 
                   month = datetime.datetime.today().month,
                   day = datetime.datetime.today().day)
     
    cal.place(x = 20, y = 100)
    date_button = tk.Button(self,
                            text = "3- Saisissez la date de la sortie",
                            command = grad_date)

    place_after(cal,date_button,dx=50,dy=250)
    
    ### Choix du type de sortie
    list_type_sortie = ["SEJOUR",
                        "VTT",
                        "SORTIE DERNIERE MINUTE",
                        "SORTIE HIVER",
                        "RANDONNEE",
                        "SORTIE CLUB",]

    default_sortie = list_type_sortie[-1]
    variable_sortie = tk.StringVar(self)
    variable_sortie.set(default_sortie)

        # Création de l'option choix du type de sortie
    OptionButton_years = tk.OptionMenu(self,
                                       variable_sortie,
                                       *list_type_sortie)

        # Création du label
    Label_years = tk.Label(self,text='1- Saisissez le type de sortie :')
    place_bellow(cal,Label_years,dx=350)
    place_after(Label_years, OptionButton_years,dy =-10)
    
    ### Choix du fichier csv
    
    label_file_explorer = tk.Label(self,
                                   text = "2- Choisisseez votre fichier",)
  
    button_explore = tk.Button(self,
                               text = "Browse Files",
                               command = browseFiles)

    place_bellow(cal,label_file_explorer,dx=350,dy=200)
    place_after(label_file_explorer, button_explore,dy =-10)

    ## Nom de sortie
    label_sortie = tk.Label(self,text='4- Saisissez le nom de la sortie :')
    place_bellow(date_button, label_sortie, dy=20)
    nom_sortie = tk.StringVar()
    textbox1 = tk.Entry(self, textvariable=nom_sortie)
    place_after(label_sortie, textbox1,dy =0) 

    ## Nombre de jours (séjour)
    label_nbr_sortie = tk.Label(self,text='5- Saisissez le nombre de sorties du séjour :')
    place_bellow(label_sortie, label_nbr_sortie, dy=20)
    nbr_sortie = tk.StringVar()
    nbr_sortie.set(1)
    textbox2 = tk.Entry(self, textvariable=nbr_sortie)
    place_after(label_nbr_sortie, textbox2,dy =0)

    ## Cout du jour (séjour)
    label_cout_sejour = tk.Label(self,text='6- Saisissez le cout du séjour :')
    place_bellow(label_nbr_sortie, label_cout_sejour, dy=20)
    cout_sejour = tk.StringVar()
    cout_sejour.set(0)
    textbox3 = tk.Entry(self, textvariable=cout_sejour)
    place_after(label_cout_sejour, textbox3,dy =0)    
    
    ### Dépacer renommer le fichier
  
    button_db = tk.Button(self,
                          text = "7- Mettre le fichier dans la BD",
                          command = put_file_in_db)

    place_bellow(label_cout_sejour, button_db,dy=10)
    
    ### Vérification des noms
  
    button_verif = tk.Button(self,
                             text = "8- Vérification des noms (optionel)",
                             command = verif_nom)

    place_bellow(button_db, button_verif, dy=10)    
    
     
    
    
     
