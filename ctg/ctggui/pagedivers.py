__all__ = ['create_divers_analysis']

# Standard library imports
import os
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import font as tkFont
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path

# 3rd party imports
import pandas as pd

# Local imports
import ctg.ctggui.guiglobals as gg
from ctg.ctggui.guitools import encadre_RL
from ctg.ctggui.guitools import font_size
from ctg.ctggui.guitools import mm_to_px
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow
from ctg.ctggui.guitools import last_available_years
from ctg.ctgfuncts.ctg_plot import plot_ctg
from ctg.ctgfuncts.ctg_synthese import nbr_sejours_adherent
from ctg.ctgfuncts.ctg_cd import plot_cd_evolution

def create_divers_analysis(self, master, page_name, institute, ctg_path):

    """
    Description : function working as a bridge between the BiblioMeter
    App and the functionalities needed for the use of the app

    Uses the following globals :
    - DIC_OUT_PARSING
    - FOLDER_NAMES

    Args: takes only self and ctg_path as arguments.
    self is the instense in which PageThree will be created
    ctg_path is a type Path, and is the path to where the folders
    organised in a very specific way are stored

    Returns : nothing, it create the page in self
    """

    # Internal functions
    def _launch_geo_analysis():

        # Getting year selection
        year_select =  variable_years.get()
        kol = plot_ctg(ctg_path,year_select)

        path_html = ctg_path / Path(str(year_select)) / Path('STATISTIQUES') / Path('HTML') 
        path_html = path_html / Path("ctg.html")
        kol.save(path_html)

        info_title = "- Information -"
        info_text  = ("L'analyse de la localisation géographique des membres a été effectuée"
                      f"pour l'année {year_select} "
                      f"\n\nLe fichier .html obtenu : 'ctg.html' "
                      "a été créé dans le dossier :"
                      f"\n\n{ctg_path}\{str(year_select)}\STATISTIQUES\HTML.")

        messagebox.showinfo(info_title, info_text)

    def _launch_participation_sejour(ctg_path):

        # Getting year selection
        year_select =  variable_years.get()

        nbr_sejours_adherent(year_select,ctg_path)
        
    def _launch_cd_analysis(ctg_path):

        # Getting year selection

        _,_ = plot_cd_evolution(ctg_path)

    from ctg.ctggui.pageclasses import AppMain

    # Setting effective font sizes and positions (numbers are reference values)
    eff_etape_font_size      = font_size(gg.REF_ETAPE_FONT_SIZE,   AppMain.width_sf_min)
    eff_launch_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE-1, AppMain.width_sf_min)
    eff_help_font_size       = font_size(gg.REF_ETAPE_FONT_SIZE-2, AppMain.width_sf_min)
    eff_select_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE, AppMain.width_sf_min)
    eff_buttons_font_size    = font_size(gg.REF_ETAPE_FONT_SIZE-3, AppMain.width_sf_min)

    synthese_analysis_x_pos_px     = mm_to_px(10 * AppMain.width_sf_mm,  gg.PPI)
    synthese_analysis_y_pos_px     = mm_to_px(40 * AppMain.height_sf_mm, gg.PPI)
    year_analysis_label_dx_px  = mm_to_px( 0 * AppMain.width_sf_mm,  gg.PPI)
    year_analysis_label_dy_px  = mm_to_px(15 * AppMain.height_sf_mm, gg.PPI)
    launch_dx_px             = mm_to_px( 0 * AppMain.width_sf_mm,  gg.PPI)
    launch_dy_px             = mm_to_px( 5 * AppMain.height_sf_mm, gg.PPI)

    year_button_x_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_X_MM * AppMain.width_sf_mm,  gg.PPI)
    year_button_y_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)
    dy_year                  = -6
    ds_year                  = 5

    # Setting common attributs
    etape_label_format = 'left'
    etape_underline    = -1

    ### Choix de l'année
    list_years = last_available_years(ctg_path,year_number=2020)[1:]
    default_year = list_years[-1]
    variable_years = tk.StringVar(self)
    variable_years.set(default_year)

        # Création de l'option button des années
    self.font_OptionButton_years = tkFont.Font(family = gg.FONT_NAME,
                                               size = eff_buttons_font_size)
    self.OptionButton_years = tk.OptionMenu(self,
                                            variable_years,
                                            *list_years)
    self.OptionButton_years.config(font = self.font_OptionButton_years)

        # Création du label
    self.font_Label_years = tkFont.Font(family = gg.FONT_NAME,
                                        size = eff_select_font_size,
                                        weight = 'bold')
    self.Label_years = tk.Label(self,
                                text = gg.TEXT_GEO_PI,
                                font = self.font_Label_years)
    self.Label_years.place(x = year_button_x_pos, y = year_button_y_pos)

    place_after(self.Label_years, self.OptionButton_years, dy = dy_year)

    ################## Synthèse des activité

    ### Titre
    synthese_analysis_font = tkFont.Font(family = gg.FONT_NAME,
                                         size = eff_etape_font_size,
                                         weight = 'bold')
    synthese_analysis_label = tk.Label(self,
                                       text = gg.TEXT_GEO,
                                       justify = etape_label_format,
                                       font = synthese_analysis_font,
                                       underline = etape_underline)

    synthese_analysis_label.place(x = synthese_analysis_x_pos_px,
                                  y = synthese_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label = tk.Label(self,
                          text = gg.HELP_GEO,
                          justify = "left",
                          font = help_label_font)
    place_bellow(synthese_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des IFs
    synthese_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                                size = eff_launch_font_size)
    synthese_analysis_launch_button = tk.Button(self,
                                                text = gg.BUTT_GEO,
                                                font = synthese_analysis_launch_font,
                                                command = lambda: _launch_geo_analysis())
    place_bellow(help_label,
                 synthese_analysis_launch_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)

    ################## Analyse des mots clefs

    ### Titre
    member_analysis_label_font = tkFont.Font(family = gg.FONT_NAME,
                                             size = eff_etape_font_size,
                                             weight = 'bold')
    member_analysis_label = tk.Label(self,
                                     text = gg.TEXT_NBR_SEJOURS,
                                     justify = "left",
                                     font = member_analysis_label_font)
    place_bellow(synthese_analysis_launch_button,
                 member_analysis_label,
                 dx = year_analysis_label_dx_px,
                 dy = year_analysis_label_dy_px)

    ### Explication de l'étape
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label = tk.Label(self,
                          text = gg.HELP_NBR_SEJOURS,
                          justify = "left",
                          font = help_label_font)
    place_bellow(member_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des mots clefs
    member_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                                size = eff_launch_font_size)
    member_analysis_button = tk.Button(self,
                                         text = gg.BUTT_NBR_SEJOURS,
                                         font = member_analysis_launch_font,
                                         command = lambda: _launch_participation_sejour(ctg_path))
    place_bellow(help_label,
                 member_analysis_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)

################## Analyse du CD

    ### Titre
    cd_analysis_label_font = tkFont.Font(family = gg.FONT_NAME,
                                          size = eff_etape_font_size,
                                          weight = 'bold')
    #gg.TEXT_AGE_ANALYSIS
    cd_analysis_label = tk.Label(self,
                                  text = "Anlyse du Comité Directeur",
                                  justify = "left",
                                  font = cd_analysis_label_font)
    place_bellow(member_analysis_button,
                 cd_analysis_label,
                 dx = year_analysis_label_dx_px,
                 dy = year_analysis_label_dy_px)
    
    ### Explication de l'étape
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                               size = eff_help_font_size)
    #gg.HELP_AGE_ANALYSIS
    help_label_cd = tk.Label(self,
                              text = "Analyse de l'évolution des effectifs du Comité Directeur de 2014 à aujourd'hui",
                              justify = "left",
                              font = help_label_font)
    place_bellow(cd_analysis_label,
                 help_label_cd)
    
    ### Bouton pour lancer l'analyse des mots clefs
    cd_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                         size = eff_launch_font_size)
    #gg.BUTT_AGE_ANALYSIS
    cd_analysis_button = tk.Button(self,
                                    text = "Lancer l'analyse",
                                    font = cd_analysis_launch_font,
                                    command = lambda: _launch_cd_analysis(ctg_path))
    place_bellow(help_label_cd,
                 cd_analysis_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)