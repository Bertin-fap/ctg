__all__ = ['create_sorties_analysis']

# Standard library imports
import os
import shutil
import tkinter as tk
from tkinter import font as tkFont
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path

# 3rd party imports
import pandas as pd

# Local imports
import ctg.ctggui.guiglobals as gg
from ctg.ctgfuncts.ctg_plot import stat_sorties_club
from ctg.ctgfuncts.ctg_synthese import synthese_randonnee
from ctg.ctggui.guitools import encadre_RL
from ctg.ctggui.guitools import font_size
from ctg.ctggui.guitools import mm_to_px
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow
from ctg.ctggui.guitools import last_available_years
from ctg.ctggui.guitools import get_available_sorties

def _launch_sorties_analysis(ctg_path,
                             year,
                             type_sortie):
    """
    """



    path_sorties =  ctg_path / Path(str(year)) / Path(type_sortie)
    file_label =  ctg_path / Path(str(year)) / Path(r'DATA/info_randos.xlsx')
    df_total = stat_sorties_club(path_sorties,ctg_path,None,file_label,year)


    #info_title = "- Information -"
    #info_text  = f"L'analyse des IFs a été effectuée pour l'année {year_select} "
    #info_text += f"avec les valeurs {analysis_if}. "
    #info_text += f"\n\nLes fichiers obtenus ont été créés dans le dossier :"
    #info_text += f"\n\n''."
    #info_text += f"\n\nLa base de donnée des indicateurs respective de l'Institut "
    #info_text += f"et de chaque département a été mise à jour "
    #info_text += f"avec les résultats de cette analyse et se trouve dans le dossier :"
    #info_text += f"\n\n''."
    #messagebox.showinfo(info_title, info_text)

def _launch_rando_analysis(ctg_path,
                           year,
                           type_sortie):
    """
    """

    if type_sortie == "SEJOUR":
        synthese_randonnee(year,ctg_path,'SEJOUR')
    else:
        synthese_randonnee(year,ctg_path,'RANDONNEE')

def create_sorties_analysis(self, master, page_name, institute, ctg_path):

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
    def _launch_sorties_analysis_try():
        # Getting year selection
        year_select = variable_years.get()
        type_sortie = variable_sortie.get()
        _launch_sorties_analysis(ctg_path,
                                 year_select,
                                 type_sortie)
        return

    def _launch_rando_analysis_try():
        # Getting year selection
        year_select = variable_years.get()
        type_sortie = variable_sortie.get()
        _launch_rando_analysis(ctg_path,
                               year_select,
                               type_sortie)
        return

    from ctg.ctggui.pageclasses import AppMain

    # Setting effective font sizes and positions (numbers are reference values)
    eff_etape_font_size      = font_size(gg.REF_ETAPE_FONT_SIZE,   AppMain.width_sf_min)
    eff_launch_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE-1, AppMain.width_sf_min)
    eff_help_font_size       = font_size(gg.REF_ETAPE_FONT_SIZE-2, AppMain.width_sf_min)
    eff_select_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE, AppMain.width_sf_min)
    eff_buttons_font_size    = font_size(gg.REF_ETAPE_FONT_SIZE-3, AppMain.width_sf_min)

    sorties_analysis_x_pos_px     = mm_to_px(10 * AppMain.width_sf_mm,  gg.PPI)
    sorties_analysis_y_pos_px     = mm_to_px(75 * AppMain.height_sf_mm, gg.PPI)
    rando_analysis_x_pos_px       = mm_to_px(10 * AppMain.width_sf_mm,  gg.PPI)
    rando_analysis_y_pos_px       = mm_to_px(120 * AppMain.height_sf_mm, gg.PPI)
    launch_dx_px             = mm_to_px( 0 * AppMain.width_sf_mm,  gg.PPI)
    launch_dy_px             = mm_to_px( 5 * AppMain.height_sf_mm, gg.PPI)

    year_button_x_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_X_MM * AppMain.width_sf_mm,  gg.PPI)
    year_button_y_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)
    sorties_button_x_pos     = mm_to_px(gg.REF_SORTIES_BUT_POS_X_MM * AppMain.width_sf_mm,  gg.PPI)
    sorties_button_y_pos     = mm_to_px(gg.REF_SORTIES_BUT_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)
    dy_year                  = -6
    ds_year                  = 5
    dy_sorties               = -6

    # Setting common attributs
    etape_label_format = 'left'
    etape_underline    = -1

    ### Choix de l'année
    list_years = last_available_years(ctg_path,year_number=2021)
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
                                text = gg.TEXT_YEAR_PI,
                                font = self.font_Label_years)
    self.Label_years.place(x = year_button_x_pos, y = year_button_y_pos)

    place_after(self.Label_years, self.OptionButton_years, dy = dy_year)

    ### Choix de du type de sortie
    list_sorties = get_available_sorties(ctg_path,variable_years.get())
    default_sortie = list_sorties[-1]
    variable_sortie = tk.StringVar(self)
    variable_sortie.set(default_sortie)

        # Création de l'option button des sorties
    self.font_OptionButton_sorties = tkFont.Font(family = gg.FONT_NAME,
                                                 size = eff_buttons_font_size)
    self.OptionButton_sorties = tk.OptionMenu(self,
                                              variable_sortie,
                                              *list_sorties)
    self.OptionButton_sorties.config(font = self.font_OptionButton_sorties)

        # Création du label
    self.font_Label_sorties = tkFont.Font(family = gg.FONT_NAME,
                                        size = eff_select_font_size,
                                        weight = 'bold')
    self.Label_sorties = tk.Label(self,
                                  text = gg.TEXT_SORTIES_PI,
                                  font = self.font_Label_sorties)
    self.Label_sorties.place(x = sorties_button_x_pos, y = sorties_button_y_pos)

    place_after(self.Label_sorties, self.OptionButton_sorties, dy = dy_sorties)

    ################## Analyse des sorties

    ### Titre
    sorties_analysis_font = tkFont.Font(family = gg.FONT_NAME,
                                        size = eff_etape_font_size,
                                        weight = 'bold')
    sorties_analysis_label = tk.Label(self,
                                      text = gg.TEXT_SORTIES,
                                      justify = etape_label_format,
                                      font = sorties_analysis_font,
                                      underline = etape_underline)

    sorties_analysis_label.place(x = sorties_analysis_x_pos_px,
                                 y = sorties_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label = tk.Label(self,
                          text = gg.HELP_SORTIES,
                          justify = "left",
                          font = help_label_font)
    place_bellow(sorties_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des sorties
    sorties_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                          size = eff_launch_font_size)
    sorties_analysis_launch_button = tk.Button(self,
                                               text = gg.BUTT_SORTIES,
                                               font = sorties_analysis_launch_font,
                                               command = lambda: _launch_sorties_analysis_try())
    place_bellow(help_label,
                 sorties_analysis_launch_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)

    ################## Analyse des randonnées

    ### Titre
    rando_analysis_font = tkFont.Font(family = gg.FONT_NAME,
                                      size = eff_etape_font_size,
                                      weight = 'bold')
    rando_analysis_label = tk.Label(self,
                                    text = gg.TEXT_RANDO,
                                    justify = etape_label_format,
                                    font = rando_analysis_font,
                                    underline = etape_underline)

    rando_analysis_label.place(x = rando_analysis_x_pos_px,
                                 y = rando_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label_rando = tk.Label(self,
                                text = gg.HELP_RANDO,
                                justify = "left",
                                font = help_label_font)
    place_bellow(rando_analysis_label,
                 help_label_rando)

    ### Bouton pour lancer l'analyse des sorties
    rando_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                             size = eff_launch_font_size)
    rando_analysis_launch_button = tk.Button(self,
                                             text = gg.BUTT_RANDO,
                                             font = sorties_analysis_launch_font,
                                             command = lambda: _launch_rando_analysis_try())
    place_bellow(help_label_rando,
                 rando_analysis_launch_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)