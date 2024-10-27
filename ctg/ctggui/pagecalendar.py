__all__ = ['create_calendar']

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

# Internal imports
import ctg.ctggui.guiglobals as gg
from ctg.ctgfuncts.ctg_effectif import statistique_vae
from ctg.ctgfuncts.ctg_effectif import builds_excel_presence_au_club
from ctg.ctgfuncts.ctg_effectif import anciennete_au_club
from ctg.ctggui.guitools import encadre_RL
from ctg.ctggui.guitools import font_size
from ctg.ctggui.guitools import mm_to_px
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow
from ctg.ctggui.guitools import last_available_years
from ctg.ctgfuncts.ctg_synthese import evolution_sorties
from ctg.ctgfuncts.make_calendar import make_calendar
from ctg.ctgfuncts.excel2word import make_calendrier

def create_calendar(self, master, page_name, institute, ctg_path):

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
    def _build_calendar(ctg_path):

        # Getting year selection
        year =  variable_year.get()
        make_calendar(int(year),ctg_path)

    def _excel2word_button(ctg_path):
        year =  variable_year.get()
        make_calendrier(year,ctg_path)

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
    list_years = list(range(2024,2035))
    default_year = list_years[0]
    variable_year = tk.StringVar(self)
    variable_year.set(default_year)

        # Création de l'option button des activités
    self.font_OptionButton_years = tkFont.Font(family = gg.FONT_NAME,
                                           size = eff_buttons_font_size)
    self.OptionButton_years = tk.OptionMenu(self,
                                         variable_year,
                                         *list_years)
    self.OptionButton_years.config(font = self.font_OptionButton_years)

        # Création du label
    self.font_Label_years = tkFont.Font(family = gg.FONT_NAME,
                                        size = eff_select_font_size,
                                        weight = 'bold')
    # gg.TEXT_ACTIVITE_PI,
    self.Label_years = tk.Label(self,
                              text = "Choix de l'année", 
                              font = self.font_Label_years)
    self.Label_years.place(x = year_button_x_pos, y = year_button_y_pos)

    place_after(self.Label_years, self.OptionButton_years, dy = dy_year)

    ################## Synthèse des activités

    ### Titre
    build_calendar_font = tkFont.Font(family = gg.FONT_NAME,
                                         size = eff_etape_font_size,
                                         weight = 'bold')
    #gg.TEXT_TENDANCE_SORTIES
    build_calendar_label = tk.Label(self,
                                 text = "Contruction de la grille Excel",
                                 justify = etape_label_format,
                                 font = build_calendar_font,
                                 underline = etape_underline)

    build_calendar_label.place(x = synthese_analysis_x_pos_px,
                             y = synthese_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    #gg.HELP_TENDANCE_SORTIES
    help_label = tk.Label(self,
                        text = "La grille Excel est utilisée pour la construction du calendrier",
                        justify = "left",
                        font = help_label_font)
    place_bellow(build_calendar_label,
                help_label)

    ### Bouton pour lancer l'analyse tendancielle des activités
    build_calendar_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                                size = eff_launch_font_size)
    #gg.BUTT_TENDANCE_SORTIES
    build_calendar_launch_button = tk.Button(self,
                                         text = "Lancer la construction",
                                         font = build_calendar_launch_font,
                                         command = lambda: _build_calendar(ctg_path))
    place_bellow(help_label,
                build_calendar_launch_button,
                dx = launch_dx_px,
                dy = launch_dy_px)

    ################## Construction du calendrier Word

    #gg.TEXT_PRESENCE_EFFECTIF
    excel2word_font = tkFont.Font(family = gg.FONT_NAME,
                                size = eff_etape_font_size,
                                weight = 'bold')
    excel2word_label = tk.Label(self,
                              text = "construction du calendrier Word",
                              justify = "left",
                              font = excel2word_font)
    place_bellow(build_calendar_launch_button,
                excel2word_label,
                dx = year_analysis_label_dx_px,
                dy = year_analysis_label_dy_px)

    ### Explication de l'étape
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    #gg.HELP_PRESENCE_EFFECTIF
    help_label = tk.Label(self,
                        text = "Construction du calendrier Word à partir des descriptifs Excel",
                        justify = "left",
                        font = help_label_font)
    place_bellow(excel2word_label,
                help_label)

    ### Bouton pour lancer l'analyse des mots clefs
    excel2word_launch_font = tkFont.Font(family = gg.FONT_NAME,
    
                                                size = eff_launch_font_size)
    #gg.BUTT_PRESENCE_EFFECTIF
    excel2word_button = tk.Button(self,
                               text = "Lancer la construction du fichier Word",
                               font = excel2word_launch_font,
                               command = lambda: _excel2word_button(ctg_path))
    place_bellow(help_label,
                excel2word_button,
                dx = launch_dx_px,
                dy = launch_dy_px)

