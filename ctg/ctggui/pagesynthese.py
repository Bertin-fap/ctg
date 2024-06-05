__all__ = ['create_synthese_analysis']

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
from ctg.ctggui.guitools import encadre_RL
from ctg.ctggui.guitools import font_size
from ctg.ctggui.guitools import mm_to_px
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import place_bellow
from ctg.ctggui.guitools import last_available_years
from ctg.ctgfuncts.ctg_synthese import synthese
from ctg.ctgfuncts.ctg_synthese import plot_pie_synthese
from ctg.ctgfuncts.ctg_synthese import synthese_adherent

def create_synthese_analysis(self, master, page_name, institute, ctg_path):

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
    def _launch_synthese_analysis():

        # Getting year selection
        year_select =  variable_years.get()

        synthese(year_select,ctg_path)
        plot_pie_synthese(year_select,ctg_path)

    def _launch_member_analysis(ctg_path):

        # Getting year selection
        year_select =  variable_years.get()

        synthese_adherent(year_select,ctg_path)
        info_title = "- Information -"
        info_text  = ("L'analyse de la participation aux sorties par membre "
                     f"a été effectuée pour l'année {year_select} "
                     f"\n\nLe fichier EXCEL obtenu : 'synthese_adherent.xlsx' "
                      "a été créé dans le dossier :"
                     f"\n\n{ctg_path}\{str(year_select)}\STATISTIQUES\EXCEL.")

        messagebox.showinfo(info_title, info_text)

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
                                text = gg.TEXT_YEAR_PI,
                                font = self.font_Label_years)
    self.Label_years.place(x = year_button_x_pos, y = year_button_y_pos)

    place_after(self.Label_years, self.OptionButton_years, dy = dy_year)

    ################## Synthèse des activité

    ### Titre
    synthese_analysis_font = tkFont.Font(family = gg.FONT_NAME,
                                         size = eff_etape_font_size,
                                         weight = 'bold')
    synthese_analysis_label = tk.Label(self,
                                       text = gg.TEXT_SYNTHESE_SORTIES,
                                       justify = etape_label_format,
                                       font = synthese_analysis_font,
                                       underline = etape_underline)

    synthese_analysis_label.place(x = synthese_analysis_x_pos_px,
                                  y = synthese_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label = tk.Label(self,
                          text = gg.HELP_SYNTHESE_SORTIES,
                          justify = "left",
                          font = help_label_font)
    place_bellow(synthese_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des IFs
    synthese_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                                size = eff_launch_font_size)
    synthese_analysis_launch_button = tk.Button(self,
                                                text = gg.BUTT_SYNTHESE_SORTIES,
                                                font = synthese_analysis_launch_font,
                                                command = lambda: _launch_synthese_analysis())
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
                                     text = gg.TEXT_MEMBER_ANALYSIS,
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
                          text = gg.HELP_MEMBER_ANALYSIS,
                          justify = "left",
                          font = help_label_font)
    place_bellow(member_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des mots clefs
    member_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                                size = eff_launch_font_size)
    member_analysis_button = tk.Button(self,
                                         text = gg.BUTT_MEMBER_ANALYSIS,
                                         font = member_analysis_launch_font,
                                         command = lambda: _launch_member_analysis(ctg_path))
    place_bellow(help_label,
                 member_analysis_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)

################## Analyse de l'évolution de la moyenne d'âge

    ### Titre
    #age_analysis_label_font = tkFont.Font(family = gg.FONT_NAME,
    #                                      size = eff_etape_font_size,
    #                                      weight = 'bold')
    #age_analysis_label = tk.Label(self,
    #                              text = gg.TEXT_AGE_ANALYSIS,
    #                              justify = "left",
    #                              font = age_analysis_label_font)
    #place_bellow(member_analysis_button,
    #             age_analysis_label,
    #             dx = year_analysis_label_dx_px,
    #             dy = year_analysis_label_dy_px)
    #
    #### Explication de l'étape
    #help_label_font = tkFont.Font(family = gg.FONT_NAME,
    #                           size = eff_help_font_size)
    #help_label_age = tk.Label(self,
    #                          text = gg.HELP_AGE_ANALYSIS,
    #                          justify = "left",
    #                          font = help_label_font)
    #place_bellow(age_analysis_label,
    #             help_label_age)
    #
    #### Bouton pour lancer l'analyse des mots clefs
    #age_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
    #                                     size = eff_launch_font_size)
    #age_analysis_button = tk.Button(self,
    #                                text = gg.BUTT_AGE_ANALYSIS,
    #                                font = age_analysis_launch_font,
    #                                command = lambda: _launch_age_analysis(ctg_path))
    #place_bellow(help_label_age,
    #             age_analysis_button,
    #             dx = launch_dx_px,
    #             dy = launch_dy_px)