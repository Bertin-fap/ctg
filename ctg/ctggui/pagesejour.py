__all__ = ['create_effectif_analysis']

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
from ctg.ctgfuncts.ctg_synthese import stat_cout_sejour

def _launch_sejour_year_analysis(ctg_path,
                        year_select):
    """
    """

    stat_cout_sejour(year_select,ctg_path)


def create_sejour_analysis(self, master, page_name, institute, ctg_path):

    """
    
    """

    # Internal functions
    def _launch_sejour_year_analysis_try():
        # Getting year selection
        year_select =  variable_years.get()

        _launch_sejour_year_analysis(ctg_path,
                            year_select)
        return

    from ctg.ctggui.pageclasses import AppMain

    # Setting effective font sizes and positions (numbers are reference values)
    eff_etape_font_size      = font_size(gg.REF_ETAPE_FONT_SIZE,   AppMain.width_sf_min)
    eff_launch_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE-1, AppMain.width_sf_min)
    eff_help_font_size       = font_size(gg.REF_ETAPE_FONT_SIZE-2, AppMain.width_sf_min)
    eff_select_font_size     = font_size(gg.REF_ETAPE_FONT_SIZE, AppMain.width_sf_min)
    eff_buttons_font_size    = font_size(gg.REF_ETAPE_FONT_SIZE-3, AppMain.width_sf_min)

    if_analysis_x_pos_px     = mm_to_px(10 * AppMain.width_sf_mm,  gg.PPI)
    if_analysis_y_pos_px     = mm_to_px(40 * AppMain.height_sf_mm, gg.PPI)
    year_analysis_label_dx_px  = mm_to_px( 0 * AppMain.width_sf_mm,  gg.PPI)
    year_analysis_label_dy_px  = mm_to_px(15 * AppMain.height_sf_mm, gg.PPI)
    launch_dx_px             = mm_to_px( 0 * AppMain.width_sf_mm,  gg.PPI)
    launch_dy_px             = mm_to_px( 5 * AppMain.height_sf_mm, gg.PPI)

    year_button_x_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_X_MM * AppMain.width_sf_mm,  gg.PPI)
    year_button_y_pos        = mm_to_px(gg.REF_YEAR_BUT_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)    #26
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

    ################## Analyse des sejours

    ### Titre
    if_analysis_font = tkFont.Font(family = gg.FONT_NAME,
                                   size = eff_etape_font_size,
                                   weight = 'bold')
    if_analysis_label = tk.Label(self,
                                 text = gg.TEXT_SEJOUR,
                                 justify = etape_label_format,
                                 font = if_analysis_font,
                                 underline = etape_underline)

    if_analysis_label.place(x = if_analysis_x_pos_px,
                            y = if_analysis_y_pos_px)

    ### Explication
    help_label_font = tkFont.Font(family = gg.FONT_NAME,
                                  size = eff_help_font_size)
    help_label = tk.Label(self,
                          text = gg.HELP_SEJOUR,
                          justify = "left",
                          font = help_label_font)
    place_bellow(if_analysis_label,
                 help_label)

    ### Bouton pour lancer l'analyse des IFs
    if_analysis_launch_font = tkFont.Font(family = gg.FONT_NAME,
                                          size = eff_launch_font_size)
    if_analysis_launch_button = tk.Button(self,
                                          text = gg.BUTT_SEJOUR,
                                          font = if_analysis_launch_font,
                                          command = lambda: _launch_sejour_year_analysis_try())
    place_bellow(help_label,
                 if_analysis_launch_button,
                 dx = launch_dx_px,
                 dy = launch_dy_px)
