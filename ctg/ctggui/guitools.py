__all__ = ['encadre_RL',
           'encadre_UD',
           'font_size',
           'get_available_sorties',
           'last_available_years',
           'mm_to_px',
           'place_after',
           'place_bellow',
           'place_bellow_LabelEntry',
           'str_size_mm',
           'show_frame',
           'create_archi',
           'window_properties'
          ]

# Standard library imports
import datetime
import itertools
import os
import re
from calendar import monthrange
from pathlib import Path
from tkinter import messagebox

# Standard library imports
import math

# Local imports
import ctg.ctggui.guiglobals as gg
from ctg.ctggui.guiglobals import IN_TO_MM
from ctg.ctggui.guiglobals import DIR_SORTIES_LIST


def show_frame(self, page_name):
    '''Show a frame for the given page name'''
    frame = self.frames[page_name]
    frame.tkraise()


def last_available_years(ctg_path, year_number=2021):

    '''
    Returns a list of the available five last available years where corpuses are stored
    '''

    # Récupérer les corpus disponibles TO DO : consolider le choix des années
    try:
        list_dir = os.listdir(ctg_path)
        years_full_list = list()
        for year in list_dir:
            if re.findall(r'^\d{4}$',year):
                years_full_list.append(year)
        if year_number is not None:
            years_list = [year for year in years_full_list if int(year)>year_number]
    except FileNotFoundError:
        warning_title = "!!! ATTENTION : Dossier de travail inaccessible !!!"
        warning_text  = f"L'accès au dossier {ctg_path} est impossible."
        warning_text += f"\nChoisissez un autre dossier de travail."
        messagebox.showwarning(warning_title, warning_text)
        years_list = []
    return years_list

def place_after(gauche, droite, dx = 5, dy = 0):
    gauche_info = gauche.place_info()
    x = int(gauche_info['x']) + gauche.winfo_reqwidth() + dx
    y = int(gauche_info['y']) + dy
    droite.place(x = x, y = y)

def place_bellow(haut, bas, dx = 0, dy = 5):
    haut_info = haut.place_info()
    x = int(haut_info['x']) + dx
    y = int(haut_info['y']) + haut.winfo_reqheight() + dy
    bas.place(x = x, y = y)

def encadre_RL(fond, gauche, droite, color = "red", dn = 10, de = 10, ds = 10, dw = 10):

    gauche_info = gauche.place_info()
    droite_info = droite.place_info()

    x1 = int(gauche_info['x']) - dw
    y1 = int(gauche_info['y']) - dn

    x2 = int(gauche_info['x']) + gauche.winfo_reqwidth() + droite.winfo_reqwidth() + de
    y2 = int(droite_info['y']) + max(gauche.winfo_reqheight(), droite.winfo_reqheight()) + ds

    rectangle = fond.create_rectangle(x1, y1, x2, y2, outline = color, width = 2)
    fond.place(x = 0, y = 0)

def encadre_UD(fond, haut, bas, color = "red", dn = 10, de = 10, ds = 10, dw = 10):

    haut_info = haut.place_info()
    bas_info = bas.place_info()

    x1 = int(haut_info['x']) - dw
    y1 = int(haut_info['y']) - dn

    x2 = int(bas_info['x']) + max(haut.winfo_reqwidth(), bas.winfo_reqwidth()) + de
    y2 = int(bas_info['y']) + haut.winfo_reqheight() + bas.winfo_reqheight() + ds

    rectangle = fond.create_rectangle(x1, y1, x2, y2, outline = color, width = 2)
    fond.place(x = 0, y = 0)

def place_bellow_LabelEntry(haut, label_entry, dx = 0, dy = 5):

    haut_info = haut.place_info()
    x = int(haut_info['x']) + dx
    y = int(haut_info['y']) + haut.winfo_reqheight() + dy
    label_entry.place(x = x, y = y)


def font_size(size, scale_factor):
    '''Set the fontsize based on scale_factor.
    If the fontsize is less than minimum_size,
    it is set to the minimum size.'''

    fontsize = int(size * scale_factor)
    if fontsize < 8:
        fontsize = 8
    return fontsize


def str_size_mm(text, font, ppi):
    '''The function `_str_size_mm` computes the sizes in mm of a string.

    Args:
        text (str): the text of which we compute the size in mm.
        font (tk.font): the font of the text.
        ppi (int): pixels per inch of the display.

    Returns:
        `(tuple)`: width in mm `(float)`, height in mm `(float)`.

    Note:
        The use of this function requires a tkinter window availability
        since it is based on a tkinter font definition.

    '''

    (w_px,h_px) = (font.measure(text),font.metrics("linespace"))
    w_mm = w_px * IN_TO_MM / ppi
    h_mm = h_px * IN_TO_MM / ppi

    return (w_mm,h_mm)


def mm_to_px(size_mm, ppi, fact = 1.0):
    '''The `mm_to_px' function converts a value in mm to a value in pixels
    using the ppi of the used display and a factor fact.

    Args:
        size_mm (float): value in mm to be converted.
        ppi ( float): pixels per inch of the display.
        fact (float): factor (default= 1).

    Returns:
        `(int)`: upper integer value of the conversion to pixels

    '''

    size_px = math.ceil((size_mm * fact / IN_TO_MM) * ppi)

    return size_px


def window_properties(screen_width_px, screen_height_px):

    # Getting number of pixels per inch screen resolution from imported global DISPLAYS
    ppi = gg.DISPLAYS[gg.BM_GUI_DISP]["ppi"]

    # Setting screen effective sizes in mm from imported global DISPLAYS
    screen_width_mm  = gg.DISPLAYS[gg.BM_GUI_DISP]["width_mm"]
    screen_height_mm = gg.DISPLAYS[gg.BM_GUI_DISP]["height_mm"]

    # Setting screen reference sizes in pixels and mm 
    #from globals internal to module "Coordinates.py"
    ref_width_px  = gg.REF_SCREEN_WIDTH_PX
    ref_height_px = gg.REF_SCREEN_HEIGHT_PX
    ref_width_mm  = gg.REF_SCREEN_WIDTH_MM
    ref_height_mm = gg.REF_SCREEN_HEIGHT_MM

    # Setting secondary window reference sizes in mm 
    #from globals internal to module "Coordinates.py"
    ref_window_width_mm  = gg.REF_WINDOW_WIDTH_MM
    ref_window_height_mm = gg.REF_WINDOW_HEIGHT_MM

    # Computing ratii of effective screen sizes to screen reference sizes in pixels
    scale_factor_width_px  = screen_width_px / ref_width_px
    scale_factor_height_px = screen_height_px / ref_height_px

    # Computing ratii of effective screen sizes to screen reference sizes in mm
    scale_factor_width_mm  = screen_width_mm / ref_width_mm
    scale_factor_height_mm = screen_height_mm / ref_height_mm

    # Computing secondary window sizes in pixels depending on scale factors
    win_width_px  = mm_to_px(ref_window_width_mm * scale_factor_width_mm, ppi)
    win_height_px = mm_to_px(ref_window_height_mm * scale_factor_height_mm, ppi)

    sizes_tuple = (win_width_px, win_height_px,
                   scale_factor_width_px, scale_factor_height_px,
                   scale_factor_width_mm, scale_factor_height_mm)

    return sizes_tuple

def create_folder(root_path, folder, verbose = False):

    folder_path = root_path / Path(folder)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        message = f"{folder_path} created"
    else:
        message = f"{folder_path} already exists"
    if verbose : print(message)
    return folder_path

def create_archi(ctg_path, year_folder, verbose = False):
    '''The `create_archi` function creates a corpus folder with the required architecture.
    It uses the global "ARCHI_YEAR" for the names of the sub_folders.

    Args:
        ctg_path (path): The full path of the working folder.
        year_folder (str): The name of the folder of the corpus.

    Returns:
        (str): The message giving the folder creation status.

    '''

    dir_list = ['DATA','SEJOUR', 'SORTIES DE DERNIERE MINUTE',
                'SORTIES DU DIMANCHE', 'SORTIES DU JEUDI',
                'SORTIES DU SAMEDI','SORTIES HIVER',
                'SORTIES VTT', 'STATISTIQUES']

    new_year_folder_path = create_folder(ctg_path, year_folder, verbose = verbose)
    for dir in dir_list:
        _ = create_folder(new_year_folder_path, dir, verbose = verbose)

    dir_list_sorties = ['SEJOUR', 'SORTIES DE DERNIERE MINUTE',
                        'SORTIES DU DIMANCHE', 'SORTIES DU JEUDI',
                        'SORTIES DU SAMEDI','SORTIES HIVER', 'SORTIES VTT']

    for dir in dir_list_sorties:
        _ = create_folder(Path(ctg_path) / Path(year_folder) / Path(dir),
                         'CSV',
                          verbose = verbose)
        _ = create_folder(Path(ctg_path) / Path(year_folder) / Path(dir),
                          'EXCEL',
                          verbose = verbose)

    jours_sortie_club(year_folder,ctg_path)
    message = f"Architecture created for {year_folder} folder"
    return message

def jours_sortie_club(year,ctg_path):

    days_dict = {0: 'dimanche',
                 1: 'Lundi',
                 2: 'mardi',
                 3: 'mercredi',
                 4: 'jeudi',
                 5: 'vendredi',
                 6: 'samedi'}
    inv_days_dict = dict(zip(days_dict.values(),days_dict.keys()))
    month_dict = dict(zip(itertools.islice(itertools.cycle(l:=range(1,13)),2,2+len(l)),l))

    root = Path(ctg_path) / Path(str(year))

    for day in ['dimanche','samedi','jeudi']:
        path = ctg_path / Path(str(year)) / Path('SORTIES DU '+day.upper()) / Path('CSV')

        r = []
        rw = []
        for m in range(1,13):
            m_save=m
            y = int(year)%100
            c = int(int(year)/100)
            m = month_dict[m]
            if m>10 : y = y-1
            _,nb_days  = monthrange(y+200*c, m_save)
            for d in range (1,nb_days+1):
                if (d + int(2.6*m - 0.2) - 2*c + y + int(c/4) + int(y/4))%7 == inv_days_dict[day]:
                    r.append(f'{str(m_save).zfill(2)}_{str(d).zfill(2)} sortie du {day}.csv')
                    try:
                        num_week = datetime.date(year,m_save, d).isocalendar().week
                    except:
                        num_week = None

                    if num_week is not None:
                        rw.append((f'{str(m_save).zfill(2)}_{str(d).zfill(2)} sortie du {day}.csv' ,
                                        num_week))

                    file_name = f'{year}_{str(m_save).zfill(2)}_{str(d).zfill(2)} sortie du {day}.csv'
                    with open(os.path.join(path,file_name), 'w') as fp:
                        fp.write(f'{str(year)} SORTIES DU {day.upper()}')

def get_available_sorties(ctg_path,year):

    year_path = ctg_path / Path(str(year))
    dir_year_list = os.listdir(year_path)

    available_sorties_list = list(set(DIR_SORTIES_LIST).intersection(dir_year_list))

    return available_sorties_list