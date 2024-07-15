__all__ = ['AppMain']

# Standard library imports
import tkinter as tk
from datetime import datetime
from functools import partial
from pathlib import Path
from tkinter import filedialog
from tkinter import messagebox
from tkinter import font as tkFont

# 3rd party imports
from screeninfo import get_monitors

# Local imports
import ctg.ctggui.guiglobals as gg
import ctg.ctggui.guitools as guf
from ctg.ctggui.guitools import font_size
from ctg.ctggui.guitools import last_available_years
from ctg.ctggui.guitools import mm_to_px
from ctg.ctggui.guitools import place_after
from ctg.ctggui.guitools import str_size_mm
from ctg.ctggui.guitools import create_archi
from ctg.ctggui.guitools import window_properties
from ctg.ctggui.page_effectif import create_effectif_analysis
from ctg.ctggui.pagesorties import create_sorties_analysis
from ctg.ctggui.pagesynthese import create_synthese_analysis
from ctg.ctggui.pagetendance import create_tendance_analysis
from ctg.ctggui.pagesejour import create_sejour_analysis
from ctg.ctggui.pagedivers import create_divers_analysis
from ctg.ctggui.pagenewactivity import get_new_activity
from ctg.ctggui.guitools import place_bellow
from ctg.ctggui.guitools import show_frame


class AppMain(tk.Tk):
    '''
    '''

    ############################### Class init - start ###############################
    def __init__(master):

        ctg_path_default = Path.home() /Path('CTG/SORTIES')

        # Setting the link with "tk.Tk"
        tk.Tk.__init__(master)


        # Identifying tk window of init class
        _ = get_monitors() # OBLIGATOIRE
        master.attributes("-topmost", True)

        # Setting the icon
        file = Path(__file__).parent.parent / Path('ctgfuncts') 
        file = file / Path('CTG_RefFiles') / Path( 'logoctg4.ico') 
        master.iconbitmap( file)

        # Setting title window
        master.title(gg.APPLICATION_WINDOW_TITLE)

        # Setting window size depending on scale factor
        screen_width_px  = master.winfo_screenwidth()
        screen_height_px = master.winfo_screenheight()

        sizes_tuple = window_properties(screen_width_px, screen_height_px)

        win_width_px  = sizes_tuple[0]
        win_height_px = sizes_tuple[1]

        master.geometry(f"{win_width_px}x{win_height_px}")
        master.resizable(False, False)

        # Defining pages classes and pages list
        AppMain.pages = ( PageHelp,
                          PageDivers,
                          PageSejours,
                          PageTendance,
                          PageSynthese,
                          PageEffectif,
                          PageSorties,
                          PageEntrerSortie,
                         )
        AppMain.pages_ordered_list = [x.__name__ for x in AppMain.pages][::-1]

        # Getting useful screen sizes and scale factors depending on displays properties
        AppMain.win_width_px  = sizes_tuple[0]
        AppMain.win_height_px = sizes_tuple[1]
        AppMain.width_sf_px   = sizes_tuple[2]
        AppMain.height_sf_px  = sizes_tuple[3]     # unused here
        AppMain.width_sf_mm   = sizes_tuple[4]
        AppMain.height_sf_mm  = sizes_tuple[5]
        AppMain.width_sf_min  = min(AppMain.width_sf_mm, AppMain.width_sf_px)


        ######################################## Title and copyright and application initialization
        PageTitle(master)
        AuthorCopyright(master)
        InitApp(master)
        SetLaunchButton(master, "CTG", ctg_path_default)


    ################################ Class init - end ################################

class InitApp(tk.Tk):

    def __init__(self,master):

        # Internal functions - start
        def _display_path(inst_bmf):
            """Shortening bmf path for easy display"""
            p = Path(inst_bmf)
            p_disp = ('/'.join(p.parts[0:2])) / Path("...") / ('/'.join(p.parts[-3:]))
            return p_disp

        def _get_file(institute_select):
            # Getting new working directory
            dialog_title = "Choisir un nouveau dossier de travail"
            bmf_str = tk.filedialog.askdirectory(title = dialog_title)
            if bmf_str == '':
                warning_title = "!!! Attention !!!"
                warning_text = "Chemin non renseigné."
                return messagebox.showwarning(warning_title, warning_text)

            # Updating bmf values using new working directory
            _set_bmf_widget_param(institute_select, bmf_str)
            _update_corpi(bmf_str)
            AppMain.institute = institute_select # FB
            AppMain.ctg_path = bmf_str # FB
            SetLaunchButton(master, institute_select, bmf_str)

        def _set_bmf_widget_param(institute_select, inst_bmf):
            # Setting bmf widgets parameters
            bmf_font = tkFont.Font(family = gg.FONT_NAME,
                                   size   = eff_bmf_font_size,
                                   weight = 'bold')
            bmf_label       = tk.Label(master,
                                       text = gg.TEXT_ROOT_CTG,
                                       font = bmf_font,)
            bmf_val         = tk.StringVar(master)
            bmf_val2        = tk.StringVar(master)
            bmf_entree      = tk.Entry(master, textvariable = bmf_val)
            bmf_entree2     = tk.Entry(master, textvariable = bmf_val2, width = eff_bmf_width)
            bmf_button_font = tkFont.Font(family = gg.FONT_NAME,
                                          size   = eff_button_font_size)
            bmf_button      = tk.Button(master,
                                        text = gg.TEXT_ROOT_CTG_CHANGE,
                                        font = bmf_button_font,
                                        command = lambda: _get_file(institute_select))
            # Placing bmf widgets
            bmf_label.place(x = eff_bmf_pos_x_px,
                            y = eff_bmf_pos_y_px,)

            text_width_mm, _ = str_size_mm(gg.TEXT_ROOT_CTG, bmf_font, gg.PPI)
            eff_path_pos_x_px = mm_to_px(text_width_mm + add_space_mm, gg.PPI)
            bmf_entree2.place(x = eff_path_pos_x_px,
                              y = eff_bmf_pos_y_px,)

            bmf_button.place(x = eff_path_pos_x_px,
                             y = eff_bmf_pos_y_px + eff_button_dy_px,)
            bmf_val.set(inst_bmf)
            bmf_val2.set((_display_path(inst_bmf)))

        def _create_corpus(inst_bmf):
            corpi_val = _set_corpi_widgets_param(inst_bmf)
            bmf_path = Path(inst_bmf)
            try:
                # Setting new corpus year folder name
                corpuses_list    = last_available_years(bmf_path, gg.FIRST_YEAR)
                last_corpus_year = corpuses_list[-1]
                current_year = datetime.now().year
                if int(last_corpus_year) == int(current_year):
                    message = ('Vous ne pouvez pas créer une nouvelle année '
                              f"tant que l'année en cours {current_year} ne sera pas "
                               "terminée.")
                    messagebox.showwarning("Warning",message)
                    corpi_val.set(str(corpuses_list))
                    return
                    
                new_corpus_year_folder = str(int(last_corpus_year) + 1)
                # Creating required folders for new corpus year
                message = create_archi(bmf_path, new_corpus_year_folder, verbose = False)
                print("\n",message)

                # Getting updated corpuses list
                corpuses_list = last_available_years(bmf_path, gg.FIRST_YEAR)

                # Setting corpi_val value to corpuses list
                corpi_val.set(str(corpuses_list))

            except FileNotFoundError:
                warning_title = "!!! ATTENTION : Dossier de travail inaccessible !!!"
                warning_text  = f"L'accès au dossier {bmf_path} est impossible."
                warning_text += f"\nChoisissez un autre dossier de travail."
                messagebox.showwarning(warning_title, warning_text)

                # Setting corpi_val value to empty string
                corpi_val.set("")

        def _set_corpi_widgets_param(inst_bmf):
            # Setting corpuses widgets parameters
            corpi_font   = tkFont.Font(family = gg.FONT_NAME,
                                       size   = eff_corpi_font_size,
                                       weight = 'bold')
            corpi_val    = tk.StringVar(master)
            corpi_entry  = tk.Entry(master, textvariable = corpi_val, width = eff_list_width)
            corpi_button_font = tkFont.Font(family = gg.FONT_NAME,
                                            size   = eff_button_font_size)
            corpi_label  = tk.Label(master,
                                    text = gg.TEXT_CORPUSES,
                                    font = corpi_font,)
            corpi_button = tk.Button(master,
                                     text = gg.TEXT_BOUTON_CREATION_CORPUS,
                                     font = corpi_button_font,
                                     command = lambda: _create_corpus(inst_bmf))
                        # Placing corpuses widgets
            corpi_label.place(x = eff_corpi_pos_x_px,
                              y = eff_corpi_pos_y_px,)

            text_width_mm, _ = str_size_mm(gg.TEXT_CORPUSES, corpi_font, gg.PPI)
            eff_list_pos_x_px = mm_to_px(text_width_mm + add_space_mm, gg.PPI)
            corpi_entry.place(x = eff_list_pos_x_px ,
                              y = eff_corpi_pos_y_px,)

            corpi_button.place(x = eff_list_pos_x_px,
                               y = eff_corpi_pos_y_px + eff_button_dy_px,)
            return corpi_val

        def _update_corpi(inst_bmf):
            corpi_val = _set_corpi_widgets_param(inst_bmf)
            bmf_path = Path(inst_bmf)
            try:
                # Getting updated corpuses list
                corpuses_list = last_available_years(bmf_path, gg.FIRST_YEAR)
                # Setting corpi_val value to corpuses list
                corpi_val.set(str(corpuses_list))

            except FileNotFoundError:
                warning_title = "!!! ATTENTION : Dossier de travail inaccessible !!!"
                warning_text  = f"L'accès au dossier {bmf_path} est impossible."
                warning_text += f"\nChoisissez un autre dossier de travail."
                messagebox.showwarning(warning_title, warning_text)

                # Setting corpi_val value to empty string
                corpi_val.set("")

        def _update_page():

            institute_select = 'CTG'
            inst_default_bmf = Path.home() / Path('CTG/SORTIES')

            # Managing working folder (bmf stands for "BiblioMeter_Files")
            _set_bmf_widget_param(institute_select, inst_default_bmf)

            # Managing corpus list
            corpi_val = _set_corpi_widgets_param(inst_default_bmf)
                # Setting and displating corpuses list initial values
            try:
                init_corpuses_list = last_available_years(inst_default_bmf, gg.FIRST_YEAR)
                corpi_val.set(str(init_corpuses_list))
            except FileNotFoundError:
                warning_title = "!!! ATTENTION : Dossier de travail inaccessible !!!"
                warning_text  = f"L'accès au dossier {inst_default_bmf} est impossible."
                warning_text += f"\nChoisissez un autre dossier de travail."
                messagebox.showwarning(warning_title, warning_text)
                # Setting corpuses list value to empty string
                corpi_val.set("")

            # Managing analysis launch button
            AppMain.institute = institute_select # FB
            AppMain.ctg_path = inst_default_bmf # FB
            SetLaunchButton(master, institute_select, inst_default_bmf)

        ############################## Internal functions - end ##############################

        ######################################## Main ########################################


        # Setting common parameters for widgets
        add_space_mm = gg.ADD_SPACE_MM
        eff_button_font_size = font_size(gg.REF_BUTTON_FONT_SIZE, AppMain.width_sf_min)

        ##################### Working-folder selection widgets parameters ######################
        # Setting effective value for bmf entry width
        eff_bmf_width = int(gg.REF_ENTRY_NB_CHAR * AppMain.width_sf_min)

        # Setting font size for bmf
        eff_bmf_font_size = font_size(gg.REF_SUB_TITLE_FONT_SIZE, AppMain.width_sf_min)

        # Setting reference positions in mm and effective ones in pixels for bmf
        eff_bmf_pos_x_px = mm_to_px(gg.REF_BMF_POS_X_MM * AppMain.height_sf_mm, gg.PPI)
        eff_bmf_pos_y_px = mm_to_px(gg.REF_BMF_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)

        # Setting reference relative positions in mm and effective relative
        # Y positions in pixels for bmf change button
        eff_button_dy_px = mm_to_px(gg.REF_BUTTON_DY_MM * AppMain.height_sf_mm, gg.PPI)

        ##################### Corpuses-list-display widgets parameters ######################
        # Setting effective value for corpi setting width
        eff_list_width = int(gg.REF_ENTRY_NB_CHAR * AppMain.width_sf_min)

        # Setting font size for corpi
        eff_corpi_font_size  = font_size(gg.REF_SUB_TITLE_FONT_SIZE, AppMain.width_sf_min)

        # Setting reference positions in mm and effective ones in pixels for corpuses
        eff_corpi_pos_x_px = mm_to_px(gg.REF_CORPI_POS_X_MM * AppMain.height_sf_mm, gg.PPI)
        eff_corpi_pos_y_px = mm_to_px(gg.REF_CORPI_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)

        _update_page()

class SetLaunchButton(tk.Tk):

    def __init__(self, master, institute, ctg_path):

        #tk.Frame.__init__(self) #FB

        # Setting font size for launch button
        eff_launch_font_size = font_size(gg.REF_LAUNCH_FONT_SIZE, AppMain.width_sf_min)

        # Setting x and y position in pixels for launch button
        launch_but_pos_x_px = AppMain.win_width_px  * 0.5
        launch_but_pos_y_px = AppMain.win_height_px * 0.8

        # Setting launch button
        launch_font = tkFont.Font(family = gg.FONT_NAME,
                                  size   = eff_launch_font_size,
                                  weight = 'bold')
        launch_button = tk.Button(master,
                                  text = gg.TEXT_BOUTON_LANCEMENT,
                                  font = launch_font,
                                  command = lambda: self._generate_pages(master,
                                                                         institute,
                                                                         ctg_path))
        # Placing launch button
        launch_button.place(x = launch_but_pos_x_px,
                            y = launch_but_pos_y_px,
                            anchor = "s")

    def _generate_pages(self, master, institute, ctg_path):

        '''Permet la génération des pages après spécification du chemin
        vers la zone de stockage.
        Vérifie qu'un chemin a été renseigné et continue le cas échant,
        sinon redemande de renseigner un chemin.
        '''

        if ctg_path == '':
            warning_title = "!!! Attention !!!"
            warning_text =  "Chemin non renseigné."
            warning_text += "\nL'application ne peut pas être lancée."
            warning_text += "\nVeuillez le définir."
            messagebox.showwarning(warning_title, warning_text)

        else:
            # Creating two frames in the tk window
            container_button = tk.Frame(master,
                                        height = gg.CONTAINER_BUTTON_HEIGHT_PX,
                                        bg = 'red')
            container_button.pack(side = "top",
                                  fill = "both",
                                  expand = False)

            container_frame = tk.Frame(master)
            container_frame.pack(side="top",
                                 fill="both",
                                 expand=True)
            container_frame.grid_rowconfigure(0,
                                              weight = 1)
            container_frame.grid_columnconfigure(0,
                                                 weight = 1)
            master.frames = {}
            for page in AppMain.pages:
                page_name = page.__name__
                frame = page(container_frame, master, container_button, institute, ctg_path)
                master.frames[page_name] = frame

                # put all of the pages in the same location;
                # the one on the top of the stacking order
                # will be the one that is visible.
                frame.grid(row = 0,
                           column = 0,
                           sticky = "nsew")


class PageTitle(tk.Tk):

    def __init__(self, parent):

        ####################### Title and copyright widgets parameters ########################
        # Setting font size for page title and copyright
        eff_page_title_font_size = guf.font_size(gg.REF_PAGE_TITLE_FONT_SIZE, AppMain.width_sf_min)

        # Setting reference Y position in mm and effective Y position in pixels for page label
        eff_page_title_pos_y_px = guf.mm_to_px(gg.REF_PAGE_TITLE_POS_Y_MM * AppMain.height_sf_mm,
                                               gg.PPI)

        # Setting x position in pixels for page title
        mid_page_pos_x_px = AppMain.win_width_px  * 0.5
        page_title = tk.Label(parent,
                              text = gg.TEXT_TITLE,
                              font = (gg.FONT_NAME, eff_page_title_font_size),
                              justify = "center")
        page_title.place(x = mid_page_pos_x_px,
                         y = eff_page_title_pos_y_px,
                         anchor = "center")

class AuthorCopyright(tk.Frame):

    def __init__(self, parent):

        # Setting font size for copyright
        ref_copyright_font_size = gg.REF_COPYRIGHT_FONT_SIZE
        eff_copyright_font_size = guf.font_size(ref_copyright_font_size, AppMain.width_sf_min)

        # Setting X and Y positions reference in mm for copyright
        ref_copyright_x_mm = gg.REF_COPYRIGHT_X_MM              #5
        eff_copyright_x_px = guf.mm_to_px(ref_copyright_x_mm * AppMain.width_sf_mm, gg.PPI)
        ref_copyright_y_mm = gg.REF_COPYRIGHT_Y_MM              #170
        eff_copyright_y_px = guf.mm_to_px(ref_copyright_y_mm * AppMain.height_sf_mm, gg.PPI)

        # Setting font size for version
        ref_version_font_size = gg.REF_VERSION_FONT_SIZE         #12
        eff_version_font_size = guf.font_size(ref_version_font_size,AppMain.width_sf_min)

        # Setting X and Y positions reference in mm for version
        ref_version_x_mm = gg.REF_VERSION_X_MM                   #185
        ref_version_y_mm = gg.REF_COPYRIGHT_Y_MM                 #170
        eff_version_y_px = guf.mm_to_px(ref_version_y_mm * AppMain.height_sf_mm, gg.PPI)
        eff_version_x_px = guf.mm_to_px(ref_version_x_mm * AppMain.width_sf_mm, gg.PPI)


        Auteurs_font_label = tkFont.Font(family = gg.FONT_NAME,
                                             size   = eff_copyright_font_size,)
        Auteurs_label = tk.Label(parent,
                                 text = gg.TEXT_COPYRIGHT,
                                 font = Auteurs_font_label,
                                 justify = "left")
        Auteurs_label.place(x = eff_copyright_x_px,
                            y = eff_copyright_y_px,
                            anchor = "sw")

        version_font_label = tkFont.Font(family = gg.FONT_NAME,
                                         size = eff_version_font_size,
                                         weight = 'bold')
        version_label = tk.Label(parent,
                                 text = gg.TEXT_VERSION,
                                 font = version_font_label,
                                 justify = "right")
        version_label.place(x = eff_version_x_px,
                            y = eff_version_y_px,
                            anchor = "sw")

class PageEntrerSortie(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        get_new_activity(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageEffectif(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_effectif_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageSejours(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_sejour_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageSorties(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_sorties_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)




class PageSynthese(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_synthese_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageTendance(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_tendance_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageDivers(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        create_divers_analysis(self, master, page_name, institute, ctg_path)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        settitleclass = SetTitltleClass(self,page_name,institute)

        quitapp = QuitApp(self, master, container_button)

class PageHelp(tk.Frame):
    '''
    '''
    def __init__(self, container_frame, master, container_button, institute, ctg_path):

        if_analysis_x_pos_px     = mm_to_px(10 * AppMain.width_sf_mm,  gg.PPI)
        if_analysis_y_pos_px     = mm_to_px(10 * AppMain.height_sf_mm, gg.PPI)

        super().__init__(container_frame)

        page_name  = self.__class__.__name__

        # Text box and y slide definition
        help_box = tk.Text(self,
                           width=90,
                           height=30,
                           highlightthickness=1,
                           foreground="black",
                           insertbackground="black",
                           font=("Helvetica", 8))

        help_box.place(x = if_analysis_x_pos_px, y = if_analysis_y_pos_px)

        tex_scroll_y = tk.Scrollbar(self,orient=tk.VERTICAL,)
        tex_scroll_y.config(command=help_box.yview, )
        help_box["yscrollcommand"] = tex_scroll_y.set
        place_after(help_box, tex_scroll_y, dx = 5, dy = 0)

        # Write into the Text box
        #version_gui = ctg.ctggui.__version__
        #help_box.insert(tk.END,f'Version : {version_gui} \n\n')
        path_help = Path(__file__).parent.parent / Path('ctgfuncts') 
        path_help = path_help / Path('CTG_RefFiles') / Path( 'help.txt')
        with open(path_help,'r',encoding="utf-8")as file_text:
            help_content =  file_text.read()
        help_box.insert(tk.END,help_content)

        setcontrollerbutton = SetControllerButton(master, container_button,page_name)

        quitapp = QuitApp(self, master, container_button)

class SetControllerButton(tk.Tk):

    def __init__(self, master, container_button,page_name):

        page_num   = AppMain.pages_ordered_list.index(page_name)
        label_text = gg.PAGES_LABELS.setdefault(page_name, page_name)
        label_button = gg.BUTTON_LABELS.setdefault(page_name, page_name)

        # Setting font size for page label and button
        eff_label_font_size  = font_size(gg.REF_LABEL_FONT_SIZE, AppMain.width_sf_min)
        eff_button_font_size = font_size(gg.REF_BUTTON_FONT_SIZE, AppMain.width_sf_min)

        button_font = tkFont.Font(family = gg.FONT_NAME,
                                  size   = eff_button_font_size)
        button = tk.Button(container_button,
                           text = label_button,
                           font = button_font,
                           command = lambda: show_frame(master, page_name))  ####FB
        button.grid(row = 0, column = page_num)


class SetTitltleClass(tk.Tk):

    def __init__(self, parent, page_name, institute):

        label_text = gg.PAGES_LABELS.setdefault(page_name, page_name)
        page_title = label_text + " du " + institute

        # Setting y_position in px for page label
        eff_label_pos_y_px = mm_to_px(gg.REF_LABEL_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)

        eff_label_font_size  = font_size(gg.REF_LABEL_FONT_SIZE, AppMain.width_sf_min)
        # Setting x position in pixels for page label
        mid_page_pos_x_px = AppMain.win_width_px / 2


        label_font = tkFont.Font(family = gg.FONT_NAME,
                                 size   = eff_label_font_size)
        label = tk.Label(parent,
                         text = page_title,
                         font = label_font)
        label.place(x = mid_page_pos_x_px,
                    y = eff_label_pos_y_px,
                    anchor = "center")


class QuitApp(tk.Frame):

    def __init__(self, parent, master, container_button):

        super().__init__(parent)
        ################## Bouton pour sortir de la page
        eff_button_font_size = font_size(gg.REF_BUTTON_FONT_SIZE, AppMain.width_sf_min)
        eff_buttons_font_size    = font_size(gg.REF_ETAPE_FONT_SIZE-3, AppMain.width_sf_min)
        exit_button_x_pos_px     = mm_to_px(gg.REF_EXIT_BUT_POS_X_MM * AppMain.width_sf_mm,  gg.PPI)
        exit_button_y_pos_px     = mm_to_px(gg.REF_EXIT_BUT_POS_Y_MM * AppMain.height_sf_mm, gg.PPI)
        quit_font = tkFont.Font(family = gg.FONT_NAME,
                                size   = eff_buttons_font_size)

        quit_button = tk.Button(parent,
                                text = gg.TEXT_PAUSE,
                                font = quit_font,
                                command = lambda: self._launch_exit(master))
        quit_button.place(x = exit_button_x_pos_px,
                          y = exit_button_y_pos_px,
                          anchor = 'n')

    def _launch_exit(self,controller):

        message =  "Vous allez fermer CTG_Meter. "
        message += "\nRien ne sera perdu et vous pourrez reprendre le traitement plus tard."
        message += "\n\nSouhaitez-vous faire une pause dans le traitement ?"
        answer_1 = messagebox.askokcancel('Information', message)
        if answer_1:
            controller.destroy()