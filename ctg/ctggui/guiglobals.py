__all__ = ['ACTIVITE_LIST',
           'ADD_SPACE_MM',
           'BM_GUI_DISP',
           'CONTAINER_BUTTON_HEIGHT_PX',
           'FIRST_YEAR',
           'DIR_SORTIES_LIST',
           'FONT_NAME',
           'HELP_EVOLUTION_EFFECTIF',
           'HELP_MEMBER_ANALYSIS',
           'HELP_RANDO',
           'HELP_SORTIES',
           'HELP_SYNTHESE_SORTIES',
           'IN_TO_MM',
           'PAGES_LABELS',
           'PPI',
           'REF_BMF_FONT_SIZE',
           'REF_BMF_POS_X_MM',
           'REF_BMF_POS_Y_MM',
           'REF_BUTTON_DX_MM',
           'REF_BUTTON_DY_MM',
           'REF_BUTTON_FONT_SIZE',
           'REF_CHECK_BOXES_SEP_SPACE',
           'REF_COPYRIGHT_FONT_SIZE',
           'REF_COPYRIGHT_X_MM',
           'REF_COPYRIGHT_Y_MM',
           'REF_CORPI_POS_X_MM',
           'REF_CORPI_POS_Y_MM',
           'REF_ENTRY_NB_CHAR',
           'REF_ETAPE_BUT_DX_MM',
           'REF_ETAPE_BUT_DY_MM',
           'REF_ETAPE_CHECK_DY_MM',
           'REF_ETAPE_FONT_SIZE',
           'REF_ETAPE_POS_X_MM',
           'REF_ETAPE_POS_Y_MM_LIST',
           'REF_EXIT_BUT_POS_X_MM',
           'REF_EXIT_BUT_POS_Y_MM',
           'REF_INST_POS_X_MM',
           'REF_INST_POS_Y_MM',
           'REF_LABEL_FONT_SIZE',
           'REF_LABEL_POS_Y_MM',
           'REF_LAUNCH_FONT_SIZE',
           'REF_PAGE_TITLE_FONT_SIZE',
           'REF_PAGE_TITLE_POS_Y_MM',
           'REF_SCREEN_WIDTH_PX',
           'REF_SCREEN_HEIGHT_PX',
           'REF_SCREEN_WIDTH_MM',
           'REF_SCREEN_HEIGHT_MM',
           'REF_SORTIES_BUT_POS_X_MM',
           'REF_SORTIES_BUT_POS_Y_MM',
           'REF_SUB_TITLE_FONT_SIZE',
           'REF_WINDOW_WIDTH_MM',
           'REF_WINDOW_HEIGHT_MM',
           'REF_VERSION_FONT_SIZE',
           'REF_VERSION_X_MM',
           'REF_YEAR_BUT_POS_X_MM',
           'REF_YEAR_BUT_POS_Y_MM',
           'BUTT_AGE_ANALYSIS',
           'TEXT_ROOT_CTG',
           'TEXT_ROOT_CTG_CHANGE',
           'TEXT_BOUTON_CREATION_CORPUS',
           'TEXT_BOUTON_LANCEMENT',
           'BUTT_MEMBER_ANALYSIS',
           'TEXT_COPYRIGHT',
           'TEXT_CORPUSES',
           'TEXT_EVOLUTION_EFFECTIF',
           'BUTT_EFFECTIF_ANALYSIS',
           'BUTT_EVOLUTION_EFFECTIF',
           'TEXT_INSTITUTE',
           'TEXT_LAUNCH_PARSING',
           'TEXT_LAUNCH_SYNTHESE',
           'TEXT_MEMBER_ANALYSIS',
           'TEXT_PAUSE',
           'TEXT_RANDO',
           'BUTT_RANDO',
           'TEXT_SYNTHESE_SORTIES',
           'BUTT_SYNTHESE_SORTIES',
           'TEXT_SORTIES',
           'TEXT_SORTIES_PI',
           'BUTT_SORTIES',
           'TEXT_ACTIVITE_PI',
           'TEXT_TITLE',
           'TEXT_UPDATE_STATUS',
           'TEXT_VERSION',
           'TEXT_YEAR_PC',
           'TEXT_YEAR_PI',
           'TEXT_TENDANCE_SORTIES',
           'HELP_TENDANCE_SORTIES',
           'BUTT_TENDANCE_SORTIES',
           'TEXT_PRESENCE_EFFECTIF',
           'HELP_PRESENCE_EFFECTIF',
           'BUTT_PRESENCE_EFFECTIF',
           'TEXT_GEO_PI',
           'TEXT_GEO',
           'HELP_GEO',
           'BUTT_GEO',
           'TEXT_NBR_SEJOURS',
           'HELP_NBR_SEJOURS',
           'BUTT_NBR_SEJOURS',
           'TEXT_SEJOUR',
           'BUTT_SEJOUR',
           'HELP_SEJOUR',]
           
# Standard HELP_SEJOURlibrary imports
import math           

# 3rd party imports
from screeninfo import get_monitors

################################## General globals ##################################

# Setting BiblioMeter version value (internal)
VERSION ='5.0.0'

# Setting the first available year -1 for activity following
FIRST_YEAR = 2021

# Setting the title of the application main window (internal)
APPLICATION_WINDOW_TITLE = f"CTG_Meter - Analyse des statistiques des effectifs et des sorties"

######################## Definition of display globals ###########################

def _get_displays(in_to_mm):

    ''' The function `get_displays` allows to identify the set of displays
        available within the user hardware and to get their parameters.
        If the width or the height of a display are not available in mm
        through the `get_monitors` method (as for Darwin platforms),
        the user is asked to specify the displays diagonal size to compute them.

    Returns:
        `list`: list of dicts with one dict per detected display,
                each dict is keyed by 8 display parameters.
    '''
    # To Do: convert prints and inputs to gui displays and inputs

    displays = [{'x':m.x,'y':m.y,'width':m.width,
                 'height':m.height,'width_mm':m.width_mm,
                 'height_mm':m.height_mm,'name':m.name,
                 'is_primary':m.is_primary} for m in get_monitors()]

    for disp in range(len(displays)):
        width_px = displays[disp]['width']
        height_px = displays[disp]['height']
        diag_px = math.sqrt(int(width_px)**2 + int(height_px)**2)
        width_mm = displays[disp]['width_mm']
        height_mm = displays[disp]['height_mm']
        if width_mm is None or height_mm is None:
            diag_in = float(input(('Enter the diagonal size of the screen n°'
                                   f'{str(disp)} (inches)')))
            width_mm = round(int(width_px) * (diag_in/diag_px) * in_to_mm,1)
            height_mm = round(int(height_px) * (diag_in/diag_px) * in_to_mm,1)
            displays[disp]['width_mm'] = str(width_mm)
            displays[disp]['height_mm'] = str(height_mm)
        else:
            diag_in = math.sqrt(float(width_mm) ** 2 + float(height_mm) ** 2) / in_to_mm
        displays[disp]['ppi'] = round(diag_px/diag_in,2)

    return displays

# Conversion factor for inch to millimeter
IN_TO_MM = 25.4

DISPLAYS = _get_displays(IN_TO_MM)

# Setting primary display
BM_GUI_DISP = 0

# Getting display resolution in pixels per inch
PPI = DISPLAYS[BM_GUI_DISP]['ppi']

# Setting display reference sizes in pixels and mm (internal)
REF_SCREEN_WIDTH_PX       = 1920
REF_SCREEN_HEIGHT_PX      = 1080
REF_SCREEN_WIDTH_MM       = 467
REF_SCREEN_HEIGHT_MM      = 267

# Application window reference sizes in mm for the display reference sizes (internal)
REF_WINDOW_WIDTH_MM       = 219
REF_WINDOW_HEIGHT_MM      = 173


####################################################################
##########################  Pages globals ##########################
####################################################################


# Setting general globals for text edition
FONT_NAME = "Helvetica"

########################## Reference coordinates for pages ##########################

# Number of characters reference for editing the entered files-folder path
REF_ENTRY_NB_CHAR          = 100     #100

# Font size references for page label and button
REF_SUB_TITLE_FONT_SIZE    = 15      #15
REF_PAGE_TITLE_FONT_SIZE   = 30      #30
REF_LAUNCH_FONT_SIZE       = 25      #25
REF_BMF_FONT_SIZE          = 15      #15        #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
REF_BUTTON_FONT_SIZE       = 12      #10
REF_COPYRIGHT_FONT_SIZE    = 12      #10
REF_VERSION_FONT_SIZE      = 12      #10

# Y position reference in mm for page label
REF_PAGE_TITLE_POS_Y_MM    = 20      #20

# Positions reference in mm for institute selection button
REF_INST_POS_X_MM          = 5       #5
REF_INST_POS_Y_MM          = 40      #45

# Positions reference in mm for bmf label and button
REF_BMF_POS_X_MM           = 5       #5
REF_BMF_POS_Y_MM           = 55      #45
REF_BUTTON_DX_MM           = -147    #-147
REF_BUTTON_DY_MM           = 10      #10

# Positions reference in mm for corpus creation button
REF_CORPI_POS_X_MM         = 5       #5
REF_CORPI_POS_Y_MM         = 85      #75

# Space between label and value
ADD_SPACE_MM               = 10      #10

# Setting X and Y positions reference in mm for copyright
REF_COPYRIGHT_X_MM         = 5       #5
REF_COPYRIGHT_Y_MM         = 170     #170
REF_VERSION_X_MM           = 185     #185

# Container button height in pixels
CONTAINER_BUTTON_HEIGHT_PX = 50      #50

# Font size references for page label and button
REF_LABEL_FONT_SIZE        = 25      #25
REF_ETAPE_FONT_SIZE        = 14      #14
REF_BUTTON_FONT_SIZE       = 10      #10

# Positions reference in mm for pages widgets
REF_LABEL_POS_Y_MM         = 15      #15
REF_ETAPE_POS_X_MM         = 10      #10
REF_ETAPE_POS_Y_MM_LIST    = [40, 74, 101, 129]   #[40, 74, 101, 129]
REF_ETAPE_BUT_DX_MM        = 10      #10
REF_ETAPE_BUT_DY_MM        = 5       #5
REF_ETAPE_CHECK_DY_MM      = -8      #-8
REF_EXIT_BUT_POS_X_MM      = 193     #193
REF_EXIT_BUT_POS_Y_MM      = 145     #145
REF_YEAR_BUT_POS_X_MM      = 10      #10
REF_YEAR_BUT_POS_Y_MM      = 26      #26
REF_SORTIES_BUT_POS_X_MM   = 10      #10
REF_SORTIES_BUT_POS_Y_MM   = 50      #26

# Separation space in mm for check boxes
REF_CHECK_BOXES_SEP_SPACE  = 25      #25

# Setting label for each gui page
PAGES_LABELS = {'page_effectif': "Analyse des effectifs",
                'pagesorties' : "Analyse des sorties",
                'pagesynthese': "Synthèse des sorties et effectifs",
                'pagetendance': "Analyse de l'évolution temporelle des pratiques ",
                'pagetendance': "Analyse tendancielle",
                'pagedivers'  : "Autres analyses",
                'PageHelp'    : "Aide",
               }


########################## Cover Page (BiblioMeter launching Page) ##########################

# Titre de la page
TEXT_TITLE                  = "- CTG_Meter -\nInitialisation de l'analyse"

# Choix de l'année de l'Institut
TEXT_INSTITUTE              = "Sélection de l'Institut"

# Titre LabelEntry of BiblioMeter_Files folder
TEXT_ROOT_CTG                    = "Dossier de travail "

# Titre bouton changement de dossier de travail
TEXT_ROOT_CTG_CHANGE             = "Changer de dossier de travail"

# Titre liste des corpus analysés
TEXT_CORPUSES               = "Liste des années "

# Titre bouton création d'un dossier nouveau de corpus
TEXT_BOUTON_CREATION_CORPUS = "Créer une nouvelle année"

# Titre bouton de lancement
TEXT_BOUTON_LANCEMENT       = "Lancer l'analyse"

# Copyright and contacts
TEXT_COPYRIGHT              =   "Contributeurs et contacts :"
TEXT_COPYRIGHT             +=  "\n- Amal Chabli"
TEXT_COPYRIGHT             +=  "\n- François Bertin : francois.bertin7@wanadoo.fr"
TEXT_COPYRIGHT             +=  "\n- Ludovic Desmeuzes"
TEXT_VERSION                = f"\nVersion {VERSION}"

##########################  Secondary pages ##########################

# Common to secondary pages
TEXT_PAUSE           = "Mettre en pause"

###############  Parsing page

# - Label ANNEE
TEXT_YEAR_PC         = "Sélection de l'année "

# - Bouton mise à jour du statut des fichiers
TEXT_UPDATE_STATUS   = "Mettre à jour le statut des fichiers"

# - Bouton lancement parsing
TEXT_LAUNCH_PARSING  = "Lancer le Parsing"

# - Bouton lancement concatenation et deduplication des parsings
TEXT_LAUNCH_SYNTHESE = "Lancer la synthèse"


###############  Consolidation page

# Choix de l'année de travail
TEXT_YEAR_PI       = "Sélection de l'année "

TEXT_SORTIES_PI       = "Sélection du type de sortie "


###############  Analysis page

### - Page effectif
TEXT_EFFECTIF     = "Analyse de l'effectif"
HELP_EFFECTIF     = " L'analyse des effectifs est effectuée à partir des fichiers"
HELP_EFFECTIF    += " EXCEL ffcyclo.org."
BUTT_EFFECTIF_ANALYSIS = "Lancer l'analyse des effectifs"
TEXT_EVOLUTION_EFFECTIF     = "Analyse de l'évolution temporelle de l'effectif"
HELP_EVOLUTION_EFFECTIF     = " L'analyse de l'évolution de l'effectif est issue"
HELP_EVOLUTION_EFFECTIF    += " des archives disponnibles à partir de 2012."
BUTT_EVOLUTION_EFFECTIF = "Lancer l'analyse de l'évolution des effectifs"
TEXT_AGE_ANALYSIS     = "Analyse de l'évolution de l'âge médian"
HELP_AGE_ANALYSIS     = " L'analyse de l'évolution de l'âge médian"
HELP_AGE_ANALYSIS    += " des archives disponnibles à partir de 2012."
BUTT_AGE_ANALYSIS = "Lancer l'analyse de l'évolution de l'âge médian"

### - sorties
TEXT_SORTIES     = "Analyse des sorties"
HELP_SORTIES     = " L'analyse des sorties est effectuée à partir des fichiers"
HELP_SORTIES    += " csv issues des Framadate"
BUTT_SORTIES = "Lancer l'analyse des sorties"
TEXT_RANDO       = "Analyse des randonnées/séjours"
HELP_RANDO       = " L'analyse des sorties/séjours est effectuée à partir des fichers"
HELP_RANDO      += " d'émargements des participants aux randonnées/séjours"
BUTT_RANDO = "Lancer l'analyse"

### - Page synthese
TEXT_SYNTHESE_SORTIES     = "Synthèse des sorties"
HELP_SYNTHESE_SORTIES     = " La synthèse est effectuée à partir des fichiers"
HELP_SYNTHESE_SORTIES    += " d'émargement aux sorties, randonnées et séjours"
BUTT_SYNTHESE_SORTIES = "Lancer la synthèse"
TEXT_EVOLUTION_EFFECTIF     = "Analyse de l'évolution temporelle de l'effectif"
HELP_EVOLUTION_EFFECTIF     = " L'analyse de l'évolution de l'effectif est issue"
HELP_EVOLUTION_EFFECTIF    += " des archives disponnibles à partir de 2012."
BUTT_EVOLUTION_EFFECTIF = "Lancer l'analyse de l'évolution des effectifs"
TEXT_MEMBER_ANALYSIS    = "Analyse des sorties par adhérent"
HELP_MEMBER_ANALYSIS    = " Analyse de la participation aux sorties, randonnées et séjour "
HELP_MEMBER_ANALYSIS   += " par adhérent. Un fichier EXCEL est généré."
BUTT_MEMBER_ANALYSIS    = "Construction du fichier EXCEL"

### - Page analyse tendancielle
TEXT_ACTIVITE_PI = "Choix de l'activité"
TEXT_TENDANCE_SORTIES     = "Analyse tendancielle des sorties"
HELP_TENDANCE_SORTIES     = (" La synthèse est effectuée à partir des fichiers"
                             " d'émargement aux sorties, randonnées et séjours")
BUTT_TENDANCE_SORTIES     = "Lancer l'analyse"
TEXT_PRESENCE_EFFECTIF    = "Analyse de la présence an club"
HELP_PRESENCE_EFFECTIF    = (" L'analyse de l'évolution de l'effectif est issue"
                             " des archives disponnibles à partir de 2012.")
BUTT_PRESENCE_EFFECTIF    = "Lancer l'analyse de l'évolution des effectifs"
TEXT_VAE_ANALYSIS         = "Analyse tendancielle depuis 2019 de la population de VAE"
HELP_VAE_ANALYSIS         = (" L'analyse de l'évolution de l'effectif est issue"
                             " des fichiers EXCEL https://ffcyclo.org/")
BUTT_VAE_ANALYSIS         = "Lancer l'analyse de l'évolution des VAE"

### - Page divers
TEXT_GEO_PI               = "Choix de l'année"
TEXT_GEO                  = "Analyse de la répartition géographique des adhérents"
HELP_GEO                  = (" La synthèse est effectuée à partir des fichiers"
                             " EXCEL extraits de https://cyclo38ffct.fr/")
BUTT_GEO                  = "Lancer la construction du fichier html"
TEXT_NBR_SEJOURS          = "Histogramme de nombre de membres ayant participé à N séjours"
HELP_NBR_SEJOURS    = (" La construction de l'histogramme est issue"
                       " de l'anlyse du fichier EXCEL : 'synthese_adherent.xl'.")
BUTT_NBR_SEJOURS = "Lancer la construction de l'histogramme"

### - Page analyse tendancielle
TEXT_SEJOUR               = "Analyse des sejours en termes de cout et de duree"
HELP_SEJOUR               = (" La synthèse est effectuée à partir du fichiers"
                             " EXCEL synthese_adherents.xlsx")
BUTT_SEJOUR               = "Lancer l'analyse"



################# Liste répertoires
DIR_SORTIES_LIST = ['SEJOUR',
                    'SORTIES DE DERNIERE MINUTE',
                    'SORTIES DU DIMANCHE',
                    'SORTIES DU JEUDI',
                    'SORTIES DU SAMEDI',
                    'SORTIES HIVER',
                    'SORTIES VTT']

ACTIVITE_LIST =    ['nbr_participations_sejours',
                    'nbr_jours_participation_sejours',
                    'sortie_dimanche_club',
                    'sortie_samedi_club',
                    'sortie_hiver_club',
                    'sortie_jeudi_club',
                    'randonnee',
                    'nbr_sejours',
                    'nbr_jours_sejours']