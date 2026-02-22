__all_  = ['create_word_from_template',
           'combine_all_docx',
           'make_sejour_docx',
           'make_list_sejour',
           'make_calendrier']

# import datetime
import datetime
import functools
import json
import os
import re
import unicodedata
from pathlib import Path
from collections import defaultdict
import pathlib
from textwrap import wrap

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from docxcompose.composer import Composer
from docx import Document as Document_compose
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Emu, Inches, Mm
import qrcode

year = 2026
Path.home()
path_effectif = Path.home() / Path(r"Nextcloud2\BASE_DOCUMENTS_CTG\1_FONCTIONNEMENT_CTG\1-1_BASE_ADHERENTS_CTG")
PATH_TEMPLATES = Path.home() / Path(r"Nextcloud2\BASE_DOCUMENTS_CTG\2_ACTIVITES_CTG\2-3_ELABORATION_CALENDRIER\PRIVE")
PATH_TEMPLATES = PATH_TEMPLATES / Path(f"Template_{str(year)}")

class EffectifCtg():


    def __init__(self,year:str,path_effectif:pathlib.WindowsPath):

        
        self.year = year
        self.path_effectif = path_effectif
        
        # get effectif of the year year
        path_root = self.path_effectif / Path(str(year))
        print (path_root)
        df = pd.read_excel(path_root / Path(str(year)+'.xlsx'))
        if 'Ville' not in df.columns:
            df['Ville'] = df['Adresse'].apply(lambda row: re.split(r'\s+\d{5,6}\s+', row)[-1])
        if 'Nom' not in df.columns:
            df['Nom'] = df['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[1])
            df['Prénom'] = df['Nom, Prénom'].apply(lambda row: re.split('\s+', row)[2])
            df['Sexe'] = df['Sexe'].apply(lambda row:row[0])
            df = df.rename(columns={'N°': 'N° Licencié',})
        
        self.effectif = df      # effectif year
        self.effectif = self.correction_effectif()
        self.effectif = self.add_membres_sympathisants()
        
    def correction_effectif(self):
        path_root = self.path_effectif / Path(str(self.year))
        correction_effectif = pd.read_excel(path_root/Path('correction_effectif.xlsx'))
        correction_effectif.index = correction_effectif['N° Licencié']
        for num_licence in correction_effectif.index:
            idx = self.effectif[self.effectif["N° Licencié"]==num_licence].index
            self.effectif.loc[idx,'Prénom'] = correction_effectif.loc[num_licence,'Prénom']
            self.effectif.loc[idx,'Nom'] = correction_effectif.loc[num_licence,'Nom']
        return self.effectif

    def add_membres_sympathisants(self):
        path_root = self.path_effectif / Path(str(self.year))
        file_path = path_root / Path('membres_sympatisants.xlsx')
        
        if os.path.isfile(file_path):
            membres_sympathisants_df = pd.read_excel(file_path)
            membres_sympathisants_df = membres_sympathisants_df[['N° Licencié',
                                                                 'Nom',
                                                                 'Prénom',
                                                                 'Sexe',
                                                                 'Pratique VAE',
                                                                 'E-mail']]
            
            effectif_tot = pd.concat([self.effectif, membres_sympathisants_df], ignore_index=True, axis=0)
            effectif_tot['Prénom'] = effectif_tot['Prénom'].str.replace(' ','-')

            return effectif_tot
            
        else:
            return self.effectif
            
nfc = functools.partial(unicodedata.normalize,'NFD')
convert_to_ascii = lambda text : nfc(text). \
                                     encode('ascii', 'ignore'). \
                                     decode('utf-8').\
                                     strip()
def reads_json():

    path_json = PATH_TEMPLATES / Path("data.json")
    with open(path_json, 'r') as file:
        data = json.load(file)
    
    PATH_DATA = Path(data["PATH_DATA"])
    PATH_IMAGES = Path(data["PATH_IMAGES"])
    PATH_OUTPUT = Path(data["PATH_OUTPUT"])
    sorties_dic = data["sorties_dic"]
    sortie_vtt = data["sortie_vtt"]
    sortie_route = data["sortie_route"]
    url_dic = data["url_dic"]
    date_ag = data["date_ag"]
    color_dic = data["color_dic"]
    org = data["organisation"]
    txt_remplacement = 'à venir' #data["txt_remplacement"]

    return (PATH_DATA, PATH_IMAGES, PATH_OUTPUT, sorties_dic,
            sortie_vtt, sortie_route, url_dic, date_ag,color_dic,org,txt_remplacement)

def make_qrcode(PATH_IMAGES, url_dic):
    
    qrc_dic = {}
    for name, url in url_dic.items():
        qrcode_img = qrcode.make(url)
        qrc_name = str(PATH_IMAGES / Path(f'qrc_{name}.jpg'))
        qrcode_img.save(qrc_name, 'JPEG')
        qrc_dic[name] = qrc_name
    
    return qrc_dic


def day_of_the_date(day:int,month:int,year:int)->str:

    '''Compute the day of the week of the date after
    "Elementary Number Theory David M. Burton Chap 6.4" [1]'''

    days_dict = {0: 'dimanche',
                 1: 'lundi',
                 2: 'mardi',
                 3: 'mercredi',
                 4: 'jeudi',
                 5: 'vendredi',
                 6: 'samedi'}

    month_dict = {3: 1, 4: 2, 5: 3, 6: 4, 7: 5,
                  8: 6, 9: 7, 10: 8, 11: 9, 12:
                  10, 1: 11, 2: 12} # [1] p. 125

    y = year % 100
    c = int(year/100)
    m = month_dict[month]
    if m>10:
        y = y-1
    day_of_the_week = days_dict[(day + int(2.6*m - 0.2) -
                                2*c + y + int(c/4) +
                                int(y/4))%7] # [1] thm 6.12

    return day_of_the_week
    
def gets_day_of_the_week(x):
    month = int(x.split('/')[1])
    day = int(x.split('/')[0])
    d = day_of_the_date(day,month,year).capitalize()
    return d

def gets_day_of_the_week2(x):
    year = int('20'+str(x.split('-')[0]))
    month = int(x.split('-')[1])
    day = int(x.split('-')[2])
    d = day_of_the_date(day,month,year).capitalize()
    return d

def computes_num_week(row):
    y_m_d = row['Date début'].split('-')
    num_week = datetime.date(int('20'+y_m_d[0]),int(y_m_d[1]), int(y_m_d[2])).isocalendar().week
    return num_week  
    
def get_nom_prenom(prenom_nom):
    
    """
     Generates the list of nanes and the list of surnames of the sejour responsibles
    """

    # extract the full name of the sejour organizers
    prenom_nom = convert_to_ascii(prenom_nom).upper() 
    nom = prenom_nom.split()[1]
    prenom = prenom_nom.split()[0] 
    
    return nom, prenom

def normalize_num_tel(s):
        s= str(int(s))
        if len(s)>0:
            digits = ''.join(char for char in s if char.isdigit())
            digits = digits.zfill(10)
            return " ".join(wrap(digits, width=2))
        else:
            return ''
            
def get_tel (df, nom, prenom , year):

    """
    Gets the phone number of prenom nom from the CTG gilda list
    """

    dg = df.query('Nom==@nom and Prénom==@prenom')
    num_portable = dg["N° Portable"].tolist()[0]
    if num_portable == 0:
        num_portable = dg["Tel fixe"].tolist()[0]
    num_portable = normalize_num_tel(num_portable)
    
    return num_portable

def get_courriel(df, nom, prenom):

    """
    Gets the email of prenom nom from the CTG gilda list
    """

    dg = df.query('Nom==@nom and Prénom==@prenom')
    email = dg["Adresse email"].tolist()[0]

    return email

def read_effectif(path_effectif, year):
    
    # reads the CTG effectit Exel file of the current year
    
    df = EffectifCtg(year,path_effectif).effectif
   
    
    df["Tel fixe"] = df["Tel fixe"].fillna(0)
    df["N° Portable"] = df["Tel portable"].fillna(0)
    
    return df
    
def computes_jour(row):
    y_m_d = row['Date début'].split('-')
    jour = day_of_the_date(int(y_m_d[2]),int(y_m_d[1]),int('20'+y_m_d[0])).capitalize()
    return jour 
    
def computes_jour_deb(row):
    y_m_d = row['Date début'].split('-')
    jour = day_of_the_date(int(y_m_d[2]),int(y_m_d[1]),int('20'+y_m_d[0])).capitalize()
    return jour
    
def computes_jour_fin(row):
    y_m_d = row['Date fin'].split('-')
    jour = day_of_the_date(int(y_m_d[2]),int(y_m_d[1]),int('20'+y_m_d[0]))
    return jour

def reads_sejour(year):

    def sets_date_debut(row):
        if row['mois début'] != row['mois fin']:
            date_debut = row['Jour_deb'] +' '+ row['Date début'].split('-')[2] \
                         + ' ' +mois_dict[int(row['Date début'].split('-')[1])]
        else:
            date_debut = row['Jour_deb'] +' '+ row['Date début'].split('-')[2] 
        return date_debut

    def sets_date_debut_simple(row):
        if row['mois début'] != row['mois fin']:
            date_debut = row['Date début'].split('-')[2] \
                         + ' ' +mois_dict[int(row['Date début'].split('-')[1])]
        else:
            date_debut = row['Date début'].split('-')[2] 
        return date_debut
                         
        
    mois_dict = {1:"janvier",
                 2:"février",
                 3:"mars",
                 4:"avril",
                 5:"mai",
                 6:"juin",
                 7:"juillet",
                 8:"août",
                 9:"septembre",
                 10:"octobre",
                 11:"novembre",
                 12:"décembre"}
    
    sejour_df = pd.read_excel(PATH_DATA / "sejours.xlsx", 
                             sheet_name=str(year))
    
    sejour_df['nbr jours'] = pd.to_datetime(sejour_df['Date fin']+f'/{str(year)}',format="%d/%m/%Y") -\
                         pd.to_datetime(sejour_df['Date début']+f'/{str(year)}',format="%d/%m/%Y")

    sejour_df['nbr jours'] = sejour_df['nbr jours'].apply(lambda row: row.days+1)
    
    sejour_df['Nom du sejour'] = sejour_df['Nom du sejour'].str.replace(' (sub)', '')
    sejour_df['Nom du sejour'] = sejour_df['Nom du sejour'].str.replace('Séjour ', '',n=1)

    sejour_df['mois début'] = sejour_df.apply(lambda row: row['Date début'].split('/')[1],axis=1)
    sejour_df['mois fin'] = sejour_df.apply(lambda row: row['Date fin'].split('/')[1],axis=1)
    sejour_df['Date début old'] = sejour_df['Date début']

    prefixe =str(year)[2:4]+'-'
    sejour_df['Date début'] = sejour_df.apply(lambda row: prefixe+row['Date début'].split('/')[1]+'-'+
                                              row['Date début'].split('/')[0],axis=1)
    
    sejour_df['Date début simple'] = sejour_df.apply(sets_date_debut_simple,axis=1)
    sejour_df['Date fin'] = sejour_df.apply(lambda row: prefixe+row['Date fin'].split('/')[1]+'-'+
                                              row['Date fin'].split('/')[0],axis=1)
    sejour_df['Date fin simple'] =  sejour_df.apply(lambda row: row['Date fin'].split('-')[2],axis=1) \
                             + ' ' +sejour_df.apply(lambda row: mois_dict[int(row['Date fin'].split('-')[1])],axis=1)
    
    sejour_df['Jour_deb'] = sejour_df.apply(computes_jour_deb,axis=1)
    sejour_df['Jour_fin'] = sejour_df.apply(computes_jour_fin,axis=1)
    sejour_df['Date début'] = sejour_df.apply(sets_date_debut,axis=1)
    
    
    sejour_df['Date fin'] = sejour_df['Jour_fin'] +' '+ sejour_df.apply(lambda row: row['Date fin'].split('-')[2],axis=1) \
                             + ' ' +sejour_df.apply(lambda row: mois_dict[int(row['Date fin'].split('-')[1])],axis=1)

    

    sejour_df['date'] = sejour_df['Date début'] + '-' + sejour_df['Date fin']
    sejour_df['date simple'] = sejour_df['Date début simple'] + '-' + sejour_df['Date fin simple']

    sejour_df = sejour_df.rename(columns={'Nom du sejour': 'titre',})
    
    sejour_df = sejour_df[['titre',
                           'date',
                           'date simple',
                           'Description',
                           'Hébergement',
                           'Type de séjour',
                           'Coût',
                           'Responsable_1',
                           'Responsable_2',
                           'Responsable_3',
                           'Date début old',
                           'nbr jours',
                           'URL']]
    #sejour_df['Description'] = sejour_df['Description'].apply(crlf)
    sejour_df = sejour_df.fillna(txt_remplacement)
    return sejour_df

def crlf(x):
    if isinstance(x,str):
        return x.replace('\n', '\r\n')
    return x
    
def makes_calendrier_developpe(we_flag, jeudi_flag, sejour_flag, vtt_flag, year, save_flag):

    list_df = []
    if jeudi_flag:
        jeudi_df = pd.read_excel(PATH_DATA / "sorties_jeudi.xlsx",
                                 sheet_name=str(year))
        list_df.append(jeudi_df)
        
    if we_flag:
        we_df = pd.read_excel(PATH_DATA / "sorties_we.xlsx",
                              sheet_name=str(year))
        list_df.append(we_df)
        
    if vtt_flag :
        vtt_df = pd.read_excel(PATH_DATA / "sorties_VTT.xlsx",
                               sheet_name=str(year))
        list_df.append(vtt_df)

    if sejour_flag :
        sejour_df = pd.read_excel(PATH_DATA / "sejours.xlsx",
                                  sheet_name=str(year))
        prefixe =str(year)[2:4]+'-'
        sejour_df['nbr jours'] = pd.to_datetime(sejour_df['Date fin']+f'/{str(year)}',format="%d/%m/%Y") -\
                         pd.to_datetime(sejour_df['Date début']+f'/{str(year)}',format="%d/%m/%Y")
        sejour_df['nbr jours'] = sejour_df['nbr jours'].apply(lambda row: row.days+1)
        sejour_df['nbr jours'] = 'Nombre de jours : '+sejour_df['nbr jours'].astype(str)+'\nType de séjour : '+ sejour_df['Type de séjour']

        sejour_df['Date début'] = sejour_df.apply(lambda row: prefixe+row['Date début'].split('/')[1]+'-'+
                                               row['Date début'].split('/')[0],axis=1)
        sejour_df['Date fin'] = sejour_df.apply(lambda row: prefixe+row['Date fin'].split('/')[1]+'-'+
                                               row['Date fin'].split('/')[0],axis=1)
        sejour_df['Semaine'] = sejour_df.apply(computes_num_week,axis=1)
        sejour_df['Jour'] = sejour_df.apply(computes_jour,axis=1)
        
    
        sejour_df = sejour_df.rename(columns={"Nom du sejour": "Nom",
                                              "Type de séjour": "Départ",
                                              "Date début":"Date",
                                              "nbr jours":"GP",
                                              "Hébergement":"MP",
                                              "Date fin":"PP"})
        sejour_df['MP'] = ''
        sejour_df['PP'] = ''
        sejour_df['Départ'] = ''
        sejour_df = sejour_df [['Semaine','Jour','Date','Nom','GP','MP','PP','Départ']]
        sejour_df = sejour_df[sejour_df['Nom'].notna()]
        list_df.append(sejour_df)
        
    result = pd.concat(list_df, axis=0)
    result['Jour'] = result['Date'].apply(gets_day_of_the_week2)
    result = result.sort_values(['Date'])
    
    result = result.fillna('')
    if save_flag:
        file = PATH_DATA / Path(str(year)) / Path(f"calendrier_{str(year)}_developpe.xlsx")
        result.to_excel(file,index=False)
        file = PATH_DATA / Path(str(year)) / Path("randonnées_subventionnees_"+str(year)+".xlsx")
        result[result['Nom'].str.contains('(sub)')][['Jour','Date','Nom','Départ']].to_excel(file,index=False)
    
    return result
    
def builds_responsable_list(sejour_df, effectif_df, year):
    
    responsable_list = []
    for responsable1, responsable2, responsable3 in zip(sejour_df['Responsable_1'],
                                                        sejour_df['Responsable_2'],
                                                        sejour_df['Responsable_3']):
        
        internal_list = []                             
        nom, prenom = get_nom_prenom(responsable1)
        num_portable = get_tel(effectif_df, nom, prenom,year)
        email = get_courriel(effectif_df, nom, prenom)
        internal_list.append(dict(nom=responsable1,
                                  tel=num_portable,
                                  mail=email))
        if responsable2 != txt_remplacement:
            nom, prenom = get_nom_prenom(responsable2)
            num_portable = get_tel(effectif_df, nom, prenom,year)
            email = get_courriel(effectif_df, nom, prenom)
            internal_list.append(dict(nom=responsable2,
                                      tel=num_portable,
                                      mail=email))
        if responsable3 != txt_remplacement:
            nom, prenom = get_nom_prenom(responsable3)
            num_portable = get_tel(effectif_df, nom, prenom,year)
            email = get_courriel(effectif_df, nom, prenom)
            internal_list.append(dict(nom=responsable3,
                                      tel=num_portable,
                                      mail=email))
        responsable_list.append(internal_list)
    return responsable_list

def builds_list_sejour_qrc(sejour_df,doc,i_dep,i_fin):
    num_img = 1
    qrc_sejour = []
    qrc_flag =[]
    for url in sejour_df['URL'].tolist()[i_dep:i_fin]:
        if url != txt_remplacement:
            qrcode_img = qrcode.make(url)
            qrc_name = str(PATH_IMAGES / Path(f'qrc_sejour{num_img}.jpg'))
            qrcode_img.save(qrc_name, 'JPEG')
            qrc_sejour.append(InlineImage(doc,qrc_name,Cm(1.5)))
            qrc_flag.append(True)
            num_img += 1
        else:
            qrc_sejour.append('')
            qrc_flag.append(False)

    return qrc_sejour, qrc_flag
                
    

def builds_sejour_docx(list_docx,effectif_df,year):
    
    sejour_df = reads_sejour(year)
    responsable_list = builds_responsable_list(sejour_df,effectif_df,year)
    long = 2
    
    file_list = []
    idx = 0
    for i_dep in range(0,len(sejour_df),long):
        if i_dep+long <= len(sejour_df):
            i_fin = i_dep+long
            template_path_docx = PATH_TEMPLATES / "5-Template_sejour2.docx"
            doc = DocxTemplate(template_path_docx)
            qrc_sejour, qrc_flag = builds_list_sejour_qrc(sejour_df,doc,i_dep,i_fin)
            
            f ={
                'titre': sejour_df['titre'].tolist()[i_dep:i_fin],
                'date': sejour_df['date'].tolist()[i_dep:i_dep+long],
                'description': sejour_df['Description'].tolist()[i_dep:i_fin],
                'type_sejour': sejour_df['Type de séjour'].tolist()[i_dep:i_fin],
                'hebergement': sejour_df['Hébergement'].tolist()[i_dep:i_fin],
                'responsable': responsable_list[i_dep:i_fin],
                'cout': sejour_df['Coût'].tolist()[i_dep:i_fin],
                'qrc_code': qrc_sejour,
                'qrc_flag':qrc_flag}
            
        else:
            i_fin = len(sejour_df)
            template_path_docx = PATH_TEMPLATES / "6-Template_sejour1.docx"
            doc = DocxTemplate(template_path_docx)
            qrc_sejour, qrc_flag = builds_list_sejour_qrc(sejour_df,doc,i_dep,i_fin)
            
            f ={
                'titre': sejour_df['titre'].tolist()[i_dep:i_fin],
                'date': sejour_df['date'].tolist()[i_dep:i_fin],
                'description': sejour_df['Description'].tolist()[i_dep:i_fin],
                'type_sejour': sejour_df['Type de séjour'].tolist()[i_dep:i_fin],
                'hebergement': sejour_df['Hébergement'].tolist()[i_dep:i_fin],
                'responsable': responsable_list[i_dep:i_fin],
                'cout': sejour_df['Coût'].tolist()[i_dep:i_fin],
                'qrc_code': qrc_sejour,
                'qrc_flag':qrc_flag}
            
        doc.render(f)
        #    
        output_file = PATH_OUTPUT / Path("5-sejour"+str(idx)+".docx")
        doc.save(output_file)
        list_docx.append(output_file)
    
        idx += 1
        
def make_list_sejour(year,list_docx,qrc_dic,output_file):

    sejour_df = reads_sejour(year)
    template_path_docx = PATH_TEMPLATES / "4-Template_liste_sejours.docx"
    frameworks = []
    for row in sejour_df.iterrows():
            frameworks.append(dict(date=row[1]['date'],
                                   name=row[1]['titre'],
                                   type=row[1]['Type de séjour'],
                                   duree=row[1]['nbr jours']))
                       
    doc = DocxTemplate(template_path_docx)  
    context ={'year': year,
              'qrc_sejour': InlineImage(doc, qrc_dic['url_sejour'], Cm(3)),
              'nbr_sejours': len(sejour_df),
              'frameworks': frameworks}
    
    doc.render(context)
    output_file = PATH_OUTPUT / output_file
    doc.save(output_file)
    list_docx.append(output_file)

def combine_all_docx(list_docx,year):

    """
    Merges all the docx file contained in `list_docxt` with
    the mater docx document `filename_master
    """

    filename_master = list_docx[0]
    number_of_sections = len(list_docx)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(1, number_of_sections):
        doc_temp = Document_compose(list_docx[i])
        composer.append(doc_temp)
    file_name = f'CALENDRIER_{str(year)}.docx'
    combine_file = PATH_DATA / Path(str(year)) / file_name
    composer.save(combine_file)

def make_debut_calendrier(year,sorties_dic,sortie_vtt,sortie_route,effectif_df,qrc_dic,list_docx,file_output):

    mois_dict = {"janvier":"01",
                 "février":"02",
                 "mars":"03",
                 "avril":"04",
                 "mai":"05",
                 "juin":"06",
                 "juillet":"07",
                 "août":"08",
                 "septembre":"09",
                 "octobre":"10",
                 "novembre":"11",
                 "décembre":"12"}

    sejour_df = reads_sejour(year)
    sejour_df['Date début old'] = sejour_df.apply(lambda row: row['Date début old'].split('/')[1] + 
                                            "-" + row['Date début old'].split('/')[0],axis=1)
    org_dic = defaultdict(list)
    cols_name = list(sejour_df.columns)
    cols_name .remove('titre')
    cols_name.remove('Date début old')
    cols_name.remove('date simple')
    for key, val in sorties_dic.items():
        for col_name in cols_name:
            org_dic[col_name].append(" ")
        org_dic['titre'].append(key)
        d = mois_dict[val.split()[1]]+"-"+str(val.split()[0]).zfill(2)
        org_dic['Date début old'].append(d)
        org_dic['date simple'].append(val)
    org_df = pd.DataFrame.from_dict(org_dic)
    sejour_df = pd.concat([sejour_df, org_df], axis=0)   
    sejour_df = sejour_df.sort_values('Date début old')
    
    template_path_docx = PATH_TEMPLATES / "3-Template_debut.docx"
    frameworks = []
    for idx,row in enumerate(sejour_df.iterrows()):
        frameworks.append(dict(date=row[1]['date simple'],
                                   name=row[1]['titre'],
                                   id=idx+1))
                
            
    doc = DocxTemplate(template_path_docx) 

    context ={'year': year,
              'org1':org[0],
              'org2':org[1],
              'route': [(f"{x[0][0]}. {x[1]}" ,get_tel (effectif_df, x[1], x[0], year)) for x in sortie_route],
              'vtt': [(f"{x[0][0]}. {x[1]}" ,get_tel (effectif_df, x[1], x[0], year)) for x in sortie_vtt],
              'frameworks': frameworks,
              'qrc_asso': InlineImage(doc, qrc_dic['url_assoconnect'], Cm(3)),
              'qrc_ffct': InlineImage(doc, qrc_dic['url_securite'], Cm(2)),
              'qrc_securite_ffct': InlineImage(doc, qrc_dic['url_securite_ffct'], Cm(2)),
              'qrc_rond_point': InlineImage(doc, qrc_dic['url_rond_point'], Cm(2)),
              'qrc_sejour': InlineImage(doc, qrc_dic['url_sejour'], Cm(2)),}
   
    doc.render(context)
    output_file = PATH_OUTPUT / file_output
    doc.save(output_file)
    list_docx.append(output_file)

def gets_circuit_makers_dict(effectif_df, year):
    circuit_makers_path = PATH_TEMPLATES / "traceurs_circuits.xlsx"
    circuit_makers_df = pd.read_excel(circuit_makers_path, sheet_name=str(year))
    circuit_makers_list = []
    for row in circuit_makers_df.iterrows():
        nom, prenom =get_nom_prenom(row[1]['Nom'])
        tel = get_tel (effectif_df, nom, prenom,year)
        courriel = get_courriel (effectif_df, nom, prenom)
        circuit_makers_list.append(dict(nom=row[1]['Nom'],
                               tel=tel,
                               fonction=row[1]['fonction'],
                               mail=courriel,
                               initiales=f'({prenom[0]}{nom[0]})'))
    return circuit_makers_list

def make_middle_calendrier(year, list_docx, file_output, effectif_df):

    def chooses_color(row):
        if row['Nom'].startswith("Rando"):
            return color_dic['rando']
        elif row['Nom'].startswith("Séjour"):
            return color_dic['sejour']
        elif  '-VTT' in row['Nom']:
            return color_dic['vtt']
        else:
            return color_dic['club']

    template_path_docx = PATH_TEMPLATES / "7-Template_middle.docx"

    we_flag = True
    jeudi_flag = True
    sejour_flag = True
    vtt_flag = True
    save_flag = True
    
    result = makes_calendrier_developpe(we_flag, jeudi_flag, sejour_flag, vtt_flag, year, save_flag)
    circuit_makers_list = gets_circuit_makers_dict(effectif_df, year)

    mois_selections = [['02','03'],['04'],['05'],['06'],['07'],['08','09'],['10','11']]
    mois_dict = {1:"janvier",
                     2:"février",
                     3:"mars",
                     4:"avril",
                     5:"mai",
                     6:"juin",
                     7:"juillet",
                     8:"août",
                     9:"septembre",
                     10:"octobre",
                     11:"novembre",
                     12:"décembre"}
    
    result['mois'] = result.apply(lambda row: row['Date'].split('-')[1],axis=1)
    
    result['Date'] = result.apply(lambda row: '/'.join(row['Date'].split('-')[::-1][0:2]),axis=1)
    
    #result['Jour'] = result.apply(lambda row: row['Jour'].capitalize(),axis=1)
    
    result['Jour'] = result['Date'].apply(gets_day_of_the_week)
    result['Type'] = result.apply(lambda row: 1 if row['Nom'].startswith("Rando") else 0 ,axis=1)
    
    result['color'] = result.apply(chooses_color ,axis=1)
    result['Nom'] = result['Nom'].str.replace(' (sub)', '')
    result['Nom'] = result['Nom'].apply(lambda x: x[7:] if x.startswith('Séjour')  else x)
    result = result[result['Nom'] !=""]
    activity_dic = {}
    for mois_selection in mois_selections:
        k = ' '.join([mois_dict[int(x)] for x in mois_selection]).capitalize()
        dg = result.query("mois in @mois_selection")
        activity_dic[k] = dg.to_dict('records')
    
    context = {'year': year,
               'titre': activity_dic.keys(),
               'activity': activity_dic,
               'circuit': circuit_makers_list} 
    
    doc = DocxTemplate(template_path_docx) 
    doc.render(context)
    doc.save(file_output)
    list_docx.append(file_output)

def make_fin_calendrier(year,date_ag,list_docx,effectif_df,qrc_dic, file_output):

    template_path_docx = PATH_TEMPLATES / "10-Template_fin.docx"
    cd_path = PATH_TEMPLATES / "CD.xlsx"
    cd_df = pd.read_excel(cd_path, sheet_name=f'CD_{year}')
    
    frameworks = []
    for row in cd_df.iterrows():
        nom, prenom =get_nom_prenom(row[1]['Nom'])
        tel = get_tel (effectif_df, nom, prenom, year)
        courriel = get_courriel (effectif_df, nom, prenom)
        frameworks.append(dict(nom=row[1]['Nom'],
                               tel=tel,
                               fonction=row[1]['fonction'],
                               mail=courriel,))

    doc = DocxTemplate(template_path_docx) 
    context = {'year': year,
               'date_ag':date_ag,
               'qrc': InlineImage(doc, qrc_dic['url_assoconnect'], Cm(3)),
               'frameworks': frameworks}
       
    doc.render(context)
    output_file = PATH_OUTPUT / file_output
    doc.save(output_file)
    list_docx.append(output_file)

def makes_liste_sorties_cool(list_docx, year, file_output):

    template_path_docx = PATH_TEMPLATES / "8-Template_randos_cool.docx"
    we_flag = True
    jeudi_flag = True
    sejour_flag = False
    vtt_flag = False
    save_flag = False
    
    result = makes_calendrier_developpe(we_flag, jeudi_flag, sejour_flag, vtt_flag, year, save_flag)
    sieve_list = []
    dist_max = 60
    dev_max = 800
    for row in result.iterrows():
        row = row[1]
        for col in ['GP','MP','PP']:
            if '/' in row[col]:
                try:
                    dist = int(row[col].split('/')[0].split('km')[0])
                    dev = int(row[col].split('/')[1].split('m')[0])
                    if dist<dist_max and dev < dev_max:
                        row_c = row.copy()
                        row_c['cool'] = row[col]
                        sieve_list.append(row_c)
                except:
                    pass
    cool_df = pd.DataFrame(sieve_list)[['Jour','Date','Nom','cool','Départ']]
    cool_df['Date'] = cool_df.apply(lambda row: '/'.join(row['Date'].split('-')[::-1][0:2]),axis=1)
    #cool_df.to_excel(PATH_DATA / 'cool.xlsx',index=False)
    frameworks = []
    for row in cool_df.iterrows():
            frameworks.append(dict(Jour=row[1]['Jour'].capitalize(),
                                   Date=row[1]['Date'],
                                   Nom=row[1]['Nom'],
                                   PPP=row[1]['cool'],
                                   Départ=row[1]['Départ'],
                              ))
        
    context ={'year': year,
              'nbr_rando': len(cool_df),
              'frameworks': frameworks}
    
    doc = DocxTemplate(template_path_docx) 
    doc.render(context)
    output_file = PATH_OUTPUT / file_output
    doc.save(output_file)
    list_docx.append(output_file)
    return


def make_liste_randonnées(list_docx,year, file_output):
    
    template_path_docx = PATH_TEMPLATES / "9-Template_liste_randonnées.docx"
    file = PATH_DATA / Path(str(year)) / Path("randonnées_subventionnees_"+str(year)+".xlsx")
    mois_dict = {1:"janvier",
                 2:"février",
                 3:"mars",
                 4:"avril",
                 5:"mai",
                 6:"juin",
                 7:"juillet",
                 8:"août",
                 9:"septembre",
                 10:"octobre",
                 11:"novembre",
                 12:"décembre"}
    rando_df = pd.read_excel(file)
    rando_df['Nom'] = rando_df['Nom'].str.replace(' (sub)', '')
    rando_df['day'] = rando_df['Date'].apply(lambda row: row.split('-')[2])
    rando_df['mois'] = rando_df['Date'].apply(lambda row: mois_dict[int(row.split('-')[1])])
    
    rando_df['Date'] = rando_df['Jour'].str.capitalize()+' '+rando_df['day']+' '+rando_df['mois']
    frameworks = []
    for idx,row in enumerate(rando_df.iterrows()):
            frameworks.append(dict(date=row[1]['Date'],
                                   name=row[1]['Nom'],
                                   lieu=row[1]['Départ'],
                                   id=idx+1))
        
    context ={'year': year,
              'nbr_rando': len(rando_df),
              'frameworks': frameworks}
    
    doc = DocxTemplate(template_path_docx) 
    doc.render(context)
    doc.save(file_output)
    list_docx.append(file_output)

def make_liste_vtt(list_docx,year, file_output):
    
    template_path_docx = PATH_TEMPLATES / "9b-Template_liste_VTT.docx"
    file = PATH_DATA / Path(str(year)) / Path("calendrier_"+str(year)+"_developpe.xlsx")
    mois_dict = {1:"janvier",
                 2:"février",
                 3:"mars",
                 4:"avril",
                 5:"mai",
                 6:"juin",
                 7:"juillet",
                 8:"août",
                 9:"septembre",
                 10:"octobre",
                 11:"novembre",
                 12:"décembre"}
    result = pd.read_excel(file)
    result['Nom'] = result['Nom'].str.replace(' (sub)', '')
    result['day'] = result['Date'].apply(lambda row: row.split('-')[2])
    result['mois'] = result['Date'].apply(lambda row: mois_dict[int(row.split('-')[1])])

    
    result = result.fillna('  ')
    vtt_df = result[result['GP'].str.contains("VTT")]
    vtt_df['Nom'] = vtt_df['Nom'].apply(lambda x : x[0:len(x)-4] if x.endswith('-VTT') else x)
    
    vtt_df['Date'] = vtt_df['Jour'].str.capitalize()+' '+vtt_df['day']+' '+vtt_df['mois']
    frameworks = []
    for idx,row in enumerate(vtt_df.iterrows()):
            frameworks.append(dict(date=row[1]['Date'],
                                   name=row[1]['Nom'],
                                   lieu=row[1]['Départ'],
                                   id=idx+1))
        
    context ={'year': year,
              'nbr_rando': len(vtt_df),
              'frameworks': frameworks}
    
    doc = DocxTemplate(template_path_docx) 
    doc.render(context)
    doc.save(file_output)
    list_docx.append(file_output)

def make_calendar(year):

    flag_make_debut = True
    flag_make_list_sejour = False
    flag_make_middle_calendrier = True
    flag_makes_liste_sorties_cool = False
    flag_make_liste_randonnées = True
    flag_make_liste_vtt = True
    flag_builds_sejour_docx = True
    flag_make_fin_calendrier = True

    
    effectif_df = read_effectif(path_effectif, year)
        
    qrc_dic = make_qrcode(PATH_IMAGES, url_dic)
    
    list_docx = []
    list_docx.append(PATH_TEMPLATES / "1-Template_couverture couleur extérieure première.docx")
    list_docx.append(PATH_TEMPLATES / "2-Template_couverture couleur intérieure première.docx")
    
    file_output = "3-debut_calendrier.docx"
    if flag_make_debut:
        make_debut_calendrier(year,
                              sorties_dic,
                              sortie_vtt,
                              sortie_route,
                              effectif_df,
                              qrc_dic,
                              list_docx,
                              file_output)
    else:
        list_docx.append(file_output)
    
    file_output = "4-liste_sejours.docx"
    if flag_make_list_sejour:
        make_list_sejour(year,
                         list_docx,
                         qrc_dic,
                         file_output)
    else:
        #list_docx.append(file_output)
        pass
        
    
    if flag_builds_sejour_docx:
        builds_sejour_docx(list_docx,
                           effectif_df,
                           year)
    else:
        pass
         
    file_output = PATH_OUTPUT / "6-middle.docx"
    if flag_make_middle_calendrier:
        make_middle_calendrier(year,list_docx,file_output,effectif_df)
    else:
        list_docx.append(file_output)

    
    #file_output = "7-liste_randonnées_cool.docx"
    #if flag_makes_liste_sorties_cool:
        #makes_liste_sorties_cool(list_docx, year, file_output)
    #else:
        #list_docx.append(file_output)séjour
    
    file_output = PATH_OUTPUT / "8-liste_randonnées.docx"
    if flag_make_liste_randonnées:
        make_liste_randonnées(list_docx,year, file_output)
    else:
        list_docx.append(file_output)

    file_output = PATH_OUTPUT / "8a-liste_randonnées.docx"
    if flag_make_liste_vtt:
        make_liste_vtt(list_docx,year, file_output)
    else:
        list_docx.append(file_output)

    file_output = "9-fin.docx"
    if flag_make_fin_calendrier:
        make_fin_calendrier(year,
                            date_ag,
                            list_docx,
                            effectif_df,
                            qrc_dic,
                            file_output)
    else:
        list_docx.append(file_output)
        
    list_docx.append(PATH_TEMPLATES / "11-Template_couverture couleur intérieure dernière.docx")
    list_docx.append(PATH_TEMPLATES / "12-Template_couverture couleur extérieure dernière.docx")
    combine_all_docx(list_docx,year)

    file_name = f'CALENDRIER_{str(year)}.docx'
    combine_file = PATH_DATA  / Path(str(year)) / file_name
    print(f"votre fichier est prêt sous le nom : {combine_file}")
    

(PATH_DATA, PATH_IMAGES, PATH_OUTPUT, sorties_dic,
                sortie_vtt, sortie_route, url_dic, date_ag,color_dic,org,txt_remplacement) = reads_json()
make_calendar(year)
print( "Le calendrier a été créé")

    