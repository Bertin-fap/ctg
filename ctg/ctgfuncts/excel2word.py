__all_  = ['create_word_from_template',
           'combine_all_docx',
           'make_sejour_docx',
           'make_list_sejour',
           'make_calendrier']

from pathlib import Path
import functools
import os
import unicodedata
import zipfile
import datetime
import shutil

import jinja2
import pandas as pd
from docx.shared import Cm, Emu, Inches, Mm
from docxtpl import DocxTemplate, InlineImage
from docxcompose.composer import Composer
from docx import Document as Document_compose

from openpyxl import load_workbook

from ctg.ctgfuncts.ctg_classes import   EffectifCtg

nfc = functools.partial(unicodedata.normalize,'NFD')
convert_to_ascii = lambda text : nfc(text). \
                                     encode('ascii', 'ignore'). \
                                     decode('utf-8').\
                                     strip()

def get_placeholder_dict(file):

    """
    Gets, from the Excel file, the placeholder needed in the Word template 
    """
    
    workbook = load_workbook(filename=file)
    workbook.sheetnames
    sheet = workbook.active
    
    place_holder_dict = {}
    for value in sheet.iter_rows(min_row=1,
                                 max_row=len( list(sheet.rows)),
                                 values_only=True):
        if value[1] is not None:
            place_holder_dict[value[0]] = value[1]
        else:
            place_holder_dict[value[0]] = '  '
    
    return place_holder_dict

def get_images_from_excel(full_path,path_images):

    """
    Extacts all the images from an Excel file
    """
    path = Path(full_path).parent
    EmbeddedFiles = zipfile.ZipFile(full_path).namelist()
    image_files = [F for F in EmbeddedFiles if F.count('.jpg') 
                                           or F.count('.jpeg') 
                                           or F.count('.png') ]
    
    with zipfile.ZipFile(full_path, 'r') as zip_ref:
        zip_ref.extractall(path_images)

    images_path_list = []
    for image in image_files:
        image_name = os.path.basename(image)
        src = Path(path_images) / Path(image)
        dst = Path(path_images) / Path(image_name)
        images_path_list.append(dst)
        shutil.copyfile(src, dst)
        
    shutil.rmtree(Path(path_images) / Path('xl'))
    shutil.rmtree(Path(path_images) / Path('_rels'))
    shutil.rmtree(Path(path_images) / Path('docProps'))
    os.remove(Path(path_images) / Path('[Content_Types].xml'))
    return images_path_list


        
def get_nom_prenom(place_holder_dict):
    
    """
     Generates the list of nanes and the list of surnames of the sejour responsibles
    """

    # extract the full name of the sejour organizers
    prenom_nom = [convert_to_ascii(v).upper() for k,v in place_holder_dict.items()
                  if 'responsable_' in k ]
    prenom_nom = [x for x in prenom_nom if len(x)>0] # reject empty fieds
    nom = [x.split()[1] for x in prenom_nom if len(prenom_nom)>0]
    prenom = [x.split()[0] for x in prenom_nom if len(prenom_nom)>0]
    
    return nom, prenom

def get_tel (df, nom, prenom):

    """
    Gets the phone number of prenom nom from the CTG gilda list
    """
    
    dg = df.query('Nom==@nom and Prénom==@prenom')
    
    num_portable = dg["N° Portable"].tolist()[0]
    if num_portable==0:
        num_portable = str(int(dg["N° Tél"].tolist()[0])).zfill(10)

    return num_portable

def get_courriel(df, nom, prenom):

    """
    Gets the email of prenom nom from the CTG gilda list
    """

    dg = df.query('Nom==@nom and Prénom==@prenom')
    email = dg["E-mail"].tolist()[0]

    return email

def add_acompte(place_holder_dict):
    
    
    acompte_montant = [v for k,v in place_holder_dict.items() if 'acompte_' in k ]
    date_encaissement = [v for k,v in place_holder_dict.items() if 'date_encaissement_' in k ]
    frameworks = []
    idx = 1
    for n, p in zip(acompte_montant,date_encaissement):
        if len(n.strip())>0:
            frameworks.append(dict(name=f'Acompte_{idx}',
                                   acompte_montant=n,
                                   date_encaissement=p))
            idx += 1

    context ={'acompte': frameworks}
    place_holder_dict = place_holder_dict | context
    
    return place_holder_dict  
    
def read_effectif(ctg_path, year):
    
    # reads the CTG effectit Exel file of the current year
    effectif_path = ctg_path / 'SORTIES' 
    
    effectif = EffectifCtg(year,effectif_path)
    df = effectif.effectif

    # deals with empty celles
    df["N° Tél"] = df["N° Tél"].fillna(0)
    df["N° Portable"] = df["N° Portable"].fillna(0)
    
    return df
    
def add_tel_mail(place_holder_dict,ctg_path,year):

    """
    creates the placeholder phone number courriel from the CGT list of the
    current year
    """

    
    df = read_effectif(ctg_path, year)
    
    nom, prenom = get_nom_prenom(place_holder_dict)
    
    frameworks = []
    for n, p in zip(nom,prenom):
        nom_prenom_abbr = f'{n.capitalize()} {p[0]}.'
        num_portable = get_tel (df, n, p)
        email = get_courriel(df, n, p)
        
        frameworks.append(dict(name=nom_prenom_abbr,
                               tel=num_portable,
                               email=email))
    context ={'frameworks': frameworks}

    place_holder_dict = place_holder_dict | context

    return place_holder_dict
    

def create_word_from_template(template_path,
                              filename,
                              current_dir,
                              output_path,
                              ctg_path,
                              year):

    template_docx = template_path / 'Template_sejour.docx'
    doc = DocxTemplate(template_docx) 
    images_path_list = get_images_from_excel(filename,current_dir)
    img_placeholder1_path = images_path_list[0]
    img_placeholder2_path = images_path_list[1]
    placeholder_1 = InlineImage(doc, str(img_placeholder1_path), Cm(8))
    placeholder_2 = InlineImage(doc, str(img_placeholder2_path), Cm(8))
    
    # builds the placeholder dict
    place_holder_dict = get_placeholder_dict(filename)
    place_holder_dict["placeholder_1"]= placeholder_1
    place_holder_dict["placeholder_2"]= placeholder_2
    
    # adds name tel emil 
    place_holder_dict = add_tel_mail(place_holder_dict,ctg_path,year)
    place_holder_dict = add_acompte(place_holder_dict)

    doc.render(place_holder_dict)
    doc.save(output_path)
    os.remove(images_path_list[0])
    os.remove(images_path_list[1])
    
def combine_all_docx(ctg_path,files_list,year):

    """
    Merges all the docx file contained in `files_list` with
    the mater docx document `filename_master
    """

    result_path = ctg_path / 'CALENDRIER' / str(year)
    filename_master = result_path / f'Calendrier {str(year)}' / 'debut.docx'
    number_of_sections=len(files_list)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(files_list[i])
        composer.append(doc_temp)
    file_name = f'CALENDRIER_{str(year)}.docx'
    combine_file = result_path / file_name
    composer.save(combine_file)
    
def make_list_sejour(ctg_path,
                     year,
                     list_docx):

    result_path = ctg_path  / 'CALENDRIER' / str(year)
    output_file = result_path / 'sejours_list.docx'
    template_path_docx = ctg_path / 'CALENDRIER' / 'Template_sejours'
    descp_sejour_xlsx_path = result_path / 'Description_sejour_xlsx'
    frameworks = []
    for sejour_xlsx in descp_sejour_xlsx_path.iterdir():
        if not str(sejour_xlsx).startswith("~$"):
            place_holder_dict = get_placeholder_dict(sejour_xlsx)
            frameworks.append(dict(date=place_holder_dict['date'],
                                   name=place_holder_dict['titre'],
                                   type=place_holder_dict['type_sejour']))
            
        else:
            print ('filter',file_sejour)
    
    context ={'year': year,
             'frameworks': frameworks}

    template_docx = template_path_docx / 'Template_liste_sejours.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    list_docx.append(output_file)

def make_list_randonnée(ctg_path,
                        year,
                        list_docx):

    result_path = ctg_path  / 'CALENDRIER' / str(year)
    output_file = result_path / 'randonnees.docx'
    template_path_docx = ctg_path / 'CALENDRIER' / 'Template_sejours'
    descp_randonnees_xlsx_path = result_path / 'randonnées.xlsx'
    frameworks = []
    df = pd.read_excel(descp_randonnees_xlsx_path)
    for idx,row in df.iterrows():
        frameworks.append(dict(date=row['date'],
                               name=row['nom'],))
            
    
    context ={'year': year,
             'frameworks': frameworks}
    
    template_docx = template_path_docx / 'Template_liste_randonnées.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    list_docx.append(output_file)

def make_cd(ctg_path,
            year,
            list_docx):

    result_path = ctg_path  / 'CALENDRIER' / str(year)
    output_file = result_path / 'cd.docx'
    template_path_docx = ctg_path / 'CALENDRIER' / 'Template_sejours'
    descp_cd_xlsx_path = result_path / 'CD.xlsx'

    effectif = read_effectif(ctg_path, year)   

    frameworks = []
    df = pd.read_excel(descp_cd_xlsx_path)
    for idx,row in df.iterrows():
        nom = convert_to_ascii(row['Nom']).upper()
        prenom = convert_to_ascii(row['Prénom']).upper()
        frameworks.append(dict(fonction=row['fonction'],
                               nom=f"{row['Nom']}  {row['Prénom']}",
                               tel = get_tel (effectif, nom, prenom),
                               courriel = get_courriel(effectif, nom, prenom)))
            
    
    context ={'year': year,
             'frameworks': frameworks}
    
    template_docx = template_path_docx / 'Template_CD.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    list_docx.append(output_file)
    
    
def make_sejour_docx(ctg_path,
                     ctg_list_membres,
                     year,
                     list_docx):

    result_path = ctg_path  / 'CALENDRIER' / str(year)
    template_path_docx = ctg_path / 'CALENDRIER' / 'Template_sejours'
    descp_sejour_xlsx_path = result_path / 'Description_sejour_xlsx'
    
    for sejour_xlsx in descp_sejour_xlsx_path.iterdir():
        docx_name = Path(sejour_xlsx).stem+'.docx'
        result_file_path = result_path / docx_name
        file_sejour = sejour_xlsx.name
        if not file_sejour.startswith("~$"):
            list_docx.append(result_file_path)
            create_word_from_template(template_path_docx,
                                      sejour_xlsx,
                                      ctg_path,
                                      result_file_path,
                                      ctg_list_membres,
                                      year)
        else:
            print ('filter',file_sejour)

def make_calendrier(year):
    ctg_path = Path.home() / 'CTG'
    ctg_list_membres = ctg_path
    
    result_path = ctg_path / 'CALENDRIER' / str(year)
    
    list_docx = []
    list_docx_to_keep = []
    
    make_list_sejour(ctg_path,
                     year,
                     list_docx)
    make_sejour_docx(ctg_path,
                     ctg_list_membres,
                     year,
                     list_docx)
    file = result_path / f'Calendrier {str(year)}'/'middle.docx'
    list_docx.append(file)
    list_docx_to_keep.append(file)
    make_list_randonnée(ctg_path,
                        year,
                        list_docx)
    make_cd(ctg_path,
            year,
            list_docx)
    
    combine_all_docx(ctg_path,list_docx,year)
    
    for file in set(list_docx)-set(list_docx_to_keep):
        os.remove(file)


    