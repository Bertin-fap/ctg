__all_  = ['create_word_from_template']

from pathlib import Path
import functools
import os
import unicodedata
import zipfile
import datetime
import shutil

import jinja2
from docx.shared import Cm, Emu, Inches, Mm
from docxtpl import DocxTemplate, InlineImage
from openpyxl import load_workbook

import ctg.ctggui as ctgg
import ctg.ctgfuncts as ctg

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
    

def add_tel_mail(place_holder_dict):

    """
    creates the placeholder phone number courriel from the CGT list of the
    current year
    """

    # reads the CTG effectit Exel file of the current year
    today = datetime.date.today()
    year = today.year
    effectif = ctg.EffectifCtg(year,ctg_path)
    df = effectif.effectif

    # deals with empty celles
    df["N° Tél"] = df["N° Tél"].fillna(0)
    df["N° Portable"] = df["N° Portable"].fillna(0)

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
                              ctg_path):

    doc = DocxTemplate(template_path) 
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
    place_holder_dict = add_tel_mail(place_holder_dict)
    
    doc.render(place_holder_dict)
    doc.save(output_path)
    os.remove(images_path_list[0])
    os.remove(images_path_list[1])
