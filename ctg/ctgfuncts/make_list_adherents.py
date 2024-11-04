_all_ = ["make_list_adherents"]

from pathlib import Path
from datetime import datetime

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

def make_list_adherents(ctg_path):

    current_year = datetime.now().year
    effectif_file = ctg_path / Path(str(current_year)) / 'DATA' /Path(str(current_year)+'.xlsx')
    df = pd.read_excel(effectif_file,usecols= ['N° Licencié','Nom','Prénom','Date de naissance','Sexe'])
    df = df.sort_values(by=['Nom','Prénom'])
    df.rename(columns={'N° Licencié': 'Licence',
                       'Nom': 'Nom',
                       'Prénom':'Prénom',
                       'Date de naissance':'D de N',
                       'Sexe':'S',}, inplace=True)
    
    result_path = Path(ctg_path) / str(current_year) / 'DATA'
    output_file = result_path / f'liste_adherents_CTG_{current_year}.docx'
    template_path_docx = Path(r"c:\users\franc\Temp")
    
    frameworks = []
    for idx,row in df.iterrows():
        frameworks.append(dict(id=row['Licence'],
                                  surname=row['Nom'],
                                  name=row['Prénom'],
                                  ddn=row['D de N'],
                                  s=row['S'] ))
            
    
    context ={'year': current_year,
              'frameworks': frameworks}
    
    template_docx = template_path_docx / 'template_Liste_CTG.docx'
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    convert(output_file)
    print(output_file)
