import ctg.ctggui as ctgg
import ctg.ctgfuncts as ctg

import os
import random
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from docxcompose.composer import Composer
from docx import Document as Document_compose
from docx2pdf import convert
from PyPDF4 import PdfFileMerger

def make_sample(n_sample):

    
    URL = r"I:\Mon Drive\DRIVE DU CTG\3-PROJETS DU CTG\2-BRA 2025-PDU\5-HEBERGEMENT - JAR\hebergement.xlsx"
    hebergement_df = pd.read_excel(URL)
    hebergement = list(zip(hebergement_df['Ville'],hebergement_df['Hôtel']))
    ctg_path = Path.home() / r"CTG\SORTIES"
    effectif = ctg.EffectifCtg(2024,ctg_path)
    sample = effectif.effectif.sample(n=n_sample)[['Nom','Prénom','N° Portable']]
    sample['ville-hotel'] = [random.choice(hebergement) for _ in range(n_sample)]
    sample[['ville', 'hotel']] = pd.DataFrame(sample['ville-hotel'].tolist(), index=sample.index)
    sample['Nom-Prenom'] = sample.apply(lambda row: f"{row['Nom']} {row['Prénom']}",axis=1)
    sample = sample.drop(['ville-hotel','Nom','Prénom'],axis=1)
    sample["Dossard"] = [random.randint(1,800) for _ in range(n_sample)]
    sample = sample.fillna('inconnu')
    sample = sample.sort_values(by=['Dossard'], ascending=True)
    return sample

def make_etiquettes(sample):
    
    template_path_docx = Path.home() / r"CTG\BRA-BRO-BG\BRA\DATA\transport\template_EtiquetteBagage_BRA.docx"
    long = 10
    
    year = 2025
    file_list = []
    idx = 0
    for i_dep in range(0,len(sample),long):
        if i_dep+long < len(sample):
            f ={'year': year,
                'dossard': sample['Dossard'].tolist()[i_dep:i_dep+long],
                'name': sample['Nom-Prenom'].tolist()[i_dep:i_dep+long],
                'tel': sample['N° Portable'].tolist()[i_dep:i_dep+long],
                'ville': sample['ville'].tolist()[i_dep:i_dep+long],
                'lieu': sample['hotel'].tolist()[i_dep:i_dep+long],}
        else:
            f ={'year': year,
                'dossard': sample['Dossard'].tolist()[i_dep:len(sample)],
                'name': sample['Nom-Prenom'].tolist()[i_dep:len(sample)],
                'tel': sample['N° Portable'].tolist()[i_dep:len(sample)],
                'ville': sample['ville'].tolist()[i_dep:len(sample)],
                'lieu': sample['hotel'].tolist()[i_dep:len(sample)],}

        print(f)
        print()
        doc = DocxTemplate(template_path_docx) 
        doc.render(f)
    
        output_file = Path.home() / Path(r"CTG\BRA-BRO-BG\BRA\DATA\transport\essai"+str(idx)+".docx")
        doc.save(output_file)
        convert(output_file) # Creates a pdf file
        output_file = Path.home() / Path(r"CTG\BRA-BRO-BG\BRA\DATA\transport\essai"+str(idx)+".pdf")
        file_list.append(output_file)
        idx += 1
        
        
        
    
    merger = PdfFileMerger(strict=False)
    
    for pdf in file_list:
        merger.append(str(pdf))

    output_file = Path.home() / Path(r"CTG\BRA-BRO-BG\BRA\DATA\transport\merge.pdf")
    merger.write(str(output_file))
    merger.close()
    
    for pdf in file_list:
        os.remove(pdf)
        file = Path(pdf).stem+".docx"
        path = Path(os.path.dirname(pdf)) / file
        os.remove(path)


n_sample = 27
sample = make_sample(n_sample)
sample
make_etiquettes(sample)
print("SUCCESSFULLY ENDED")