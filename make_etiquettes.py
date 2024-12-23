import ctg.ctggui as ctgg
import ctg.ctgfuncts as ctg
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
    n_sample = 150
    year = 2024
    hebergement = [('St julien de Montdenis','Hôtel Lancheton'),
                   ('St Michel de Maurienne','Hôtel  le Galibier'),
                   ('St Michel de Maurienne','Hôtel  Varcin'),
                   ('St Michel de Maurienne','Hôtel le Savoy'),
                   ('St Michel de Maurienne','Hôtel camping le  Marintan'),
                   ('St Michel de Maurienne','Lycée'),
                   ('St Michel de Maurienne','Collège'),
                   ('VALMEINIER 1500 et 1800','le grand Fourchon'),
                   ('VALMEINIER 1500 et 1800','Hôtel les Carretes'),
                   ('VALLOIRE-les Granges','Hôtel le Tatami'),
                   ('VALLOIRE (centre village)','Hôtel de la Poste'),
                   ('VALLOIRE (centre village)','Hôtel les Mélèzes'),
                   ('VALLOIRE (centre village)','Hôtel Christiana'),
                   ('VALLOIRE (centre village)',"Hôtel L'Aiguille Noire"),
                   ('VALLOIRE (centre village)',"Hôtel le Centre"),
                   ('VALLOIRE (centre village)',"Hôtel La Maison Rapin"),
                   ('VALLOIRE (centre village)',"Le Grand Hôtel"),
                   ('VALLOIRE - Les Verneys',"Hôtel le Relais du Galibier"),
                   ('VALLOIRE - Les Verneys',"Hôtel le Crêt Rond"),
                   ('VALLOIRE - Les Verneys',"Gîte les Réaux"),
                   ('VALLOIRE - Les Verneys',"Pulka"),
                   ('VALLOIRE - Les Verneys',"Le Val D'Or"),
                   ('Valmeinier 1500',"L'Arména"),
                   ]
    ctg_path = r"c:\users\franc\CTG\SORTIES"
    effectif = ctg.EffectifCtg(year,ctg_path)
    sample = effectif.effectif.sample(n=n_sample)[['Nom','Prénom','N° Portable']]
    sample['ville-hotel'] = [random.choice(hebergement) for _ in range(n_sample)]
    sample[['ville', 'hotel']] = pd.DataFrame(sample['ville-hotel'].tolist(), index=sample.index)
    sample['Nom-Prenom'] = sample.apply(lambda row: f"{row['Nom']} {row['Prénom']}",axis=1)
    sample = sample.drop(['ville-hotel','Nom','Prénom'],axis=1)
    sample["Dossard"] = [random.randint(1,800) for _ in range(n_sample)]
    sample = sample.fillna('inconnu')
    return sample

def make_etiquettes(sample):
    template_path_docx = Path.home() / r"CTG\BRA-BRO-BG\BRA\DATA\EtiquetteBagage_BRA.docx"
    output_file = Path.home() / r"Temp/essai.docx"
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
        
        doc = DocxTemplate(template_path_docx) 
        doc.render(f)
        output_file = Path.home() / Path(r"Temp/essai"+str(idx)+".docx")
        doc.save(output_file)
        file_list.append(Path.home() / Path(r"Temp/essai"+str(idx)+".pdf"))
        idx += 1
        
        convert(output_file) # Creates a pdf file
        
    
    merger = PdfFileMerger(strict=False)
    
    for pdf in file_list:
        merger.append(str(pdf))
    
    merger.write(str(Path.home() / Path(r"Temp/merge.pdf")))
    merger.close()

n_sample = 105
sample = make_sample(n_sample)
make_etiquettes(sample)