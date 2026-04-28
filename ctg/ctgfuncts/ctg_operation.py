__all__ = ['create_justificatif']

import os
import shutil
from datetime import date, datetime
from pathlib import Path
from tkinter import messagebox

from PyPDF2 import PdfMerger
from PIL import Image
from docx2pdf import convert

import ctg.ctggui.guiglobals as gg


def create_justificatif(ymd,year,intitulle,somme):
    ctg_path_finance =Path.home() / Path(gg.nextcloud) 
    ctg_path_finance = ctg_path_finance / Path('BASE_FINANCES_CTG') / Path(str(year)) / Path('COMPTABILITE-COURANTE')
    ctg_path_finance = ctg_path_finance / Path('1_JUSTIFICATIFS COMPTABILITE COURANTE/Temp') 
    
    if os.path.isdir(ctg_path_finance):

        docx_list = [ctg_path_finance / Path(x) for x in os.listdir(ctg_path_finance) if x.endswith(".docx")]
        for idx,file in enumerate(docx_list):
            convert(file, ctg_path_finance / Path(f"docx_{idx}.pdf"))
            
        image_list = [Image.open(ctg_path_finance / Path(x)) for x in os.listdir(ctg_path_finance) if x.endswith(".jpg")]
        if image_list:
            image_list[0].save(ctg_path_finance / Path("jpg.pdf"), save_all=True, append_images=image_list[1:])
        
        image_list = [Image.open(ctg_path_finance / Path(x)) for x in os.listdir(ctg_path_finance) if x.endswith(".png")]
        if image_list:
            image_list[0].save(ctg_path_finance / Path("png.pdf"), save_all=True, append_images=image_list[1:])        
 
        file_list = [x for x in os.listdir(ctg_path_finance) if x.endswith('.pdf')]
        merger = PdfMerger()
        for file in file_list:
            merger.append(ctg_path_finance / Path(file))
        merger.write(ctg_path_finance / Path("merge.pdf"))
        merger.close()
                
            
    else:
        os.mkdir(ctg_path_finance)
    
    
    
    # Get today's date
    
    new_name = f'xxx-{ymd}-{intitulle}-({somme}€).pdf'
    
    # Source and destination paths
    source = ctg_path_finance / Path("merge.pdf")
    destination = ctg_path_finance.parent / Path(new_name)
    
    # Move the file
    _ =shutil.move(source, destination)
    messagebox.showinfo("showinfo", f'le fichier\n {destination }\n a été créé')

