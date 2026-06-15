# -*- coding: utf-8 -*-
"""
NnBim: EXPORTAR BIBLIOTECA
DESCRICAO:
Extrai familias carregaveis e organiza em pastas por categoria.
"""

__title__ = 'Catalogar Família\nExportar-RFA' # O acento aqui pode, pois esta dentro da string
__author__ = 'NnBim Dev'

import os
import clr
from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc

def exportar():
    # Usar nomes simples para variaveis evita erros de decode
    pasta = forms.pick_folder(title="Selecione o destino")
    if not pasta:
        return

    familias = FilteredElementCollector(doc).OfClass(Family).ToElements()
    
    with forms.ProgressBar(title="Exportando...", total=len(familias)) as pb:
        count = 0
        for fam in familias:
            count += 1
            pb.update_progress(count, len(familias))
            
            if fam.IsEditable:
                try:
                    # Garantir que o nome da categoria nao tenha caracteres invalidos
                    cat_raw = fam.FamilyCategory.Name if fam.FamilyCategory else "Sem_Categoria"
                    cat_nome = "".join([c for c in cat_raw if c.isalnum() or c in " _-"]).strip()
                    
                    caminho_cat = os.path.join(pasta, cat_nome)
                    if not os.path.exists(caminho_cat):
                        os.makedirs(caminho_cat)
                    
                    caminho_arq = os.path.join(caminho_cat, fam.Name + ".rfa")
                    
                    if not os.path.exists(caminho_arq):
                        fam_doc = doc.EditFamily(fam)
                        if fam_doc:
                            opt = SaveAsOptions()
                            opt.OverwriteExistingFile = True
                            fam_doc.SaveAs(caminho_arq, opt)
                            fam_doc.Close(False)
                except:
                    continue

    forms.alert("Exportacao concluida!")

if __name__ == "__main__":
    exportar()

