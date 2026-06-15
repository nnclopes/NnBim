# -*- coding: utf-8 -*-
"""
NnBim: CATALOGO DE SISTEMAS
DESCRICAO:
Gera amostras fisicas de tipos de Parede organizadas em grade.
"""

__title__ = 'Catalogar Família\n de Sistema' # O acento aqui pode, pois esta dentro da string
__author__ = 'NnBim Dev'

import clr
from Autodesk.Revit.DB import *
from pyrevit import revit, forms

# Referencias globais
doc = revit.doc

def gerar_mostruario():
    # Coleta tipos de parede
    wall_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()
    
    # Filtro para ignorar paredes empilhadas (Stacked) que podem dar erro na criacao simples
    tipos_validos = [t for t in wall_types if t.Kind != WallKind.Stacked]
    
    if not tipos_validos:
        forms.alert("Nenhum tipo de parede encontrado.")
        return

    nivel = doc.ActiveView.GenLevel
    if not nivel:
        forms.alert("Por favor, execute em uma Planta de Piso.")
        return

    espacamento_y = 5.0 
    comprimento = 3.28 
    y_atual = 0

    with revit.Transaction("NnBim: Gerar Mostruario"):
        for w_type in tipos_validos:
            p1 = XYZ(0, y_atual, 0)
            p2 = XYZ(comprimento, y_atual, 0)
            linha = Line.CreateBound(p1, p2)
            
            try:
                # Criar Parede
                Wall.Create(doc, linha, w_type.Id, nivel.Id, 8.0, 0, False, False)
                
                # Criar Texto (Usando ID fixo para evitar erro de busca)
                txt_type = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElementId()
                ponto_txt = XYZ(-1, y_atual, 0)
                
                # Pega o nome do parametro de tipo com seguranca
                nome_param = w_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                nome_str = nome_param.AsString() if nome_param else w_type.Name
                
                TextNote.Create(doc, doc.ActiveView.Id, ponto_txt, nome_str, txt_type)
                
                y_atual -= espacamento_y
            except:
                continue

    forms.alert("Catalogo de sistemas gerado!", title="Sucesso NnBim")

if __name__ == "__main__":
    gerar_mostruario()