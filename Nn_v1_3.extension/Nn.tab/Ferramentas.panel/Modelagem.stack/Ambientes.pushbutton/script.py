# -*- coding: utf-8 -*-
"""
NnBim Modelagem: Auto Forros
DESCRIÇÃO:
Automatiza a modelagem de forros baseando-se nos limites dos Ambientes (Rooms). 
Identifica automaticamente ilhas (vãos) e recortes internos.

COMO USAR:
1. Esteja em uma Planta de Piso ou de Forro Refletido.
2. Selecione os Ambientes (Rooms) desejados.
3. Clique no botão 'Ambientes para Forro'.
4. Na janela, escolha o Tipo de Forro e digite a Altura (ex: 2.70).
5. Clique em 'Gerar Forros'.
"""

__title__ = "Ambientes para Forro"
__author__ = "NnBim Dev"

# 1. IMPORTS DO CORE
from Autodesk.Revit.DB import *
from pyrevit import forms, revit
import clr

# 2. IMPORTS DA PASTA LIB (Nn_v1_3.extension)
# Estes módulos devem estar em Nn_v1_3.extension\lib
from Snippets._selection import get_selected_rooms
from Snippets._context_manager import ef_Transaction
from GUI.Tools.CreateFromRooms import CreateFromRooms

# .NET COLLECTIONS
clr.AddReference("System")
from System.Collections.Generic import List

# VARIÁVEIS DE CONTEXTO
doc = revit.doc
uidoc = revit.uidoc
active_level = doc.ActiveView.GenLevel

def create_ceilings(rooms, ceil_type, offset):
    """Cria os elementos de forro no Revit."""
    ceilings = []
    
    # Usando o Context Manager personalizado da sua lib
    with ef_Transaction(doc, __title__, debug=False):
        for room in rooms:
            try:
                # Validação de área para evitar erros em rooms não colocados
                area_param = room.get_Parameter(BuiltInParameter.ROOM_AREA)
                if not area_param or area_param.AsDouble() <= 0:
                    continue

                # Obtenção dos contornos (Boundary)
                room_boundaries = room.GetBoundarySegments(SpatialElementBoundaryOptions())
                curveLoopList = List[CurveLoop]()

                for roomBoundary in room_boundaries:
                    room_curve_loop = CurveLoop()
                    for segment in roomBoundary:
                        curve = segment.GetCurve()
                        room_curve_loop.Append(curve)
                    curveLoopList.Add(room_curve_loop)

                # Criação do Forro (API Revit 2022+)
                if curveLoopList:
                    ceiling = Ceiling.Create(doc, curveLoopList, ceil_type.Id, active_level.Id)
                    ceilings.append(ceiling)
                    
                    # Define a altura (Offset from Level)
                    param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
                    if param:
                        param.Set(offset)
            except Exception as e:
                print("Erro no Room {}: {}".format(room.Id, e))
                
    return ceilings

def main():
    # 1. Verifica se há nível associado à vista atual
    if not active_level:
        forms.alert("Execute este comando em uma Planta de Piso ou Forro Refletido.", 
                    title=__title__, exitscript=True)

    # 2. Busca Ambientes (Snippet da lib)
    selected_rooms = get_selected_rooms()
    if not selected_rooms:
        return

    # 3. Mapeia Tipos de Forro
    all_ceil_types = FilteredElementCollector(doc).OfClass(CeilingType).OfCategory(BuiltInCategory.OST_Ceilings)
    dict_ceil_types = {Element.Name.GetValue(fr): fr for fr in all_ceil_types}

    if not dict_ceil_types:
        forms.alert("Nenhum Tipo de Forro carregado no projeto.", title=__title__, exitscript=True)

    # 4. Abre a Interface Customizada (GUI da lib)
    ui = CreateFromRooms(dict_ceil_types, 
                         title=__title__, 
                         label="Selecione o Tipo de Forro:",
                         button_name='Gerar Forros')
    
    # Resgata dados da UI
    selected_type = ui.selected_type
    offset_value = ui.offset

    if not selected_type:
        return

    # 5. Executa a criação
    new_ceilings = create_ceilings(selected_rooms, selected_type, offset_value)

    # 6. Seleciona os novos elementos e feedback
    if new_ceilings:
        new_ids = List[ElementId]([c.Id for c in new_ceilings])
        uidoc.Selection.SetElementIds(new_ids)
        forms.toast("{} Forros criados com sucesso!".format(len(new_ceilings)))

if __name__ == "__main__":
    main()
