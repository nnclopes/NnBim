# -*- coding: utf-8 -*-
"""
NnBim Modelagem: Auto Forros
DESCRICAO:
Automatiza a modelagem de forros baseando-se nos limites dos Ambientes (Rooms).
Identifica automaticamente ilhas (vaos) e recortes internos.

COMO USAR:
1. Esteja em uma Planta de Piso ou de Forro Refletido.
2. Selecione os Ambientes (Rooms) desejados.
3. Clique no botao 'Ambientes para Forro'.
4. Na janela, escolha o Tipo de Forro e digite a Altura (ex: 2.70).
5. Clique em 'Gerar Forros'.
"""
__title__ = "Ambientes para Forro"
__author__ = "NnBim Dev"

from Autodesk.Revit.DB import *
from pyrevit import forms, revit
import clr

from Snippets._selection import get_selected_rooms
from Snippets._context_manager import ef_Transaction
from GUI.Tools.CreateFromRooms import CreateFromRooms

clr.AddReference("System")
from System.Collections.Generic import List

doc          = revit.doc
uidoc        = revit.uidoc
active_level = doc.ActiveView.GenLevel


def create_ceilings(rooms, ceil_type, offset):
    """Cria os elementos de forro no Revit."""
    ceilings = []

    with ef_Transaction(doc, __title__, debug=False):
        for room in rooms:
            try:
                area_param = room.get_Parameter(BuiltInParameter.ROOM_AREA)
                if not area_param or area_param.AsDouble() <= 0:
                    continue

                room_boundaries = room.GetBoundarySegments(SpatialElementBoundaryOptions())
                curveLoopList   = List[CurveLoop]()

                for roomBoundary in room_boundaries:
                    room_curve_loop = CurveLoop()
                    for segment in roomBoundary:
                        room_curve_loop.Append(segment.GetCurve())
                    curveLoopList.Add(room_curve_loop)

                if curveLoopList:
                    ceiling = Ceiling.Create(doc, curveLoopList, ceil_type.Id, active_level.Id)
                    ceilings.append(ceiling)

                    param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
                    if param:
                        param.Set(offset)

            except Exception as e:
                print("Erro no Room {}: {}".format(room.Id, e))

    return ceilings


def main():
    if not active_level:
        forms.alert(
            "Execute este comando em uma Planta de Piso ou Forro Refletido.",
            title=__title__,
            exitscript=True
        )

    selected_rooms = get_selected_rooms()
    if not selected_rooms:
        return

    all_ceil_types  = FilteredElementCollector(doc).OfClass(CeilingType).OfCategory(BuiltInCategory.OST_Ceilings)
    dict_ceil_types = {Element.Name.GetValue(fr): fr for fr in all_ceil_types}

    if not dict_ceil_types:
        forms.alert("Nenhum Tipo de Forro carregado no projeto.", title=__title__, exitscript=True)

    ui = CreateFromRooms(
        dict_ceil_types,
        title=__title__,
        label="Selecione o Tipo de Forro:",
        button_name='Gerar Forros'
    )

    selected_type = ui.selected_type
    offset_value  = ui.offset

    if not selected_type:
        return

    new_ceilings = create_ceilings(selected_rooms, selected_type, offset_value)

    if new_ceilings:
        new_ids = List[ElementId]([c.Id for c in new_ceilings])
        uidoc.Selection.SetElementIds(new_ids)
        forms.toast("{} Forros criados com sucesso!".format(len(new_ceilings)))


if __name__ == "__main__":
    main()