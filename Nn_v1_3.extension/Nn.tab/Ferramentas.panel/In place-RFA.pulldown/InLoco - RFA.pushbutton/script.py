# -*- coding: utf-8 -*-
"""
NnBim: InLoco - RFA
DESCRICAO:
Substitui elementos In-Place por familias carregaveis existentes no projeto.
Posiciona a familia no centro do bounding box do elemento original
e oculta o original na vista atual.

COMO USAR:
1. Selecione os elementos In-Place na vista.
2. Execute o script.
3. Escolha a familia carregavel de destino.
"""
__title__ = 'InLoco\nRFA'
__author__ = 'NnBim Dev'

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
# CORRECAO: import que estava faltando
import System.Collections.Generic

doc = __revit__.ActiveUIDocument.Document


def get_solid(element):
    opt      = Options()
    geometry = element.get_Geometry(opt)
    if geometry:
        for obj in geometry:
            if isinstance(obj, Solid) and obj.Volume > 0:
                return obj
    return None


selection = revit.get_selection()
if not selection:
    script.exit()

cat_id   = selection[0].Category.Id
familias = FilteredElementCollector(doc).OfClass(FamilySymbol)
opcoes   = {
    "{} : {}".format(s.Family.Name, s.Name): s
    for s in familias
    if s.Category and s.Category.Id == cat_id
}

if not opcoes:
    forms.alert("Nenhuma familia carregavel da mesma categoria encontrada.")
    script.exit()

escolha = forms.SelectFromList.show(sorted(opcoes.keys()), title="Escolha a Familia")
if not escolha:
    script.exit()

symb = opcoes[escolha]

# CORRECAO: transaction duplicada removida
with revit.Transaction("NnBim: InLoco para RFA"):
    if not symb.IsActive:
        symb.Activate()

    elementos_ocultar = System.Collections.Generic.List[ElementId]()

    for el in selection:
        bbox = el.get_BoundingBox(None)
        if not bbox:
            continue
        pt = (bbox.Max + bbox.Min) / 2.0
        pt = XYZ(pt.X, pt.Y, bbox.Min.Z)
        doc.Create.NewFamilyInstance(pt, symb, Structure.StructuralType.NonStructural)
        elementos_ocultar.Add(el.Id)

    if elementos_ocultar.Count > 0:
        doc.ActiveView.HideElements(elementos_ocultar)

forms.toast("{} elemento(s) substituido(s)!".format(len(selection)))