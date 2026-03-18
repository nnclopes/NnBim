# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = __revit__.ActiveUIDocument.Document

def get_solid(element):
    opt = Options()
    geometry = element.get_Geometry(opt)
    if geometry:
        for obj in geometry:
            if isinstance(obj, Solid) and obj.Volume > 0: return obj
    return None

selection = revit.get_selection()
if not selection:
    script.exit()

cat_id = selection[0].Category.Id
familias = FilteredElementCollector(doc).OfClass(FamilySymbol)
opcoes = {"{} : {}".format(s.Family.Name, s.Name): s for s in familias if s.Category.Id == cat_id}

escolha = forms.SelectFromList.show(sorted(opcoes.keys()), title="Escolha a Família")
if escolha:
    symb = opcoes[escolha]
    with Transaction(doc, "NnBim: Substituir"):
        t = Transaction(doc, "Substituir")
        t.Start()
        for el in selection:
            bbox = el.get_BoundingBox(None)
            pt = (bbox.Max + bbox.Min) / 2.0
            pt = XYZ(pt.X, pt.Y, bbox.Min.Z)
            if not symb.IsActive: symb.Activate()
            doc.Create.NewFamilyInstance(pt, symb, Structure.StructuralType.NonStructural)
            doc.ActiveView.HideElements(System.Collections.Generic.List[ElementId]([el.Id]))
        t.Commit()