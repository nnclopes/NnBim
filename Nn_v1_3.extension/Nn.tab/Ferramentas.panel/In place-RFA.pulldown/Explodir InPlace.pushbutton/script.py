# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
import System.Collections.Generic

doc = __revit__.ActiveUIDocument.Document

def get_solids(element):
    opt = Options()
    opt.ComputeReferences = True
    geometry = element.get_Geometry(opt)
    solids = []
    def extract_solid(geom_obj):
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            try:
                separados = SolidUtils.SplitVolumes(geom_obj)
                for sep in separados:
                    if sep.Volume > 0: solids.append(sep)
            except: solids.append(geom_obj)
        elif isinstance(geom_obj, GeometryInstance):
            for inst_geom in geom_obj.GetInstanceGeometry():
                extract_solid(inst_geom)
    if geometry:
        for obj in geometry: extract_solid(obj)
    return solids

selection = revit.get_selection()
if not selection:
    forms.alert("Selecione o In-Loco.")
    script.exit()

in_place_el = selection[0]
solidos = get_solids(in_place_el)

if solidos:
    with Transaction(doc, "NnBim: Explodir") as t:
        t.Start()
        for s in solidos:
            geo_list = System.Collections.Generic.List[GeometryObject]()
            geo_list.Add(s)
            ds = DirectShape.CreateElement(doc, in_place_el.Category.Id)
            ds.SetShape(geo_list)
        doc.Delete(in_place_el.Id)
        t.Commit()
    forms.toast("Sucesso!")