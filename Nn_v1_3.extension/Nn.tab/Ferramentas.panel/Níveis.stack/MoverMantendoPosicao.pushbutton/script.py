# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
import System
import tempfile
import os

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

def get_solids(element):
    opt = Options()
    opt.ComputeReferences = True
    geometry = element.get_Geometry(opt)
    solids = []
    def extract_solid(geom_obj):
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            solids.append(geom_obj)
        elif isinstance(geom_obj, GeometryInstance):
            for inst_geom in geom_obj.GetInstanceGeometry():
                extract_solid(inst_geom)
    if geometry:
        for obj in geometry: extract_solid(obj)
    return solids

selection = revit.get_selection()
if not selection:
    forms.alert("Selecione os elementos primeiro.")
    script.exit()

template_path = forms.pick_file(file_ext='rft')
if not template_path:
    script.exit()

with revit.TransactionGroup("NnBim: Converter para RFA"):
    for el in selection:
        solidos = get_solids(el)
        if not solidos: continue
        bbox = el.get_BoundingBox(None)
        ponto_insercao = (bbox.Max + bbox.Min) / 2.0
        ponto_insercao = XYZ(ponto_insercao.X, ponto_insercao.Y, bbox.Min.Z)
        fam_doc = app.NewFamilyDocument(template_path)
        with Transaction(fam_doc, "Criar Geometria") as t_fam:
            t_fam.Start()
            mover = Transform.CreateTranslation(-ponto_insercao)
            for s in solidos:
                FreeFormElement.Create(fam_doc, SolidUtils.CreateTransformed(s, mover))
            t_fam.Commit()
        temp_file = os.path.join(tempfile.gettempdir(), "Nn_RFA_{}.rfa".format(System.Guid.NewGuid()))
        fam_doc.SaveAs(temp_file)
        loaded_fam = fam_doc.LoadFamily(doc)
        fam_doc.Close(False)
        symb = doc.GetElement(loaded_fam.GetFamilySymbolIds()[0])
        with Transaction(doc, "Inserir RFA") as t_proj:
            t_proj.Start()
            if not symb.IsActive: symb.Activate()
            doc.Create.NewFamilyInstance(ponto_insercao, symb, Structure.StructuralType.NonStructural)
            t_proj.Commit()

