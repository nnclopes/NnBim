# -*- coding: utf-8 -*-
__title__ = 'Explodir\nInPlace'
__author__ = 'Nivea'

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
import System.Collections.Generic

doc = __revit__.ActiveUIDocument.Document

def get_solids(element):
    """Lê a geometria pesada e separa cada sólido em uma lista"""
    opt = Options()
    opt.ComputeReferences = True
    geometry = element.get_Geometry(opt)
    solids = []
    
    def extract_solid(geom_obj):
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            try:
                # Força a quebra de volumes fundidos
                separados = SolidUtils.SplitVolumes(geom_obj)
                for sep in separados:
                    if sep.Volume > 0: solids.append(sep)
            except:
                solids.append(geom_obj)
        elif isinstance(geom_obj, GeometryInstance):
            for inst_geom in geom_obj.GetInstanceGeometry():
                extract_solid(inst_geom)

    if geometry:
        for obj in geometry: extract_solid(obj)
    return solids

# 1. PEGAR A FAMÍLIA IN-LOCO
selection = revit.get_selection()

if not selection:
    forms.alert("Selecione a família in-loco gigante primeiro!")
    script.exit()

in_place_el = selection[0]
cat_id = in_place_el.Category.Id

solidos = get_solids(in_place_el)

if not solidos:
    forms.alert("Nenhum sólido encontrado para explodir.")
    script.exit()

confirmar = forms.alert(
    "Vou explodir este in-loco em {} elementos individuais.\nPodemos prosseguir?".format(len(solidos)),
    title="Análise de Geometria", options=["Sim, explodir", "Cancelar"]
)

if confirmar != "Sim, explodir":
    script.exit()

# 2. EXPLODIR E CRIAR DIRECTSHAPES INDIVIDUAIS
t = Transaction(doc, "Explodir In-Loco Nn")
t.Start()

try:
    novos_ids = []
    total = 0
    
    with forms.ProgressBar(total=len(solidos), title="Individualizando peças...") as pb:
        for s in solidos:
            # Lista tipada do .NET que o Revit exige para criar o DirectShape
            geo_list = System.Collections.Generic.List[GeometryObject]()
            geo_list.Add(s)
            
            # Cria a forma direta na mesma categoria do in-loco original
            ds = DirectShape.CreateElement(doc, cat_id)
            ds.SetShape(geo_list)
            ds.Name = "Brise Individual"
            
            novos_ids.append(ds.Id)
            total += 1
            pb.update_progress(total)

    # Deleta a família in-loco aglomerada
    doc.Delete(in_place_el.Id)
    
    t.Commit()
    
    # Seleciona todas as novas peças soltas
    revit.get_selection().set_to(novos_ids)
    
    forms.alert("Pronto! {} elementos foram separados com sucesso e a massa original foi apagada.".format(total), title="Sucesso")

except Exception as e:
    t.RollBack()
    forms.alert("Erro ao explodir: {}".format(str(e)), title="Erro")