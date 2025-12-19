# -*- coding: utf-8 -*-
__title__   = 'Gerador de\nDetalhes'
__doc__     = 'Gera Planta, Corte e Elevação selecionando elementos na tela.'
__author__  = 'Nívea Lopes - NnBim'
__credits__ = ['Erik Frits', 'Gemini']
__version__ = '3.4.0'

import math
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script, DB

doc = revit.doc
uidoc = revit.uidoc

# ==============================================================================
# 1. CLASSE DE FILTRO
# ==============================================================================
class NnBim_SelectionFilter(ISelectionFilter):
    def __init__(self, categoria_id):
        self.categoria_id = categoria_id

    def AllowElement(self, element):
        if element.Category and element.Category.Id.IntegerValue == self.categoria_id.IntegerValue:
            return True
        return False

    def AllowReference(self, reference, point):
        return False

# ==============================================================================
# 2. GEOMETRIA
# ==============================================================================
def rotate_vector(vector, rotation_rad):
    vector_x = vector.X
    vector_y = vector.Y
    rotated_x = vector_x * math.cos(rotation_rad) - vector_y * math.sin(rotation_rad)
    rotated_y = vector_x * math.sin(rotation_rad) + vector_y * math.cos(rotation_rad)
    return XYZ(rotated_x, rotated_y, vector.Z)

class ElementProperties():
    origin = None 
    vector = None 
    width  = None 
    height = None
    depth  = None
    
    def __init__(self, el):
        self.el = el
        self.get_geometry()

    def get_geometry(self):
        BB = self.el.get_BoundingBox(None)
        if not BB: return

        self.width  = (BB.Max.X - BB.Min.X)
        self.height = (BB.Max.Z - BB.Min.Z)
        self.depth  = (BB.Max.Y - BB.Min.Y)
        self.origin = (BB.Max + BB.Min) / 2
        self.vector = XYZ.BasisX

        try:
            if hasattr(self.el.Location, 'Rotation'):
                self.vector = rotate_vector(self.vector, self.el.Location.Rotation)
            elif hasattr(self.el.Location, 'Curve') and self.el.Location.Curve:
                p0 = self.el.Location.Curve.GetEndPoint(0)
                p1 = self.el.Location.Curve.GetEndPoint(1)
                self.vector = (p1 - p0).Normalize()
        except:
            pass

# ==============================================================================
# 3. GERADOR
# ==============================================================================
class SectionGenerator():
    def __init__(self, doc, props):
        self.doc = doc
        self.props = props
        self.offset = 1.5 

    def create_transform(self, mode='elevation'):
        trans = Transform.Identity
        trans.Origin = self.props.origin
        vector = self.props.vector.Normalize()

        if mode == 'elevation':
            trans.BasisX = vector
            trans.BasisY = XYZ.BasisZ
            trans.BasisZ = vector.CrossProduct(XYZ.BasisZ)
        elif mode == 'cross':
            vec_cross = vector.CrossProduct(XYZ.BasisZ)
            trans.BasisX = vec_cross
            trans.BasisY = XYZ.BasisZ
            trans.BasisZ = vec_cross.CrossProduct(XYZ.BasisZ)
        elif mode == 'plan':
            trans.BasisX = vector
            trans.BasisY = vector.CrossProduct(XYZ.BasisZ)
            trans.BasisZ = XYZ.BasisZ
        return trans

    def create_section_box(self, mode):
        bbox = BoundingBoxXYZ()
        bbox.Transform = self.create_transform(mode)
        
        W, H, D = self.props.width/2, self.props.height/2, self.props.depth/2
        off = self.offset

        if mode == 'elevation':
            bbox.Min, bbox.Max = XYZ(-W-off, -H-off, 0), XYZ(W+off, H+off, D+off)
        elif mode == 'cross':
            bbox.Min, bbox.Max = XYZ(-D-off, -H-off, 0), XYZ(D+off, H+off, W+off)
        elif mode == 'plan':
            bbox.Min, bbox.Max = XYZ(-W-off, -D-off, 0), XYZ(W+off, D+off, H+off)
        return bbox

    def generate(self, name_base, view_type_id):
        views = []
        try:
            for suffix, mode in [("_Elevacao", "elevation"), ("_Corte", "cross"), ("_Planta", "plan")]:
                bbox = self.create_section_box(mode)
                view = ViewSection.CreateSection(self.doc, view_type_id, bbox)
                self.rename(view, name_base + suffix)
                views.append(view)
        except:
            pass
        return views

    def rename(self, view, name):
        for i in range(20):
            try:
                view.Name = name if i==0 else name + " (" + str(i) + ")"
                break
            except: continue

# ==============================================================================
# 4. EXECUÇÃO
# ==============================================================================

# PASSO 1: Escolher Categoria
all_cats = doc.Settings.Categories
cat_dict = {}

for cat in all_cats:
    try:
        c_name = cat.Name 
        c_type = cat.CategoryType
        is_model = (c_type == CategoryType.Model)
        is_room  = ("Rooms" in c_name or "Ambientes" in c_name)
        
        if (is_model or is_room) and (cat.CanAddSubcategory or is_room): 
            cat_dict[c_name] = cat.Id
    except:
        continue

nomes_ordenados = sorted(cat_dict.keys())
if not nomes_ordenados: forms.alert("Erro categorias.", exitscript=True)

escolha = forms.SelectFromList.show(
    nomes_ordenados, 
    title="1. Qual categoria?",
    button_name="Selecionar na Tela",
    multiselect=False
)

if not escolha: script.exit()

cat_id_selecionada = cat_dict[escolha]

# PASSO 2: Seleção na Tela
try:
    filtro = NnBim_SelectionFilter(cat_id_selecionada)
    
    mensagem = "Clique nos itens ({}) e em CONCLUIR.".format(escolha)
    
    with forms.WarningBar(title=mensagem):
        referencias = uidoc.Selection.PickObjects(ObjectType.Element, filtro, "Selecione")
    
    elementos = [doc.GetElement(ref) for ref in referencias]
except:
    script.exit()

if not elementos: script.exit()

# PASSO 3: Tipo de Vista (CORREÇÃO SEGURA AQUI)
view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
section_types = {}

# Loop seguro para evitar erro de .Name
for v in view_types:
    try:
        if v.ViewFamily == ViewFamily.Section:
            # Tenta pegar o nome de forma segura
            v_name = Element.Name.GetValue(v)
            if v_name:
                section_types[v_name] = v.Id
    except:
        continue # Se um tipo de vista estiver corrompido, pula ele

if not section_types:
    forms.alert("Não encontrei nenhum Tipo de Vista de Corte no projeto.", exitscript=True)

tipo_vista = forms.SelectFromList.show(
    sorted(section_types.keys()), 
    title="2. Template de Vista",
    button_name="Gerar"
)

if not tipo_vista: script.exit()

# PASSO 4: Gerar
t = Transaction(doc, "NnBim: Gerar Vistas")
t.Start()

count = 0
for el in elementos:
    try:
        props = ElementProperties(el)
        if not props.width: continue

        gen = SectionGenerator(doc, props)
        
        nome = "DET_" + str(el.Id)
        if "Ambientes" in escolha or "Rooms" in escolha:
            p = el.get_Parameter(BuiltInParameter.ROOM_NAME)
            if p: nome = "DET_" + p.AsString()
        else:
            try:
                typ = doc.GetElement(el.GetTypeId())
                if typ: nome = "DET_" + typ.Name
            except: pass

        gen.generate(nome, section_types[tipo_vista])
        count += 1
    except:
        pass

t.Commit()

if count > 0:
    forms.alert("Sucesso! {} elementos.".format(count))