# -*- coding: utf-8 -*-
__title__   = 'Gerador de\nDetalhes'
__doc__     = 'Gera Vistas Típicas (evita duplicatas) de Locais ou Vínculos.'
__author__  = 'Nívea Lopes - NnBim'
__credits__ = ['Erik Frits', 'Gemini']
__version__ = '4.4.0'

import math
import re 
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script, DB

doc = revit.doc
uidoc = revit.uidoc

# ==============================================================================
# 1. CLASSES AUXILIARES
# ==============================================================================

class NnBim_LinkFilter(ISelectionFilter):
    def __init__(self, cat_id, doc_alvo):
        self.cat_id = cat_id
        self.doc_alvo = doc_alvo
    def AllowElement(self, element):
        if element.Category and element.Category.Id.IntegerValue == self.cat_id.IntegerValue:
            return True
        return False
    def AllowReference(self, reference, point):
        try:
            elem = self.doc_alvo.GetElement(reference.LinkedElementId)
            if elem and elem.Category and elem.Category.Id.IntegerValue == self.cat_id.IntegerValue:
                return True
        except: return False
        return False

def rotate_vector(vector, rotation_rad):
    vx, vy = vector.X, vector.Y
    rx = vx * math.cos(rotation_rad) - vy * math.sin(rotation_rad)
    ry = vx * math.sin(rotation_rad) + vy * math.cos(rotation_rad)
    return XYZ(rx, ry, vector.Z)

class ElementProperties():
    origin = None; vector = None; width = None; height = None; depth = None
    def __init__(self, el, transform=None):
        self.el = el
        self.transform = transform
        self.get_geometry()
    def get_geometry(self):
        BB = self.el.get_BoundingBox(None)
        if not BB: return
        self.width  = (BB.Max.X - BB.Min.X)
        self.height = (BB.Max.Z - BB.Min.Z)
        self.depth  = (BB.Max.Y - BB.Min.Y)
        local_origin = (BB.Max + BB.Min) / 2
        local_vector = XYZ.BasisX
        try:
            if hasattr(self.el.Location, 'Rotation'):
                local_vector = rotate_vector(local_vector, self.el.Location.Rotation)
            elif hasattr(self.el.Location, 'Curve') and self.el.Location.Curve:
                p0 = self.el.Location.Curve.GetEndPoint(0)
                p1 = self.el.Location.Curve.GetEndPoint(1)
                local_vector = (p1 - p0).Normalize()
        except: pass
        if self.transform:
            self.origin = self.transform.OfPoint(local_origin)
            self.vector = self.transform.OfVector(local_vector)
        else:
            self.origin = local_origin
            self.vector = local_vector

# ==============================================================================
# 2. MOTOR DE GERAÇÃO
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

    def generate(self, name_base, view_type_id, template_id=None):
        views = []
        sufixos = [("_Elevacao", "elevation"), ("_Corte", "cross"), ("_Planta", "plan")]
        try:
            for suffix, mode in sufixos:
                bbox = self.create_section_box(mode)
                view = ViewSection.CreateSection(self.doc, view_type_id, bbox)
                final_name = name_base + suffix
                self.safe_rename(view, final_name)
                if template_id and template_id != ElementId.InvalidElementId:
                    try: view.ViewTemplateId = template_id
                    except: pass
                views.append(view)
        except Exception as e:
            print("Erro ao criar vista: " + str(e))
        return views

    def safe_rename(self, view, name):
        for i in range(50):
            try:
                new_name = name if i==0 else "{} ({})".format(name, i)
                view.Name = new_name
                break
            except: continue

# ==============================================================================
# 3. FUNÇÕES DE NOMENCLATURA & LEITURA
# ==============================================================================
def clean_name(texto):
    if not texto: return ""
    return re.sub(r'[\\:{}\[\]|;<>?`~]', '', str(texto))

def get_param_value(element, option_name):
    if not option_name or option_name == "(Nenhum)": return None
    if option_name == "ID do Elemento": return str(element.Id)
    
    val = None
    try:
        # Instância
        if option_name == "Marca (Mark)":
            p = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p: val = p.AsString()
        elif option_name == "Comentários":
            p = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p: val = p.AsString()
        elif option_name == "Nome do Ambiente":
            p = element.get_Parameter(BuiltInParameter.ROOM_NAME)
            if p: val = p.AsString()
        elif option_name == "Número do Ambiente":
            p = element.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            if p: val = p.AsString()
        
        # Tipo
        if not val:
            doc_elem = element.Document
            typ = doc_elem.GetElement(element.GetTypeId())
            if typ:
                if option_name == "Nome do Tipo":
                    val = Element.Name.GetValue(typ)
                elif option_name == "Marca de Tipo":
                    p = typ.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
                    if p: val = p.AsString()
                elif option_name == "Descrição":
                    p = typ.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION)
                    if p: val = p.AsString()
    except: pass
    return val

# ==============================================================================
# 4. EXECUÇÃO
# ==============================================================================

# PASSO 0: MODO
modo_origem = forms.CommandSwitchWindow.show(['Modelo Atual (Local)', 'Vínculo Revit (Link)'], message="Origem dos elementos?")
if not modo_origem: script.exit()
is_link = (modo_origem == 'Vínculo Revit (Link)')

# PASSO 1: CATEGORIA
all_cats = doc.Settings.Categories
cat_dict = {}
for cat in all_cats:
    try:
        if (cat.CategoryType == CategoryType.Model or "Rooms" in cat.Name) and (cat.CanAddSubcategory or "Rooms" in cat.Name):
            cat_dict[cat.Name] = cat.Id
    except: continue

cat_escolhida = forms.SelectFromList.show(sorted(cat_dict.keys()), title="1. Categoria", button_name="Selecionar", multiselect=False)
if not cat_escolhida: script.exit()

# PASSO 2: SELEÇÃO
elementos_brutos = [] 

try:
    if is_link:
        with forms.WarningBar(title="Selecione VÁRIOS elementos no VÍNCULO."):
            refs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Selecione")
        for ref in refs:
            link_instance = doc.GetElement(ref.ElementId) 
            link_transform = link_instance.GetTotalTransform()
            link_doc = link_instance.GetLinkDocument()
            elem = link_doc.GetElement(ref.LinkedElementId)
            if elem.Category.Id.IntegerValue == cat_dict[cat_escolhida].IntegerValue:
                elementos_brutos.append( (elem, link_transform) )
    else:
        filtro = NnBim_SelectionFilter(cat_dict[cat_escolhida])
        with forms.WarningBar(title="Selecione VÁRIOS elementos LOCAIS."):
            refs = uidoc.Selection.PickObjects(ObjectType.Element, filtro, "Selecione")
        for ref in refs:
            elementos_brutos.append( (doc.GetElement(ref), None) )

except: script.exit()

if not elementos_brutos: 
    forms.alert("Nada selecionado.")
    script.exit()

# PASSO 3: NOMENCLATURA & FILTRO
prefixo = forms.ask_for_string(default="DET_", prompt="Prefixo:", title="Nome")
if prefixo is None: script.exit()

opcoes_params = ['Marca (Mark)', 'Nome do Tipo', 'Marca de Tipo', 'Comentários', 'ID do Elemento']
p1 = forms.SelectFromList.show(opcoes_params, title="AGRUPADOR (Itens com esse valor igual serão ignorados):", button_name="Usar este agrupador", multiselect=False)
if not p1: script.exit()

p_opcionais = forms.SelectFromList.show(opcoes_params, title="(Opcional) Sufixos do nome:", button_name="Continuar", multiselect=True)
separador_final = "_"

# --- OTIMIZAÇÃO: FILTRAGEM DE DUPLICATAS ---
elementos_unicos = []
chaves_vistas = [] # Lista para guardar o que já foi processado

print("Processando {} elementos selecionados...".format(len(elementos_brutos)))

for item in elementos_brutos:
    el = item[0]
    
    # Descobre o valor do Parâmetro Principal (Agrupador)
    chave = get_param_value(el, p1)
    
    # Se o parâmetro estiver vazio, usa o ID (para não perder o elemento)
    if not chave: chave = str(el.Id)
    
    if chave not in chaves_vistas:
        chaves_vistas.append(chave)
        elementos_unicos.append(item)
    else:
        # Pula silenciosamente pois é duplicata
        pass

# PASSO 4: TIPO DE VISTA
view_types = {}
for v in FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements():
    if v.ViewFamily == ViewFamily.Section:
        name = Element.Name.GetValue(v)
        if name: view_types[name] = v

tipo_vista_nome = forms.SelectFromList.show(sorted(view_types.keys()), title="Tipo de Vista", button_name="Confirmar")
if not tipo_vista_nome: script.exit()
view_type_obj = view_types[tipo_vista_nome]

# PASSO 5: TEMPLATE
template_id_final = None
try:
    param_default = view_type_obj.get_Parameter(BuiltInParameter.VIEW_DEFAULT_TEMPLATE_ID)
    if param_default and param_default.AsElementId() != ElementId.InvalidElementId:
        template_id_final = param_default.AsElementId()
except: pass

if not template_id_final:
    templates = {v.Name: v.Id for v in FilteredElementCollector(doc).OfClass(View).ToElements() if v.IsTemplate}
    if templates:
        res_t = forms.SelectFromList.show(['(Nenhum)'] + sorted(templates.keys()), title="Template?", button_name="Gerar")
        if res_t and res_t != '(Nenhum)': template_id_final = templates[res_t]

# PASSO 6: EXECUÇÃO DOS ÚNICOS
t = DB.Transaction(doc, "NnBim: V4.4 Gerar Vistas Típicas")
t.Start()
count = 0

for item in elementos_unicos:
    try:
        el, trans = item[0], item[1]
        props = ElementProperties(el, transform=trans)
        if not props.width: continue

        nome_parts = []
        val1 = get_param_value(el, p1) # Já sabemos que é único
        if not val1: val1 = str(el.Id)
        nome_parts.append(clean_name(val1))

        if p_opcionais:
            for opt in p_opcionais:
                val = get_param_value(el, opt)
                if val: nome_parts.append(clean_name(val))
        
        nome_base = prefixo + separador_final.join(nome_parts)
        
        gen = SectionGenerator(doc, props)
        gen.generate(nome_base, view_type_obj.Id, template_id_final)
        count += 1
    except: pass

t.Commit()

# Relatório Final Inteligente
economizados = len(elementos_brutos) - count
forms.alert(
    "Sucesso!\n\nSelecionados: {}\nGerados: {} (Itens Típicos)\nIgnorados: {} (Duplicatas)".format(
        len(elementos_brutos), count, economizados
    ), 
    title="NnBim Otimização"
)