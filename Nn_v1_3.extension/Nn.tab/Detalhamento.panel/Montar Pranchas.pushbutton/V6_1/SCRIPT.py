# -*- coding: utf-8 -*-
"""
Nome: Montar Prancha
Descricao: Diagramacao V6.1 (Calibrada 140mm + Lista Auto)
Autor: NnBim Dev
"""

import clr
import re
from collections import defaultdict

# Imports do Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Imports do pyRevit
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc

# --- 1. CONFIGURACOES GERAIS ---

# [GAPS EM MILIMETROS]
CFG_GAP_INT_X = 25  # Entre Elevacao e Corte
CFG_GAP_INT_Y = 45  # Entre Elevacao e Planta
CFG_GAP_GRID_X = 20 # Corredor vertical entre grupos
CFG_GAP_GRID_Y = 55 # Corredor horizontal

# [MARGENS E ZONAS PROIBIDAS]
CFG_MARGIN_LEFT = 25 # Margem esquerda (Encadernacao 0,02 + 0,005)
CFG_MARGIN_TOP = 20  # Margem superior

# [CALIBRACAO TECHPROJ]
# Baseado na cota 0,13 (Carimbo) + 0,01 (Borda) = 140mm
# Adicionamos +5mm de folga para seguranca visual.
CFG_MARGIN_RIGHT = 145 

# Altura reservada na base (alem do carimbo)
CFG_MARGIN_BOTTOM = 20

# NOME DO PARAMETRO PARA A LISTA DE DESENHOS
# (Deve ser Texto Multilinha e estar nas Folhas)
PARAM_LISTA_NOME = "Nn_IndiceDesenhos"

# Fator de Conversao (Mm -> Feet)
MM_TO_FT = 0.00328084

# --- 2. FUNCOES AUXILIARES ---

def is_view_placed(view):
    """Verifica se a vista ja esta em folha."""
    try:
        if view.ViewSheetId != ElementId.InvalidElementId: return True
    except AttributeError:
        try:
            p = view.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
            if p and p.AsString() and p.AsString() != "---": return True
        except: pass
    return False

def get_element_name(element):
    """Leitura segura do nome."""
    return Element.Name.GetValue(element)

def set_sheet_index(sheet, text_content):
    """Escreve a lista de desenhos no parametro da folha."""
    p = sheet.LookupParameter(PARAM_LISTA_NOME)
    if p:
        try:
            p.Set(text_content)
        except:
            print("Aviso: Nao foi possivel escrever no parametro '{}'.".format(PARAM_LISTA_NOME))
    else:
        # Tenta procurar por parametros compartilhados se o nome simples falhar
        pass 

def get_viewport_info(doc, sheet_id, view_id):
    """Recupera o numero e o titulo do viewport recem criado."""
    # Procura o viewport criado nesta folha para esta vista
    vps = FilteredElementCollector(doc).OfClass(Viewport).WhereElementIsNotElementType().ToElements()
    for vp in vps:
        if vp.SheetId == sheet_id and vp.ViewId == view_id:
            # Achou!
            number = vp.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_DETAIL_NUMBER).AsString()
            
            # Tenta pegar o Titulo na Folha
            p_title = vp.get_Parameter(BuiltInParameter.VIEWPORT_VIEW_NAME) # Parametro "Title on Sheet"
            title = p_title.AsString() if p_title and p_title.AsString() else ""
            
            # Se titulo na folha estiver vazio, pega nome da vista
            if not title:
                v = doc.GetElement(view_id)
                title = v.Name
            
            return number, title
    return "-", "DESENHO"

# --- 3. CLASSES ---

class ViewAnalysis:
    def __init__(self, view):
        self.view = view
        self.width = 0.0
        self.height = 0.0
        self.calculate_size()

    def calculate_size(self):
        bbox = self.view.get_BoundingBox(None)
        if bbox:
            w = (bbox.Max.X - bbox.Min.X)
            h = (bbox.Max.Y - bbox.Min.Y)
            scale = self.view.Scale
            self.width = w / scale
            self.height = h / scale
        else:
            self.width = 0.5
            self.height = 0.5

class ViewGroup:
    def __init__(self, name_base):
        self.name = name_base
        self.views = {'Planta': None, 'Corte': None, 'Elevacao': None}
        self.total_width = 0.0
        self.total_height = 0.0
    
    def add_view(self, view, type_key):
        self.views[type_key] = view

    def calculate_dimensions(self):
        ve = ViewAnalysis(self.views['Elevacao']) if self.views['Elevacao'] else None
        vc = ViewAnalysis(self.views['Corte']) if self.views['Corte'] else None
        vp = ViewAnalysis(self.views['Planta']) if self.views['Planta'] else None
        
        w_elev = ve.width if ve else 0
        w_corte = vc.width if vc else 0
        gap_x = (CFG_GAP_INT_X * MM_TO_FT) if (ve and vc) else 0
        self.total_width = w_elev + gap_x + w_corte
        
        h_elev = ve.height if ve else 0
        h_planta = vp.height if vp else 0
        gap_y = (CFG_GAP_INT_Y * MM_TO_FT) if (ve and vp) else 0
        self.total_height = h_elev + gap_y + h_planta
        
        return ve, vc, vp

class SheetEngine:
    def __init__(self, titleblock_symbol):
        self.tb_symbol = titleblock_symbol
        self.sheet_num_start = 101
        
        # Dimensoes da Folha
        p_w = titleblock_symbol.LookupParameter("Sheet Width") or titleblock_symbol.LookupParameter("Largura da folha")
        p_h = titleblock_symbol.LookupParameter("Sheet Height") or titleblock_symbol.LookupParameter("Altura da folha")
        
        self.sheet_w = p_w.AsDouble() if p_w else 2.75
        self.sheet_h = p_h.AsDouble() if p_h else 1.95
        
        # DEFINICAO DA AREA UTIL (Com Margem Direita de 145mm)
        self.min_x = CFG_MARGIN_LEFT * MM_TO_FT
        self.max_x = self.sheet_w - (CFG_MARGIN_RIGHT * MM_TO_FT) 
        
        self.max_y = self.sheet_h - (CFG_MARGIN_TOP * MM_TO_FT)
        self.min_y = CFG_MARGIN_BOTTOM * MM_TO_FT
        
        self.cursor_x = self.min_x
        self.cursor_y = self.max_y
        self.row_max_h = 0.0
        self.current_sheet = None
        self.current_drawing_list = [] 

    def create_sheet(self):
        # Salva lista da folha anterior
        if self.current_sheet and self.current_drawing_list:
            # Ordena a lista pelo numero antes de salvar (1, 2, 3...)
            # A lista ja deve estar ordenada pela ordem de insercao, mas garante:
            full_text = "\n".join(self.current_drawing_list)
            set_sheet_index(self.current_sheet, full_text)
        
        # Reseta
        self.current_drawing_list = []
        
        s_num = "A-{:03d}".format(self.sheet_num_start)
        self.sheet_num_start += 1
        try:
            self.current_sheet = ViewSheet.Create(doc, self.tb_symbol.Id)
            self.current_sheet.Name = "Automatico NnBim"
            self.current_sheet.SheetNumber = s_num
        except:
            s_num = "A-{:03d}".format(self.sheet_num_start + 50)
            try:
                self.current_sheet = ViewSheet.Create(doc, self.tb_symbol.Id)
                self.current_sheet.SheetNumber = s_num
            except: pass
            
        self.cursor_x = self.min_x
        self.cursor_y = self.max_y
        self.row_max_h = 0.0
        return self.current_sheet

    def finalize(self):
        if self.current_sheet and self.current_drawing_list:
            full_text = "\n".join(self.current_drawing_list)
            set_sheet_index(self.current_sheet, full_text)

    def _place_and_record(self, view, center):
        """Coloca a vista e grava no indice."""
        if not view: return
        try:
            Viewport.Create(doc, self.current_sheet.Id, view.Id, center)
            # Pega dados para o Indice
            num, title = get_viewport_info(doc, self.current_sheet.Id, view.Id)
            entry = "{} - {}".format(num, title.upper()) # CAIXA ALTA
            self.current_drawing_list.append(entry)
        except Exception as e:
            pass

    def place_views_ordered(self, group, start_x, start_y, ve, vc, vp):
        """Desenha na ordem: Planta (#1) -> Elevacao (#2) -> Corte (#3)."""
        
        # 1. CALCULO DE POSICOES VISUAIS
        coords = {}
        
        # Elevacao (Topo Esq)
        if ve:
            cx = start_x + (ve.width / 2)
            cy = start_y - (ve.height / 2)
            coords['Elevacao'] = XYZ(cx, cy, 0)

        # Corte (Topo Dir)
        if vc:
            offset_x = (ve.width if ve else 0) + (CFG_GAP_INT_X * MM_TO_FT)
            cx = start_x + offset_x + (vc.width / 2)
            cy = start_y - (vc.height / 2)
            coords['Corte'] = XYZ(cx, cy, 0)

        # Planta (Baixo Esq)
        if vp:
            cx = start_x + (vp.width / 2)
            offset_y = (ve.height if ve else 0) + (CFG_GAP_INT_Y * MM_TO_FT)
            cy = start_y - offset_y - (vp.height / 2)
            coords['Planta'] = XYZ(cx, cy, 0)

        # 2. ORDEM DE CRIACAO (Para definir numeracao)
        if 'Planta' in coords:
            self._place_and_record(group.views['Planta'], coords['Planta'])
            
        if 'Elevacao' in coords:
            self._place_and_record(group.views['Elevacao'], coords['Elevacao'])
            
        if 'Corte' in coords:
            self._place_and_record(group.views['Corte'], coords['Corte'])

    def process_grid(self, groups):
        if not self.current_sheet: self.create_sheet()
        
        for grp in groups:
            ve, vc, vp = grp.calculate_dimensions()
            
            # Verifica Largura (Respeitando a Margem 145mm)
            if (self.cursor_x + grp.total_width) > self.max_x:
                self.cursor_x = self.min_x
                self.cursor_y -= (self.row_max_h + (CFG_GAP_GRID_Y * MM_TO_FT))
                self.row_max_h = 0.0
            
            # Verifica Altura
            if (self.cursor_y - grp.total_height) < self.min_y:
                self.create_sheet()
            
            self.place_views_ordered(grp, self.cursor_x, self.cursor_y, ve, vc, vp)
            
            self.cursor_x += grp.total_width + (CFG_GAP_GRID_X * MM_TO_FT)
            if grp.total_height > self.row_max_h:
                self.row_max_h = grp.total_height
        
        self.finalize()

    def process_centered(self, groups):
        for grp in groups:
            sheet = self.create_sheet()
            sheet.Name = grp.name 
            
            ve, vc, vp = grp.calculate_dimensions()
            
            center_x = self.min_x + ((self.max_x - self.min_x) / 2)
            center_y = self.min_y + ((self.max_y - self.min_y) / 2)
            
            start_x = center_x - (grp.total_width / 2)
            start_y = center_y + (grp.total_height / 2)
            
            self.place_views_ordered(grp, start_x, start_y, ve, vc, vp)
        
        self.finalize()

# --- 4. MAIN ---

def main():
    sel_views = forms.select_views(title="Selecione Vistas", use_selection=True)
    if not sel_views: return

    tblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    if not tblocks: forms.alert("Sem Carimbos.", exitscript=True)
    
    dict_tb = {}
    for t in tblocks:
        try:
            full_name = "{} : {}".format(t.FamilyName, get_element_name(t))
            dict_tb[full_name] = t
        except: continue
    
    selected_tb = forms.SelectFromList.show(sorted(dict_tb.keys()), title="Escolha o Carimbo", multiselect=False)
    if not selected_tb: return
    tb_symbol = dict_tb[selected_tb]

    ops = {'Modo GRID (Varios Detalhes)': 'GRID', 'Modo CENTRALIZADO (Executivo)': 'CENTER'}
    res_mode = forms.CommandSwitchWindow.show(sorted(ops.keys()), message="Escolha o Modo")
    if not res_mode: return
    mode = ops[res_mode]

    groups = {} 
    pattern = re.compile(r"(.+)[_ -](Planta|Corte|Elevacao|Elev|Section|Plan)", re.IGNORECASE)
    
    for v in sel_views:
        if is_view_placed(v): continue
        match = pattern.search(v.Name)
        if match:
            base = match.group(1).strip()
            suf = match.group(2).lower()
            if base not in groups: groups[base] = ViewGroup(base)
            if "planta" in suf or "plan" in suf: groups[base].add_view(v, 'Planta')
            elif "corte" in suf or "section" in suf: groups[base].add_view(v, 'Corte')
            elif "elev" in suf: groups[base].add_view(v, 'Elevacao')

    if not groups:
        forms.alert("Nenhum grupo encontrado.")
        return

    with revit.Transaction("NnBim V6.1 Layout"):
        engine = SheetEngine(tb_symbol)
        sorted_groups = [groups[k] for k in sorted(groups.keys())]
        
        if mode == "GRID":
            engine.process_grid(sorted_groups)
        else:
            engine.process_centered(sorted_groups)

    forms.alert("Sucesso! Verifique se o parametro 'Nn_IndiceDesenhos' foi preenchido.")

if __name__ == '__main__':
    main()