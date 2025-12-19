# -*- coding: utf-8 -*-
"""
Nome: Montar Prancha
Descricao: Diagramacao Automatica V5.11 (Revit 2025 Verified)
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

# --- 1. CONFIGURACOES & CALIBRACAO (Milimetros) ---

# [GAPS INTERNOS DO GRUPO]
CFG_GAP_INT_X = 25  # Entre Elevacao e Corte
CFG_GAP_INT_Y = 45  # Entre Elevacao e Planta

# [GAPS DO GRID] (Espacos entre os modulos)
CFG_GAP_GRID_X = 20 # Corredor vertical
CFG_GAP_GRID_Y = 55 # Corredor horizontal

# [MARGENS DA FOLHA]
CFG_MARGIN_LEFT = 25 # Margem esquerda (Encadernacao)
CFG_MARGIN_TOP = 20  # Margem superior
CFG_MARGIN_RIGHT = 15 # Margem direita

# [ZONA DE PROTECAO DA TABELA]
# Altura livre reservada na parte de baixo da folha para tabelas inseridas manualmente
CFG_TABLE_ZONE_HEIGHT = 100 

# Fator de Conversao (Mm -> Feet)
MM_TO_FT = 0.00328084

# --- 2. FUNCOES BLINDADAS (Para Revit 2025 e anteriores) ---

def is_view_placed(view):
    """
    Verifica se a vista ja esta em folha.
    Usa estrategia dupla para evitar erros do IronPython na versao 2025.
    """
    # TENTATIVA 1: Metodo Moderno (API 2023+)
    try:
        if view.ViewSheetId != ElementId.InvalidElementId:
            return True
    except AttributeError:
        # TENTATIVA 2: Fallback Classico (Ler Parametro)
        # Se o Python falhar ao ler a propriedade nova, lemos o parametro de texto.
        try:
            param = view.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
            if param:
                val = param.AsString()
                # Se tem valor e nao sao tracos, esta colocada
                if val and val != "---":
                    return True
        except:
            pass # Se tudo falhar, assume que nao esta colocada
            
    return False

def get_element_name(element):
    """Le o nome do elemento de forma segura."""
    return Element.Name.GetValue(element)

# --- 3. CLASSES DE LOGICA ---

class ViewAnalysis:
    """Analisa dimensoes reais da vista no papel."""
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
    """O Modulo (Conjunto de Vistas: Planta+Corte+Elev)."""
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
    """Motor de Criacao e Posicionamento (Tetris)."""
    def __init__(self, titleblock_symbol):
        self.tb_symbol = titleblock_symbol
        self.sheet_num_start = 101
        
        # Leitura segura de dimensoes da folha
        p_w = titleblock_symbol.LookupParameter("Sheet Width") or titleblock_symbol.LookupParameter("Largura da folha")
        p_h = titleblock_symbol.LookupParameter("Sheet Height") or titleblock_symbol.LookupParameter("Altura da folha")
        
        self.sheet_w = p_w.AsDouble() if p_w else 2.75
        self.sheet_h = p_h.AsDouble() if p_h else 1.95
        
        # Define Area Util (Canvas)
        self.min_x = CFG_MARGIN_LEFT * MM_TO_FT
        self.max_x = self.sheet_w - (CFG_MARGIN_RIGHT * MM_TO_FT)
        
        self.max_y = self.sheet_h - (CFG_MARGIN_TOP * MM_TO_FT)
        self.min_y = (CFG_TABLE_ZONE_HEIGHT * MM_TO_FT) + (20 * MM_TO_FT)
        
        self.cursor_x = self.min_x
        self.cursor_y = self.max_y
        self.row_max_h = 0.0
        self.current_sheet = None

    def create_sheet(self):
        s_num = "A-{:03d}".format(self.sheet_num_start)
        self.sheet_num_start += 1
        try:
            self.current_sheet = ViewSheet.Create(doc, self.tb_symbol.Id)
            self.current_sheet.Name = "Automatico NnBim"
            self.current_sheet.SheetNumber = s_num
        except:
            # Tenta pular numeracao se der erro
            s_num = "A-{:03d}".format(self.sheet_num_start + 50)
            try:
                self.current_sheet = ViewSheet.Create(doc, self.tb_symbol.Id)
                self.current_sheet.SheetNumber = s_num
            except: pass
            
        self.cursor_x = self.min_x
        self.cursor_y = self.max_y
        self.row_max_h = 0.0
        return self.current_sheet

    def place_views_generic(self, sheet, group, start_x, start_y, ve, vc, vp):
        # 1. Elevacao (Mestre)
        if ve:
            cx = start_x + (ve.width / 2)
            cy = start_y - (ve.height / 2)
            self._safe_create_viewport(sheet, group.views['Elevacao'], XYZ(cx, cy, 0))

        # 2. Corte (Direita)
        if vc:
            offset_x = (ve.width if ve else 0) + (CFG_GAP_INT_X * MM_TO_FT)
            cx = start_x + offset_x + (vc.width / 2)
            cy = start_y - (vc.height / 2)
            self._safe_create_viewport(sheet, group.views['Corte'], XYZ(cx, cy, 0))

        # 3. Planta (Abaixo)
        if vp:
            cx = start_x + (vp.width / 2)
            offset_y = (ve.height if ve else 0) + (CFG_GAP_INT_Y * MM_TO_FT)
            cy = start_y - offset_y - (vp.height / 2)
            self._safe_create_viewport(sheet, group.views['Planta'], XYZ(cx, cy, 0))

    def _safe_create_viewport(self, sheet, view, center):
        try:
            Viewport.Create(doc, sheet.Id, view.Id, center)
        except: pass

    def process_grid(self, groups):
        if not self.current_sheet: self.create_sheet()
        
        for grp in groups:
            ve, vc, vp = grp.calculate_dimensions()
            
            # Checa Largura (Quebra de Linha)
            if (self.cursor_x + grp.total_width) > self.max_x:
                self.cursor_x = self.min_x
                self.cursor_y -= (self.row_max_h + (CFG_GAP_GRID_Y * MM_TO_FT))
                self.row_max_h = 0.0
            
            # Checa Altura (Quebra de Folha)
            if (self.cursor_y - grp.total_height) < self.min_y:
                self.create_sheet()
            
            # Insere
            self.place_views_generic(self.current_sheet, grp, self.cursor_x, self.cursor_y, ve, vc, vp)
            
            # Avanca Cursor
            self.cursor_x += grp.total_width + (CFG_GAP_GRID_X * MM_TO_FT)
            if grp.total_height > self.row_max_h:
                self.row_max_h = grp.total_height

    def process_centered(self, groups):
        for grp in groups:
            sheet = self.create_sheet()
            sheet.Name = grp.name 
            
            ve, vc, vp = grp.calculate_dimensions()
            
            center_sheet_x = self.min_x + ((self.max_x - self.min_x) / 2)
            center_sheet_y = self.min_y + ((self.max_y - self.min_y) / 2)
            
            start_x = center_sheet_x - (grp.total_width / 2)
            start_y = center_sheet_y + (grp.total_height / 2)
            
            self.place_views_generic(sheet, grp, start_x, start_y, ve, vc, vp)

# --- 4. MAIN ---

def main():
    # 1. Seleciona Vistas
    sel_views = forms.select_views(title="Selecione Vistas", use_selection=True)
    if not sel_views: return

    # 2. Seleciona Carimbo
    tblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    if not tblocks: forms.alert("Sem Carimbos carregados.", exitscript=True)
    
    dict_tb = {}
    for t in tblocks:
        try:
            fam_name = t.FamilyName
            type_name = get_element_name(t) # Uso da funcao segura
            full_name = "{} : {}".format(fam_name, type_name)
            dict_tb[full_name] = t
        except: continue
    
    if not dict_tb:
        forms.alert("Erro ao ler nomes dos Carimbos.", exitscript=True)

    selected_tb_name = forms.SelectFromList.show(
        sorted(dict_tb.keys()),
        title="1. Escolha o Carimbo",
        multiselect=False
    )
    if not selected_tb_name: return
    tb_symbol = dict_tb[selected_tb_name]

    # 3. Seleciona Modo
    ops = {'Modo GRID (Varios Detalhes)': 'GRID', 'Modo CENTRALIZADO (Executivo)': 'CENTER'}
    res_mode = forms.CommandSwitchWindow.show(
        sorted(ops.keys()),
        message="2. Escolha o Modo de Diagramacao"
    )
    if not res_mode: return
    mode = ops[res_mode]

    # 4. Agrupamento Inteligente
    groups = {} 
    pattern = re.compile(r"(.+)[_ -](Planta|Corte|Elevacao|Elev|Section|Plan)", re.IGNORECASE)
    
    for v in sel_views:
        # AQUI ESTA A CORRECAO:
        if is_view_placed(v): 
            # Pula vista se ja estiver em folha
            continue
            
        match = pattern.search(v.Name)
        if match:
            base = match.group(1).strip()
            suf = match.group(2).lower()
            if base not in groups: groups[base] = ViewGroup(base)
            
            if "planta" in suf or "plan" in suf: groups[base].add_view(v, 'Planta')
            elif "corte" in suf or "section" in suf: groups[base].add_view(v, 'Corte')
            elif "elev" in suf: groups[base].add_view(v, 'Elevacao')

    if not groups:
        forms.alert("Nenhum grupo identificado ou vistas ja estao em folhas.")
        return

    # 5. Execucao
    with revit.Transaction("NnBim V5.11 Layout"):
        engine = SheetEngine(tb_symbol)
        sorted_groups = [groups[k] for k in sorted(groups.keys())]
        
        if mode == "GRID":
            engine.process_grid(sorted_groups)
        else:
            engine.process_centered(sorted_groups)

    forms.alert("Sucesso! Pranchas geradas no modo {}.".format(mode))

if __name__ == '__main__':
    main()