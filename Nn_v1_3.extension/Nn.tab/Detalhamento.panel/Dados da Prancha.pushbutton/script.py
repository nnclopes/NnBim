# -*- coding: utf-8 -*-
"""
Nome: Atualizar Dados PRANCHA
Descricao: Indice (Com Numeros) + Colunas + Auditoria (V17.2)
Autor: NnBim Dev
"""

import clr
import re
from datetime import datetime

# Imports do Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

# Imports do pyRevit
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc
my_config = script.get_config()

# --- CONFIGURACAO PADRAO ---
LIMIT_LINES_DEFAULT = 8 

# --- 1. FUNCOES AUXILIARES ---

def natural_sort_key(s):
    # Ordena 1, 2, 10 corretamente
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split('([0-9]+)', s)]

def get_viewport_info(viewport):
    # --- 1. CAPTURA DO NUMERO (REFORÇADA) ---
    num = ""
    try:
        # Tentativa A: Pelo ID interno (Mais confiavel)
        p = viewport.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_DETAIL_NUMBER)
        if p: num = p.AsString()
        
        # Tentativa B: Se falhar, tenta pelo nome do parametro (PT/EN)
        if not num:
            p = viewport.LookupParameter("Detail Number")
            if not p: p = viewport.LookupParameter("Número do detalhe")
            if p: num = p.AsString()
            
    except: pass
    
    if num is None: num = ""

    # --- 2. CAPTURA DO TITULO ---
    title = ""
    try:
        view = doc.GetElement(viewport.ViewId)
        if view:
            # Tenta Titulo na Folha
            p_title = view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
            if p_title: title = p_title.AsString()
            
            # Se vazio, pega Nome da Vista
            if not title: title = view.Name
    except: pass
    
    if not title: title = "DESENHO S/ TÍTULO"
    
    return num, title

def get_all_text_parameters(doc):
    params = {}
    sample_sheet = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).FirstElement()
    if sample_sheet:
        for p in sample_sheet.Parameters:
            if p.IsReadOnly: continue 
            if p.StorageType == DB.StorageType.String:
                p_name = p.Definition.Name
                params[p_name] = "SHEET"

    proj_info = doc.ProjectInformation
    if proj_info:
        for p in proj_info.Parameters:
            if p.IsReadOnly: continue
            if p.StorageType == DB.StorageType.String:
                p_name = p.Definition.Name
                key_name = "[GLOBAL] " + p_name
                params[key_name] = "PROJECT"
    return params

def get_config_safe(key_name):
    try: return my_config.get_option(key_name)
    except: return None

def select_param_design(key_name, title_text, param_dict, default_filter=None, force_ask=False):
    saved_value = get_config_safe(key_name)
    param_keys = sorted(param_dict.keys())
    
    if saved_value and saved_value in param_keys and not force_ask:
        return saved_value
    
    sug_default = next((x for x in param_keys if default_filter and default_filter.lower() in x.lower()), None)
    
    selected = forms.SelectFromList.show(
        param_keys,
        title=title_text,
        button_name="Confirmar",
        multiselect=False,
        default=sug_default
    )
    
    if selected: my_config.set_option(key_name, selected)
    return selected

def check_and_fill_audit(sheets, filled_params):
    if not sheets: return
    sample = sheets[0]
    audit_list = [
        ("CLIENTE 👤", ["cliente", "client"]),
        ("PROJETO 🏗️", ["obra", "projeto", "project name"]),
        ("DESENHISTA ✏️", ["desenhado", "drawn"]),
        ("REVISOR 👀", ["verificado", "checked"]),
    ]
    updates = {}
    for friendly, keywords in audit_list:
        found_param = None
        for p in sample.Parameters:
            if p.IsReadOnly or p.Definition.Name in filled_params: continue
            if p.StorageType != DB.StorageType.String: continue
            if any(k in p.Definition.Name.lower() for k in keywords):
                found_param = p
                break
        if found_param:
            val = found_param.AsString()
            if not val or val.strip() == "":
                new_val = forms.ask_for_string(title="Auditoria", prompt="O campo '{}' está vazio. Preencher:".format(friendly))
                if new_val: updates[found_param.Definition.Name] = new_val.upper()
    return updates

# --- 2. MAIN ---

def main():
    force_reset = False
    if globals().get('__shiftclick__', False): force_reset = True
        
    # A. SELECAO
    selection = uidoc.Selection.GetElementIds()
    sheets = []
    if selection:
        for el_id in selection:
            el = doc.GetElement(el_id)
            if isinstance(el, DB.ViewSheet): sheets.append(el)
    
    if not sheets:
        sheets = forms.select_sheets(title='Selecione as Pranchas 📑', include_placeholder=False, button_name='Iniciar')
    if not sheets: return

    # DADOS AUTO
    count_total = len(sheets)
    str_total = "{:02d}".format(count_total)
    str_date = datetime.now().strftime("%d/%m/%Y")
    doc_title = doc.Title
    str_file = doc_title[:-4] if doc_title.lower().endswith(".rvt") else doc_title

    # MAPEAR
    all_params_map = get_all_text_parameters(doc)
    if not all_params_map:
        forms.alert("Erro: Sem parâmetros editáveis.")
        return

    # --- SETUP PERGUNTAS ---
    
    # 1. COLUNA 1
    target_list_1 = select_param_design('cfg_lista', "1/6. LISTA DE DESENHOS (Coluna 1):", all_params_map, "conteudo", force_reset)
    if not target_list_1: return

    # 2. COLUNA 2
    try:
        target_list_2 = select_param_design('cfg_lista_2', "2/6. LISTA DE DESENHOS (Coluna 2 - Opcional):", all_params_map, "conteudo", force_reset)
    except: target_list_2 = None 

    # Limite de Linhas
    limit_lines = get_config_safe('cfg_limit_lines')
    if not limit_lines or force_reset:
        res = forms.ask_for_string(
            default=str(LIMIT_LINES_DEFAULT),
            title="Limite de Linhas",
            prompt="Quantas linhas cabem na Coluna 1 antes de pular?"
        )
        limit_lines = int(res) if res and res.isdigit() else LIMIT_LINES_DEFAULT
        my_config.set_option('cfg_limit_lines', limit_lines)
    else:
        limit_lines = int(limit_lines)

    # 3. DISCIPLINA
    target_disc = select_param_design('cfg_disc', "3/6. DISCIPLINA:", all_params_map, "disciplina", force_reset)
    if not target_disc: return
    val_disc = forms.ask_for_string(default="ARQUITETURA", title="Disciplina", prompt="Nome da Disciplina:")
    if not val_disc: return

    # 4. ARQUIVO
    target_file = select_param_design('cfg_file', "4/6. NOME DO ARQUIVO:", all_params_map, "arquivo", force_reset)
    if not target_file: return
    
    # 5. TOTAL
    target_total = select_param_design('cfg_total', "5/6. TOTAL DE FOLHAS:", all_params_map, "total", force_reset)
    if not target_total: return

    # 6. DATA
    target_date = select_param_design('cfg_data', "6/6. DATA DE EMISSÃO:", all_params_map, "emiss", force_reset)
    if not target_date: return

    # AUDITORIA
    filled_list = [target_list_1, target_disc, target_file, target_total, target_date]
    if target_list_2: filled_list.append(target_list_2)
    clean_filled = [x.replace("[GLOBAL] ", "") for x in filled_list]
    extra_updates = check_and_fill_audit(sheets, clean_filled)

    # --- EXECUCAO ---
    processed_count = 0
    t = DB.Transaction(doc, "NnBim: Atualizar V17.2")
    t.Start()
    
    try:
        actions = [
            (target_disc, val_disc.upper(), all_params_map[target_disc]),
            (target_file, str_file, all_params_map[target_file]),
            (target_total, str_total, all_params_map[target_total]),
            (target_date, str_date, all_params_map[target_date])
        ]

        # 1. Globais
        for param_key, val, p_type in actions:
            if p_type == "PROJECT":
                real_name = param_key.replace("[GLOBAL] ", "")
                p_proj = doc.ProjectInformation.LookupParameter(real_name)
                if p_proj: p_proj.Set(val)

        # 2. Folhas
        for sheet in sheets:
            # GERA LISTA COM NUMERO
            vp_ids = sheet.GetAllViewports()
            drawing_list = []
            if vp_ids:
                temp_data = []
                for vid in vp_ids:
                    vp = doc.GetElement(vid)
                    if vp:
                        n, title = get_viewport_info(vp)
                        temp_data.append((n, title))
                
                # Ordena pelo numero
                temp_data.sort(key=lambda x: natural_sort_key(x[0]))
                
                for n, title in temp_data:
                    # Formata: "01 - TITULO" ou "TITULO" se sem numero
                    if n and n.strip() != "":
                        line = "{} - {}".format(n, title.upper())
                    else:
                        line = title.upper()
                    drawing_list.append(line)
            
            # --- QUEBRA DE COLUNAS ---
            text_col1 = ""
            text_col2 = ""
            
            if len(drawing_list) <= limit_lines:
                text_col1 = "\n".join(drawing_list)
                text_col2 = "" 
            else:
                part1 = drawing_list[:limit_lines]
                part2 = drawing_list[limit_lines:]
                text_col1 = "\n".join(part1)
                text_col2 = "\n".join(part2)

            # Grava
            p1 = sheet.LookupParameter(target_list_1)
            if p1: p1.Set(text_col1)
            
            if target_list_2:
                p2 = sheet.LookupParameter(target_list_2)
                if p2: p2.Set(text_col2)

            # Outros Params
            for param_key, val, p_type in actions:
                if p_type == "SHEET":
                    p = sheet.LookupParameter(param_key)
                    if p: p.Set(val)

            # Auditoria
            if extra_updates:
                for pname, pval in extra_updates.items():
                    p = sheet.LookupParameter(pname)
                    if p: p.Set(pval)

            processed_count += 1
            
        t.Commit()
        forms.alert("✅ Sucesso!\n{} pranchas atualizadas.".format(processed_count), title="Concluído")
        
    except Exception as e:
        t.RollBack()
        forms.alert("Erro: {}".format(e))

if __name__ == '__main__':
    main()