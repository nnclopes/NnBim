# -*- coding: utf-8 -*-
"""
NnBim Dev - Ferramenta de Simplificacao de Modelos
NOME: Modelo Slim
VERSAO: V2.5 (Correcao Definitiva de Delecao de Vistas)
AUTOR: Nn_Dev (Engenharia de Software NnBim)
"""

__title__ = 'Modelo\nSlim'
__author__ = 'NnBim Dev'

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = revit.doc

def selecionar_flexivel(titulo, modo, tipo_filtro):
    """Interface com filtros precisos para separar Vistas, Tabelas e Templates"""
    # Coletor base de todas as vistas
    coletor = FilteredElementCollector(doc).OfClass(View).ToElements()
    opcoes = {}
    
    for c in coletor:
        # 1. Filtro para VIEW TEMPLATES
        if tipo_filtro == "Templates":
            if not c.IsTemplate: continue
        
        # 2. Filtro para TABELAS
        elif tipo_filtro == "Tabelas":
            if c.ViewType != ViewType.Schedule or c.IsTemplate: continue
            
        # 3. Filtro para VISTAS (Plantas, Cortes, Fachadas, 3D)
        elif tipo_filtro == "Vistas":
            if c.IsTemplate: continue
            if c.ViewType == ViewType.Schedule: continue
            if c.ViewType == ViewType.DrawingSheet: continue # Folhas tem classe propria
            # Pular vistas internas do sistema
            if c.ViewType in [ViewType.ProjectBrowser, ViewType.SystemBrowser, ViewType.Internal]: continue
        
        # Ignorar vista ativa e checar se pode ser apagada
        if c.Id == doc.ActiveView.Id: continue
        if not c.CanBePrinted and tipo_filtro == "Vistas": continue 

        # --- FORMATACAO DO LABEL ---
        try:
            if modo == "Familia-Tipo" and hasattr(c, "GetTypeId"):
                v_type = doc.GetElement(c.GetTypeId())
                fam = v_type.FamilyName if v_type else "Sistema"
                tipo = v_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                label = "[{} : {}] {}".format(fam, tipo, c.Name)
            else:
                label = str(c.Name)
        except:
            label = str(c.Name)
        
        opcoes[label] = c.Id # Guardamos apenas o ID

    if not opcoes: return []

    escolhidos = forms.SelectFromList.show(
        sorted(opcoes.keys()),
        title=titulo,
        multiselect=True,
        button_name='Manter Selecionados',
        width=950
    )
    
    return [opcoes[n] for n in escolhidos] if escolhidos else []

def obter_tipos_nao_usados():
    """Identifica tipos sem nenhuma instancia no projeto"""
    todos_tipos = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
    nao_usados = []
    for t in todos_tipos:
        try:
            if not t.Category: continue
            regra = ParameterFilterRuleFactory.CreateEqualsRule(ElementId(BuiltInParameter.ELEM_TYPE_PARAM), t.Id)
            filtro = ElementParameterFilter(regra)
            if FilteredElementCollector(doc).WherePasses(filtro).WhereElementIsNotElementType().GetElementCount() == 0:
                nao_usados.append(t.Id)
        except: continue
    return nao_usados

def main():
    # --- CONFIGURACOES ---
    op_vis = ["Familia-Tipo (Agrupado)", "Ordem Alfabetica (A-Z)"]
    modo_vis = forms.ask_for_one_item(op_vis, title="Configurar Visualizacao", prompt="Escolha o estilo das listas:")
    if not modo_vis: return

    # --- 1. COLETAR IDs PARA MANTER ---
    vistas_manter = selecionar_flexivel("VISTAS: Marque as que FICAM", modo_vis, "Vistas")
    tabelas_manter = selecionar_flexivel("TABELAS: Marque as que FICAM", modo_vis, "Tabelas")
    templates_manter = selecionar_flexivel("VIEW TEMPLATES: Marque os que FICAM", modo_vis, "Templates")
    
    # Coleta de Folhas (ViewSheet e uma classe separada no Revit)
    folhas_base = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    op_folhas = {str(f.Name): f.Id for f in folhas_base}
    sel_folhas = forms.SelectFromList.show(sorted(op_folhas.keys()), title="PRANCHAS: Marque as que FICAM", multiselect=True)
    folhas_manter = [op_folhas[n] for n in sel_folhas] if sel_folhas else []

    # Unificar lista de preservacao
    todos_manter = set(vistas_manter + tabelas_manter + templates_manter + folhas_manter)
    todos_manter.add(doc.ActiveView.Id)

    # --- 2. PURGE E LINKS ---
    ids_purge = obter_tipos_nao_usados()
    remover_links = forms.alert("Remover Vinculos (Revit e CAD)?", yes=True, no=True)

    # --- 3. EXECUCAO DA LIMPEZA ---
    with revit.Transaction("Nn_Modelo_Slim_V2.5"):
        total = 0
        
        # Passo A: Deletar Folhas Primeiro (Isso ja limpa as vistas dentro delas)
        todas_folhas = FilteredElementCollector(doc).OfClass(ViewSheet).ToElementIds()
        for f_id in todas_folhas:
            if f_id not in todos_manter:
                try:
                    doc.Delete(f_id)
                    total += 1
                except: pass

        # Passo B: Deletar Vistas, Tabelas e Templates restantes
        todas_vistas = FilteredElementCollector(doc).OfClass(View).ToElementIds()
        for v_id in todas_vistas:
            if v_id not in todos_manter:
                # O GetElement checa se a vista ainda existe (pode ter sido deletada com a folha)
                el = doc.GetElement(v_id)
                if el and el.CanBePrinted or isinstance(el, ViewSchedule):
                    try:
                        doc.Delete(v_id)
                        total += 1
                    except: pass

        # Passo C: Links e Purge
        links_ids = []
        if remover_links:
            links_ids.extend(FilteredElementCollector(doc).OfClass(RevitLinkType).ToElementIds())
            links_ids.extend(FilteredElementCollector(doc).OfClass(ImportInstance).ToElementIds())
        
        for p_id in list(links_ids) + list(ids_purge):
            if doc.GetElement(p_id):
                try:
                    doc.Delete(p_id)
                    total += 1
                except: pass

    forms.alert("Sucesso! {} itens removidos do modelo.".format(total))

if __name__ == '__main__':
    main()

