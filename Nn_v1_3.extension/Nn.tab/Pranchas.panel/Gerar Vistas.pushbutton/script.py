# -*- coding: utf-8 -*-
"""
NnBim: MODELO SLIM
DESCRICAO:
Ferramenta de simplificacao de modelos.
Remove vistas, tabelas, templates, folhas e tipos nao usados.

COMO USAR:
1. Execute e escolha o estilo de visualizacao das listas.
2. Marque o que deseja MANTER em cada categoria.
3. Confirme a remocao de vinculos se necessario.
"""
__title__ = 'Modelo\nSlim'
__author__ = 'Nn_Dev'

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = revit.doc


def selecionar_flexivel(titulo, modo, tipo_filtro):
    coletor = FilteredElementCollector(doc).OfClass(View).ToElements()
    opcoes  = {}

    for c in coletor:
        if tipo_filtro == "Templates":
            if not c.IsTemplate: continue
        elif tipo_filtro == "Tabelas":
            if c.ViewType != ViewType.Schedule or c.IsTemplate: continue
        elif tipo_filtro == "Vistas":
            if c.IsTemplate: continue
            if c.ViewType == ViewType.Schedule: continue
            if c.ViewType == ViewType.DrawingSheet: continue
            if c.ViewType in [ViewType.ProjectBrowser, ViewType.SystemBrowser, ViewType.Internal]: continue

        if c.Id == doc.ActiveView.Id: continue
        if not c.CanBePrinted and tipo_filtro == "Vistas": continue

        try:
            if modo == "Familia-Tipo" and hasattr(c, "GetTypeId"):
                v_type = doc.GetElement(c.GetTypeId())
                fam    = v_type.FamilyName if v_type else "Sistema"
                tipo   = v_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                label  = "[{} : {}] {}".format(fam, tipo, c.Name)
            else:
                label = str(c.Name)
        except:
            label = str(c.Name)

        opcoes[label] = c.Id

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
    todos_tipos = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
    nao_usados  = []
    for t in todos_tipos:
        try:
            if not t.Category: continue
            regra  = ParameterFilterRuleFactory.CreateEqualsRule(ElementId(BuiltInParameter.ELEM_TYPE_PARAM), t.Id)
            filtro = ElementParameterFilter(regra)
            if FilteredElementCollector(doc).WherePasses(filtro).WhereElementIsNotElementType().GetElementCount() == 0:
                nao_usados.append(t.Id)
        except:
            continue
    return nao_usados


def main():
    op_vis   = ["Familia-Tipo (Agrupado)", "Ordem Alfabetica (A-Z)"]
    modo_vis = forms.ask_for_one_item(op_vis, title="Configurar Visualizacao", prompt="Escolha o estilo das listas:")
    if not modo_vis: return

    vistas_manter    = selecionar_flexivel("VISTAS: Marque as que FICAM", modo_vis, "Vistas")
    tabelas_manter   = selecionar_flexivel("TABELAS: Marque as que FICAM", modo_vis, "Tabelas")
    templates_manter = selecionar_flexivel("VIEW TEMPLATES: Marque os que FICAM", modo_vis, "Templates")

    folhas_base  = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    op_folhas    = {str(f.Name): f.Id for f in folhas_base}
    sel_folhas   = forms.SelectFromList.show(sorted(op_folhas.keys()), title="PRANCHAS: Marque as que FICAM", multiselect=True)
    folhas_manter = [op_folhas[n] for n in sel_folhas] if sel_folhas else []

    todos_manter = set(vistas_manter + tabelas_manter + templates_manter + folhas_manter)
    todos_manter.add(doc.ActiveView.Id)

    ids_purge     = obter_tipos_nao_usados()
    remover_links = forms.alert("Remover Vinculos (Revit e CAD)?", yes=True, no=True)

    with revit.Transaction("Nn_Modelo_Slim_V2.5"):
        total = 0

        # Passo A: Deletar Folhas
        todas_folhas = FilteredElementCollector(doc).OfClass(ViewSheet).ToElementIds()
        for f_id in todas_folhas:
            if f_id not in todos_manter:
                try:
                    doc.Delete(f_id)
                    total += 1
                except:
                    pass

        # Passo B: Deletar Vistas, Tabelas e Templates
        # CORRECAO: parenteses adicionados para corrigir precedencia de operadores
        todas_vistas = FilteredElementCollector(doc).OfClass(View).ToElementIds()
        for v_id in todas_vistas:
            if v_id not in todos_manter:
                el = doc.GetElement(v_id)
                if el and (el.CanBePrinted or isinstance(el, ViewSchedule)):
                    try:
                        doc.Delete(v_id)
                        total += 1
                    except:
                        pass

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
                except:
                    pass

    forms.alert("Sucesso! {} itens removidos do modelo.".format(total))


if __name__ == '__main__':
    main()