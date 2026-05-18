# -*- coding: utf-8 -*-
"""
NnBim: SCANNER RAIO-X GLOBAL
DESCRICAO:
Extrai todos os parametros de um elemento, incluindo:
- Parametros Nativos (Built-in)
- Parametros de Projeto
- Parametros Compartilhados (Shared Parameters)

COMO USAR:
1. Selecione UM elemento no modelo, ou
2. Execute sem selecao para escanear Informacoes do Projeto ou Vista Atual.
"""
__title__ = 'Raio-X\nGlobal'
__author__ = 'NnBim Dev'

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()


def print_parameters(element, title):
    output.print_md("## " + title)

    all_params = element.GetOrderedParameters()
    if not all_params:
        print("Nenhum parametro encontrado.")
        return

    sorted_params = sorted(all_params, key=lambda p: p.Definition.Name)

    for p in sorted_params:
        try:
            nome      = p.Definition.Name
            is_shared = " (COMPARTILHADO)" if p.IsShared else ""

            valor = "N/A"
            if p.StorageType == StorageType.String:
                valor = p.AsString()
            else:
                valor = p.AsValueString()
            if valor is None or valor == "":
                valor = "[Vazio]"

            bip = p.Definition.BuiltInParameter
            if bip != BuiltInParameter.INVALID:
                cod = str(bip)
            else:
                cod = "GUID: " + str(p.GUID) if p.IsShared else "Customizado"

            print("Nome: **{}{}** | Valor: `{}` | Codigo: `{}`".format(
                nome, is_shared, valor, cod
            ))
        except Exception:
            continue


def main():
    selection = uidoc.Selection.GetElementIds()

    el       = None
    el_type  = None
    contexto = ""

    if selection:
        if len(selection) > 1:
            forms.alert("Selecione apenas UM elemento para um Raio-X preciso.")
            return

        el       = doc.GetElement(list(selection)[0])
        contexto = "Elemento Selecionado"

        if el.GetTypeId() != ElementId.InvalidElementId:
            el_type = doc.GetElement(el.GetTypeId())

    else:
        opcoes = {
            "Informacoes do Projeto": doc.ProjectInformation,
            "Vista / Prancha Atual":  doc.ActiveView
        }

        escolha = forms.SelectFromList.show(
            sorted(opcoes.keys()),
            title="O que deseja escanear?",
            multiselect=False
        )
        if not escolha:
            return

        el       = opcoes[escolha]
        contexto = escolha

    # Cabecalho
    output.print_md("# Raio-X NnBim: Scanner Global")
    output.print_md("**Contexto:** {}".format(contexto))
    output.print_md("---")

    print("- Categoria : {}".format(el.Category.Name if el.Category else "N/A"))
    print("- Nome      : {}".format(el.Name))
    print("- ID        : {}".format(el.Id))

    output.print_md("---")
    print_parameters(el, "PARAMETROS DE INSTANCIA")

    if el_type:
        output.print_md("---")
        print_parameters(el_type, "PARAMETROS DE TIPO")

    output.print_md("---")
    output.print_md("**Scanner concluido. De Ctrl+A, Ctrl+C e mande no chat.**")


if __name__ == '__main__':
    main()