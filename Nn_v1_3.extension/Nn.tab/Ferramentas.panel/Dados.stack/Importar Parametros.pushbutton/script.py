# -*- coding: utf-8 -*-
"""
NnBim: IMPORTADOR DE PARAMETROS EM MASSA
DESCRICAO:
Permite a importacao de dezenas de parametros compartilhados
com associacao inteligente (Lote Unico ou Individual).

COMO USAR:
1. Tenha um arquivo de Parametros Compartilhados (.txt) carregado
   em Gerenciar > Parametros Compartilhados.
2. Selecione os parametros desejados.
3. Escolha o modo de associacao de categorias.
4. Defina Instancia ou Tipo e o grupo de exibicao.
5. Clique em Executar.
"""
__title__ = 'Associar\nParametros'
__author__ = 'NnBim Dev'

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import *

doc = revit.doc
app = doc.Application
output = script.get_output()


def main():
    # 1. ACESSAR O ARQUIVO DE PARAMETROS COMPARTILHADOS (TXT)
    spf = app.OpenSharedParameterFile()
    if not spf:
        forms.alert(
            "Nenhum arquivo de Parametros Compartilhados (.txt) encontrado.\n"
            "Va em Gerenciar > Parametros Compartilhados e carregue um arquivo primeiro.",
            title="NnBim",
            exitscript=True
        )

    # 2. COLETAR E SELECIONAR OS PARAMETROS
    sp_dict = {}
    for group in spf.Groups:
        for defi in group.Definitions:
            nome_exibicao = "[{}] {}".format(group.Name, defi.Name)
            sp_dict[nome_exibicao] = defi

    if not sp_dict:
        forms.alert("O arquivo de parametros compartilhados esta vazio.", exitscript=True)

    sp_selecionados = forms.SelectFromList.show(
        sorted(sp_dict.keys()),
        title="NnBim: 1. Selecione os Parametros para Importar",
        multiselect=True,
        button_name="Avancar >"
    )
    if not sp_selecionados:
        script.exit()

    # 3. PERGUNTAR O MODO DE ASSOCIACAO
    modo_categoria = forms.CommandSwitchWindow.show(
        ["Mesmas Categorias para Todos", "Categorias Individuais por Parametro"],
        message="Como deseja associar as categorias aos parametros?"
    )
    if not modo_categoria:
        script.exit()

    # 4. MAPEAMENTO DE CATEGORIAS
    categorias = doc.Settings.Categories
    cat_dict = {c.Name: c for c in categorias if c.AllowsBoundParameters}

    param_cat_map = {}  # {nome_param: [categorias]}

    if modo_categoria == "Mesmas Categorias para Todos":
        cat_selecionadas = forms.SelectFromList.show(
            sorted(cat_dict.keys()),
            title="NnBim: 2. Selecione as Categorias para TODOS os parametros",
            multiselect=True,
            button_name="Avancar >"
        )
        if not cat_selecionadas:
            script.exit()

        for sp in sp_selecionados:
            param_cat_map[sp] = cat_selecionadas

    else:
        # Modo Individual: abre a janela para CADA parametro
        for sp in sp_selecionados:
            cat_selecionadas = forms.SelectFromList.show(
                sorted(cat_dict.keys()),
                title="Categorias para o parametro:\n{}".format(sp),
                multiselect=True,
                button_name="Confirmar e Proximo >"
            )
            if cat_selecionadas:
                param_cat_map[sp] = cat_selecionadas

    if not param_cat_map:
        script.exit()

    # 5. DEFINIR INSTANCIA OU TIPO
    tipo_associacao = forms.CommandSwitchWindow.show(
        ["Instancia", "Tipo"],
        message="Estes parametros serao criados como Instancia ou Tipo?"
    )
    if not tipo_associacao:
        script.exit()
    is_instance = (tipo_associacao == "Instancia")

    # 6. DEFINIR O GRUPO DE EXIBICAO
    grupos_opcoes = {
        "Dados de Identidade (Identity Data)": BuiltInParameterGroup.PG_IDENTITY_DATA,
        "Texto (Text)":                        BuiltInParameterGroup.PG_TEXT,
        "Dados (Data)":                        BuiltInParameterGroup.PG_DATA,
        "Geral (General)":                     BuiltInParameterGroup.PG_GENERAL,
        "Faseamento (Phasing)":                BuiltInParameterGroup.PG_PHASING,
        "Visibilidade (Visibility)":           BuiltInParameterGroup.PG_VISIBILITY
    }

    grupo_selecionado = forms.SelectFromList.show(
        sorted(grupos_opcoes.keys()),
        title="NnBim: Escolha o Grupo (Aba de Propriedades)",
        multiselect=False,
        button_name="EXECUTAR ASSOCIACAO"
    )
    if not grupo_selecionado:
        script.exit()

    bip_group = grupos_opcoes[
        grupo_selecionado[0] if isinstance(grupo_selecionado, list) else grupo_selecionado
    ]

    # 7. EXECUTAR A INJECAO
    sucessos = 0
    falhas   = 0

    with revit.Transaction("NnBim: Importacao de Parametros"):
        for sp_name, cats_escolhidas in param_cat_map.items():
            definition = sp_dict[sp_name]

            cat_set = app.Create.NewCategorySet()
            for c_name in cats_escolhidas:
                cat_set.Insert(cat_dict[c_name])

            binding = (
                app.Create.NewInstanceBinding(cat_set)
                if is_instance else
                app.Create.NewTypeBinding(cat_set)
            )

            try:
                inseriu = doc.ParameterBindings.Insert(definition, binding, bip_group)
                if not inseriu:
                    inseriu = doc.ParameterBindings.ReInsert(definition, binding, bip_group)

                if inseriu:
                    sucessos += 1
                else:
                    falhas += 1
            except Exception as e:
                print("Erro no parametro {}: {}".format(sp_name, e))
                falhas += 1

    # 8. RELATORIO
    forms.toast("Associacao NnBim Concluida!")

    output.print_md("### NnBim: Relatorio de Parametros")
    output.print_md("**Modo de Categoria:** {}".format(modo_categoria))
    output.print_md("**Parametros Associados com Sucesso:** {}".format(sucessos))

    if falhas > 0:
        output.print_md(
            "*Aviso: {} parametros falharam (podem ser do sistema ou estarem bloqueados).*".format(falhas)
        )


if __name__ == '__main__':
    main()