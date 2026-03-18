# -*- coding: utf-8 -*-
"""
NnBim: IMPORTADOR DE PARÂMETROS EM MASSA
DESCRICAO:
Permite a importação de dezenas de parâmetros compartilhados
com associação inteligente (Lote único ou Individual).
"""

__title__ = 'Associar\nParâmetros'
__author__ = 'NnBim Dev'

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import *

doc = revit.doc
app = doc.Application
output = script.get_output()

def main():
    # 1. ACESSAR O ARQUIVO DE PARÂMETROS COMPARTILHADOS (TXT)
    spf = app.OpenSharedParameterFile()
    if not spf:
        forms.alert("Nenhum arquivo de Parâmetros Compartilhados (.txt) encontrado.\nVá em Gerenciar > Parâmetros Compartilhados e carregue um arquivo primeiro.", title="NnBim", exitscript=True)

    # 2. COLETAR E SELECIONAR OS PARÂMETROS
    sp_dict = {}
    for group in spf.Groups:
        for defi in group.Definitions:
            nome_exibicao = "[{}] {}".format(group.Name, defi.Name)
            sp_dict[nome_exibicao] = defi

    if not sp_dict:
        forms.alert("O arquivo de parâmetros compartilhados está vazio.", exitscript=True)

    sp_selecionados = forms.SelectFromList.show(
        sorted(sp_dict.keys()),
        title="NnBim: 1. Selecione os Parâmetros para Importar",
        multiselect=True,
        button_name="Avançar >"
    )
    if not sp_selecionados: script.exit()

    # 3. PERGUNTAR O MODO DE ASSOCIAÇÃO
    modo_categoria = forms.CommandSwitchWindow.show(
        ["Mesmas Categorias para Todos", "Categorias Individuais por Parâmetro"],
        message="Como deseja associar as categorias aos parâmetros?"
    )
    if not modo_categoria: script.exit()

    # 4. MAPEAMENTO DE CATEGORIAS
    categorias = doc.Settings.Categories
    cat_dict = {c.Name: c for c in categorias if c.AllowsBoundParameters}
    
    param_cat_map = {} # Dicionário para guardar [Parâmetro : Categorias dele]

    if modo_categoria == "Mesmas Categorias para Todos":
        cat_selecionadas = forms.SelectFromList.show(
            sorted(cat_dict.keys()),
            title="NnBim: 2. Selecione as Categorias para TODOS os parâmetros",
            multiselect=True,
            button_name="Avançar >"
        )
        if not cat_selecionadas: script.exit()
        
        # Copia a mesma lista de categorias para todos os parâmetros
        for sp in sp_selecionados:
            param_cat_map[sp] = cat_selecionadas

    else:
        # Modo Individual: Abre a janela para CADA parâmetro
        for sp in sp_selecionados:
            cat_selecionadas = forms.SelectFromList.show(
                sorted(cat_dict.keys()),
                title="Categorias para o parâmetro:\n{}".format(sp),
                multiselect=True,
                button_name="Confirmar e Próximo >"
            )
            # Se o usuário cancelar a janela no meio do processo, pula o parâmetro
            if cat_selecionadas:
                param_cat_map[sp] = cat_selecionadas

    if not param_cat_map: 
        script.exit()

    # 5. DEFINIR INSTÂNCIA OU TIPO (Aplica a todos da rodada para agilizar)
    tipo_associacao = forms.CommandSwitchWindow.show(
        ["Instância", "Tipo"],
        message="Estes parâmetros serão criados como Instância ou Tipo?"
    )
    if not tipo_associacao: script.exit()
    is_instance = (tipo_associacao == "Instância")

    # 6. DEFINIR O GRUPO DE EXIBIÇÃO
    grupos_opcoes = {
        "Dados de Identidade (Identity Data)": BuiltInParameterGroup.PG_IDENTITY_DATA,
        "Texto (Text)": BuiltInParameterGroup.PG_TEXT,
        "Dados (Data)": BuiltInParameterGroup.PG_DATA,
        "Geral (General)": BuiltInParameterGroup.PG_GENERAL,
        "Faseamento (Phasing)": BuiltInParameterGroup.PG_PHASING,
        "Visibilidade (Visibility)": BuiltInParameterGroup.PG_VISIBILITY
    }

    grupo_selecionado = forms.SelectFromList.show(
        sorted(grupos_opcoes.keys()),
        title="NnBim: Escolha o Grupo (Aba de Propriedades)",
        multiselect=False,
        button_name="EXECUTAR ASSOCIAÇÃO"
    )
    if not grupo_selecionado: script.exit()

    bip_group = grupos_opcoes[grupo_selecionado[0] if isinstance(grupo_selecionado, list) else grupo_selecionado]

    # 7. EXECUTAR A INJEÇÃO (CÉREBRO DA ASSOCIAÇÃO DINÂMICA)
    sucessos = 0
    falhas = 0

    with revit.Transaction("NnBim: Importação de Parâmetros"):
        for sp_name, cats_escolhidas in param_cat_map.items():
            definition = sp_dict[sp_name]
            
            # Cria o pacote de categorias exclusivo deste parâmetro
            cat_set = app.Create.NewCategorySet()
            for c_name in cats_escolhidas:
                cat_set.Insert(cat_dict[c_name])

            # Cria a "Binding" (Cola do parâmetro)
            binding = app.Create.NewInstanceBinding(cat_set) if is_instance else app.Create.NewTypeBinding(cat_set)
            
            try:
                inseriu = doc.ParameterBindings.Insert(definition, binding, bip_group)
                if not inseriu:
                    inseriu = doc.ParameterBindings.ReInsert(definition, binding, bip_group)
                
                if inseriu:
                    sucessos += 1
                else:
                    falhas += 1
            except Exception as e:
                print("Erro no parâmetro {}: {}".format(sp_name, e))
                falhas += 1

    # 8. RELATÓRIO
    forms.toast("Associação NnBim Concluída!")
    
    output.print_md("### NnBim: Relatório de Parâmetros")
    output.print_md("**Modo de Categoria:** {}".format(modo_categoria))
    output.print_md("**Parâmetros Associados com Sucesso:** {}".format(sucessos))
    
    if falhas > 0:
        output.print_md("*Aviso: {} parâmetros falharam (podem ser do sistema ou estarem bloqueados).*".format(falhas))

if __name__ == '__main__':
    main()