# -*- coding: utf-8 -*-
"""
NnBim: Exportar Parametros Compartilhados
DESCRICAO:
Varre os parametros compartilhados da familia aberta e os exporta
para o arquivo .txt mapeado, reportando erros exatos da API.

COMO USAR:
1. Abra uma familia (.rfa) no Editor de Familias.
2. Tenha um arquivo de Parametros Compartilhados (.txt) mapeado
   em Gerenciar > Parametros Compartilhados.
3. Execute o script e selecione os parametros a exportar.
4. Informe o nome do grupo de destino no arquivo .txt.
"""
__title__ = 'Exportar\nParametros'
__author__ = 'Nn_Dev'

from pyrevit import revit, DB, forms, script

doc = revit.doc
app = doc.Application


def main():
    # 1. VALIDACAO: deve estar num Editor de Familias
    if not doc.IsFamilyDocument:
        forms.alert(
            "Execute dentro de um Editor de Familias (.rfa).",
            exitscript=True
        )

    # 2. VALIDACAO: arquivo de parametros compartilhados mapeado
    sp_file = app.OpenSharedParameterFile()
    if not sp_file:
        forms.alert(
            "Nenhum arquivo de Parametros Compartilhados (.txt) mapeado.\n"
            "Va em Gerenciar > Parametros Compartilhados e carregue um arquivo primeiro.",
            exitscript=True
        )

    # 3. COLETAR PARAMETROS COMPARTILHADOS DA FAMILIA
    sp_elements = DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement).ToElements()

    if not sp_elements:
        forms.alert("Nenhum parametro compartilhado encontrado nesta familia.", exitscript=True)

    param_dict = {}
    for sp_elem in sp_elements:
        try:
            sp_def  = sp_elem.GetDefinition()
            param_dict[sp_def.Name] = {
                "Def":  sp_def,
                "GUID": sp_elem.GuidValue
            }
        except Exception:
            pass

    if not param_dict:
        forms.alert("Nao foi possivel ler os parametros desta familia.", exitscript=True)

    # 4. SELECAO PELO USUARIO
    selected_names = forms.SelectFromList.show(
        sorted(param_dict.keys()),
        title="NnBim: Selecione os Parametros para Exportar",
        multiselect=True,
        button_name="Exportar >"
    )
    if not selected_names:
        script.exit()

    # 5. NOME DO GRUPO NO ARQUIVO .TXT
    group_name = forms.ask_for_string(
        default="NnBim_Exportados",
        prompt="Nome do Grupo de destino no arquivo .txt:",
        title="NnBim: Grupo de Destino"
    )
    if not group_name:
        script.exit()

    # 6. CRIAR GRUPO SE NAO EXISTIR
    group = sp_file.Groups.get_Item(group_name)
    if not group:
        try:
            group = sp_file.Groups.Create(group_name)
        except Exception as e:
            forms.alert(
                "Erro ao criar o grupo '{}'. O arquivo pode ser somente leitura.\n\nErro: {}".format(group_name, e),
                exitscript=True
            )

    # 7. EXPORTAR
    sucesso    = 0
    ignorados  = 0
    erros_list = []

    for name in selected_names:
        data   = param_dict[name]
        sp_def = data["Def"]
        guid   = data["GUID"]

        try:
            # Compatibilidade Revit 2022+ (GetDataType) e versoes anteriores (ParameterType)
            if hasattr(sp_def, "GetDataType"):
                opt = DB.ExternalDefinitionCreationOptions(sp_def.Name, sp_def.GetDataType())
            else:
                opt = DB.ExternalDefinitionCreationOptions(sp_def.Name, sp_def.ParameterType)

            opt.GUID           = guid
            opt.UserModifiable = True

            group.Definitions.Create(opt)
            sucesso += 1

        except Exception as e:
            ignorados += 1
            erro_str = str(e)
            if erro_str not in erros_list:
                erros_list.append(erro_str)

    # 8. RELATORIO FINAL
    msg = (
        "Exportacao concluida!\n\n"
        "Arquivo : {}\n"
        "Grupo   : {}\n\n"
        "{} parametro(s) exportado(s).\n"
        "{} parametro(s) ignorado(s) / com falha.\n"
    ).format(sp_file.Filename, group_name, sucesso, ignorados)

    if erros_list:
        msg += "\nMOTIVO(S) DA FALHA:\n"
        for erro in erros_list:
            msg += "- {}\n".format(erro)

    forms.alert(msg, title="NnBim: Relatorio de Exportacao")


if __name__ == '__main__':
    main()