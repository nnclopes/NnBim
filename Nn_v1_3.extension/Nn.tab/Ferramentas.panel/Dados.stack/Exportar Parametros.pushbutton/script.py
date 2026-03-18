# -*- coding: utf-8 -*-
"""
NnBim: Exportar Parâmetros Compartilhados (Diagnóstico)
DESCRICAO:
Varre a família (incluindo legendas), exporta para o TXT e reporta o erro exato do Revit caso a gravação falhe.
"""
__title__ = 'Exportar\nParâmetros'
__author__ = 'Nn_Dev'

import clr
from pyrevit import revit, DB, forms, script

doc = revit.doc
app = doc.Application

def main():
    if not doc.IsFamilyDocument:
        forms.alert("Execute dentro de um Editor de Famílias (.rfa).", exitscript=True)

    sp_file = app.OpenSharedParameterFile()
    if not sp_file:
        forms.alert("Nenhum arquivo de Parâmetros Compartilhados (.txt) mapeado.", exitscript=True)

    sp_elements = DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement).ToElements()

    if not sp_elements:
        forms.alert("Nenhum parâmetro compartilhado encontrado.", exitscript=True)

    param_dict = {}
    for sp_elem in sp_elements:
        try:
            sp_def = sp_elem.GetDefinition()
            param_dict[sp_def.Name] = {
                "Def": sp_def,
                "GUID": sp_elem.GuidValue
            }
        except:
            pass

    selected_param_names = forms.SelectFromList.show(
        sorted(param_dict.keys()),
        title="Selecione os Parâmetros para Exportar",
        multiselect=True
    )

    if not selected_param_names:
        script.exit()

    group_name = forms.ask_for_string(
        default="NnBim_Legendas",
        prompt="Digite o nome do Grupo para salvar no arquivo .txt:"
    )

    if not group_name:
        script.exit()

    group = sp_file.Groups.get_Item(group_name)
    if not group:
        try:
            group = sp_file.Groups.Create(group_name)
        except Exception as e:
            forms.alert("Erro ao criar o grupo. O arquivo pode ser Somente Leitura.\nErro: {}".format(e), exitscript=True)

    sucesso = 0
    ignorados = 0
    erros_reais = [] # Lista para capturar o motivo da falha
    
    for name in selected_param_names:
        data = param_dict[name]
        sp_def = data["Def"]
        guid = data["GUID"]
        
        try:
            if hasattr(sp_def, "GetDataType"):
                opt = DB.ExternalDefinitionCreationOptions(sp_def.Name, sp_def.GetDataType())
            else:
                opt = DB.ExternalDefinitionCreationOptions(sp_def.Name, sp_def.ParameterType)
            
            opt.GUID = guid
            opt.UserModifiable = True # Forçando True para evitar bloqueios de sistema
            
            group.Definitions.Create(opt)
            sucesso += 1
            
        except Exception as e:
            ignorados += 1
            erro_str = str(e)
            if erro_str not in erros_reais:
                erros_reais.append(erro_str)

    # Relatório com o Diagnóstico
    msg_final = (
        "Exportação Concluída!\n\n"
        "📁 Arquivo: {}\n"
        "📂 Grupo: {}\n\n"
        "✅ {} Parâmetros exportados.\n"
        "⚠️ {} Parâmetros falharam/ignorados.\n".format(sp_file.Filename, group_name, sucesso, ignorados)
    )

    if erros_reais:
        msg_final += "\n🛑 MOTIVO DA FALHA (Revit API):\n"
        for erro in erros_reais:
            msg_final += "- {}\n".format(erro)

    forms.alert(msg_final, title="NnBim Dev - Relatório de Diagnóstico")

if __name__ == '__main__':
    main()