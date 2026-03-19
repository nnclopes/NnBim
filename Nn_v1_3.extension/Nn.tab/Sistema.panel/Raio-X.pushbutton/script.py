
# -*- coding: utf-8 -*-
"""
NnBim: SCANNER RAIO-X GLOBAL
DESCRICAO:
Extrai todas as informações de parâmetros (Instância e Tipo)
de Famílias, Vistas, Pranchas ou Informações do Projeto.
"""

__title__ = 'Raio-X\nGlobal'
__author__ = 'NnBim Dev'

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

def main():
    selection = uidoc.Selection.GetElementIds()
    
    el = None
    el_type = None
    contexto = ""
    
    # 1. LÓGICA DE DECISÃO (Selecionado vs Global)
    if selection:
        if len(selection) > 1:
            forms.alert("Para o Raio-X ficar preciso, selecione apenas UM elemento por vez.")
            return
            
        el = doc.GetElement(selection[0])
        contexto = "Elemento Selecionado"
        
        # Tenta procurar o Tipo (Família) se existir
        if el.GetTypeId() != ElementId.InvalidElementId:
            el_type = doc.GetElement(el.GetTypeId())
            
    else:
        # Se não houver seleção, pergunta o que escannear
        opcoes = [
            "1. Informações do Projeto (Project Info)", 
            "2. Propriedades da Vista/Prancha Atual"
        ]
        
        escolha = forms.CommandSwitchWindow.show(
            opcoes, 
            message="Nenhum elemento selecionado. O que deseja escanear?"
        )
        
        if not escolha: 
            return # Utilizador cancelou
            
        if "Projeto" in escolha:
            el = doc.ProjectInformation
            contexto = "Informações do Projeto"
        else:
            el = doc.ActiveView
            contexto = "Vista / Prancha Atual"

    # 2. CABEÇALHO DO RELATÓRIO
    output.print_md("# 🚀 Raio-X NnBim: Scanner Global")
    output.print_md("### Copie as informações abaixo e mande para o seu Engenheiro Sénior (IA)")
    
    output.print_md("---")
    output.print_md("## 📦 Contexto: {}".format(contexto))
    
    categoria_nome = el.Category.Name if el.Category else "Sem Categoria Nativa"
    nome_elemento = el.Name if el.Name else "N/A"
    
    print("- Categoria: {}".format(categoria_nome))
    print("- Nome/Identificação: {}".format(nome_elemento))
    print("- ID do Elemento: {}".format(el.Id))
    
    # 3. FUNÇÃO DE EXTRAÇÃO DE PARÂMETROS
    def print_parameters(element, title):
        output.print_md("## {}".format(title))
        params = element.Parameters
        
        # Organiza por ordem alfabética para facilitar a leitura
        sorted_params = sorted(params, key=lambda p: p.Definition.Name)
        
        if not sorted_params:
            print("*Nenhum parâmetro encontrado nesta secção.*")
            return

        for p in sorted_params:
            nome = p.Definition.Name
            somente_leitura = "Sim" if p.IsReadOnly else "Não"
            
            bip = p.Definition.BuiltInParameter
            bip_str = str(bip) if str(bip) != "INVALID" else "Customizado (Shared/Project)"
            
            try:
                if p.StorageType == StorageType.String:
                    valor = p.AsString()
                else:
                    valor = p.AsValueString()
            except:
                valor = "N/A"
                
            if not valor: valor = "[Vazio]"
            
            print("🔹 Nome: **{}** | Valor: `{}` | Apenas Leitura: {} | Código: {}".format(nome, valor, somente_leitura, bip_str))

    # 4. IMPRESSÃO DOS RESULTADOS
    output.print_md("---")
    print_parameters(el, "🛠️ PARÂMETROS DIRETOS (Instância / Globais)")
    
    if el_type:
        output.print_md("---")
        print_parameters(el_type, "🏢 PARÂMETROS DE TIPO (Família)")
        
    output.print_md("---")
    output.print_md("**Fim do Raio-X. Clique no ecrã, dê Ctrl+A, depois Ctrl+C e cole no chat!**")

if __name__ == '__main__':
    main()