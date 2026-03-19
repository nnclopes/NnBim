# -*- coding: utf-8 -*-
"""
NnBim: AUDITORIA DE TÍTULOS DE VISTA
DESCRICAO:
Lê as pranchas selecionadas, encontra vistas que não possuem
um 'Título na Folha' preenchido e exibe uma tabela interativa
para que o usuário decida se quer adicionar um título ou não.
"""

__title__ = 'Título das\nVistas'
__author__ = 'NnBim Dev'

import os
import sys
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import *
import wpf

doc = revit.doc

# --- 1. CLASSE DE DADOS PARA A TABELA (DataGrid) ---
# Esta classe faz a ponte entre o Python e o visual WPF
class VistaData(object):
    def __init__(self, view_id, prancha, nome_vista):
        self.ViewId = view_id
        self.Prancha = prancha
        self.NomeVista = nome_vista
        self.TituloFolha = ""

# --- 2. CLASSE DA JANELA ---
class JanelaTitulos(forms.WPFWindow):
    def __init__(self, xaml_path, lista_dados):
        # Carrega o XAML da mesma pasta do script
        forms.WPFWindow.__init__(self, xaml_path)
        self.lista_dados = lista_dados
        
        # Injeta os dados na tabela
        self.dgVistas.ItemsSource = self.lista_dados

    def salvar_titulos(self, sender, e):
        # Quando clicar em Salvar, fecha a tela e roda a atualização
        self.Close()
        
        vistas_atualizadas = 0
        
        with revit.Transaction("NnBim: Atualização de Títulos na Folha"):
            with forms.ProgressBar(title='Aplicando Títulos... ({value} de {max_value})') as pb:
                for index, item in enumerate(self.lista_dados):
                    # Se o usuário digitou algo na coluna
                    if item.TituloFolha and item.TituloFolha.strip() != "":
                        view = doc.GetElement(item.ViewId)
                        param_titulo = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                        
                        if param_titulo and not param_titulo.IsReadOnly:
                            param_titulo.Set(item.TituloFolha.upper()) # Já salva em Maiúsculas
                            vistas_atualizadas += 1
                            
                    pb.update_progress(index + 1, len(self.lista_dados))

        if vistas_atualizadas > 0:
            forms.toast("Sucesso! {} títulos foram atualizados.".format(vistas_atualizadas))
        else:
            forms.toast("Nenhum título foi alterado.")

# --- 3. EXECUÇÃO PRINCIPAL ---
def main():
    # A. Coleta todas as pranchas do projeto para o usuário escolher
    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    
    if not sheets:
        forms.alert("Nenhuma prancha encontrada neste projeto.")
        script.exit()

    sheet_dict = {s.SheetNumber + " - " + s.Name: s for s in sheets}
    
    pranchas_selecionadas = forms.SelectFromList.show(
        sorted(sheet_dict.keys()),
        title="NnBim: Selecione as Pranchas para Auditar",
        multiselect=True,
        button_name="Analisar Vistas >"
    )
    
    if not pranchas_selecionadas:
        script.exit()

    # B. Motor de Raio-X nas vistas das pranchas selecionadas
    vistas_sem_titulo = []
    
    with forms.ProgressBar(title='Analisando Vistas...', cancellable=True) as pb:
        for i, p_nome in enumerate(pranchas_selecionadas):
            if pb.cancelled:
                break
                
            sheet = sheet_dict[p_nome]
            vp_ids = sheet.GetAllViewports()
            
            for vp_id in vp_ids:
                vp = doc.GetElement(vp_id)
                view = doc.GetElement(vp.ViewId)
                
                # Ignorar legendas e tabelas
                if view.ViewType in [ViewType.Legend, ViewType.Schedule]:
                    continue
                
                # Verifica a gaveta "Título na Folha"
                param_titulo = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                if param_titulo:
                    titulo_na_folha = param_titulo.AsString()
                    
                    # Se estiver vazio ou for nulo, é um candidato para a nossa auditoria!
                    if not titulo_na_folha or titulo_na_folha.strip() == "":
                        dado = VistaData(view.Id, sheet.SheetNumber, view.Name)
                        vistas_sem_titulo.append(dado)
                    
            pb.update_progress(i + 1, len(pranchas_selecionadas))

    # C. Abre a Janela se encontrou problemas
    if not vistas_sem_titulo:
        forms.alert("Parabéns! Todas as vistas nas pranchas selecionadas já possuem um 'Título na Folha' preenchido (ou não há vistas).", title="NnBim: Auditoria Perfeita")
    else:
        xaml_file = os.path.join(os.path.dirname(__file__), "GerirTitulos.xaml")
        
        # Verificação extra para garantir que o XAML está na mesma pasta
        if not os.path.exists(xaml_file):
            forms.alert("Arquivo 'GerirTitulos.xaml' não encontrado na pasta do botão.\nPor favor, salve-o na mesma pasta do script.py.", title="Erro de XAML")
            script.exit()
            
        janela = JanelaTitulos(xaml_file, vistas_sem_titulo)
        janela.ShowDialog()

if __name__ == '__main__':
    main()