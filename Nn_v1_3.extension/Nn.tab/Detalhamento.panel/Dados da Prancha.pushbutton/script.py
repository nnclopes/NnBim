# -*- coding: utf-8 -*-
"""
NnBim: SINCRONIZADOR DE CARIMBO
DESCRICAO:
Ferramenta de alta performance para gestão de pranchas. 
Atualiza dados de responsabilidade (Autor, Verificador), 
datas e revisões em lote. Sincroniza automaticamente os 
nomes das vistas contidas na prancha para a lista de desenhos.

COMO USAR:
1. Abra a ferramenta e selecione as pranchas na lista.
2. Na aba 'Dados em Lote', preencha apenas o que deseja alterar.
3. Deixe em branco os parâmetros que devem ser mantidos.
4. Clique em 'Executar' para sincronizar tudo instantaneamente.
"""

__title__ = 'Sincronizar\nCarimbo'
__author__ = 'NnBim Dev'

import sys
import os
import re
from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script

# --- 1. LOCALIZAÇÃO DA BIBLIOTECA NNBIM ---
cur_dir = os.path.dirname(__file__)
ext_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(cur_dir))))
lib_path = os.path.join(ext_path, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    from GUI.WPF_Base import my_WPF
except ImportError as e:
    forms.alert("Erro ao importar a base WPF:\n{}".format(e))
    script.exit()

import wpf

# --- 2. CLASSE DA INTERFACE (WPF) ---
class JanelaPranchas(my_WPF):
    def __init__(self, xaml_name):
        self.xaml_path = os.path.join(lib_path, 'GUI', xaml_name)
        wpf.LoadComponent(self, self.xaml_path)
        
        try:
            self.add_wpf_resource() 
        except:
            pass

        self._carregar_pranchas()

    def _carregar_pranchas(self):
        """Busca as pranchas e organiza de forma natural"""
        sheets = FilteredElementCollector(revit.doc).OfClass(ViewSheet).ToElements()
        
        def natural_sort(s):
            return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', s.SheetNumber)]
        
        sorted_sheets = sorted(sheets, key=natural_sort)
        self.lbPranchas.ItemsSource = [s.SheetNumber + " - " + s.Name for s in sorted_sheets]

    def executar_atualizacao(self, sender, e):
        """Cérebro da automação: Sincroniza tudo de forma segura"""
        selecionados = self.lbPranchas.SelectedItems
        if not selecionados:
            forms.alert("Por favor, selecione as pranchas na Aba 1.")
            return

        total_pranchas = len(selecionados)

        # Coleta os dados digitados
        d_desenhado = self.txtDesenhado.Text
        d_verificado = self.txtVerificado.Text
        d_projetado = self.txtProjetado.Text
        d_data_emissao = self.txtDataEmissao.Text
        d_revisao = self.txtRevisao.Text
        d_data_revisao = self.txtDataRevisao.Text
        d_disciplina = self.txtDisciplina.Text
        d_fase = self.txtFase.Text

        # 1. Parâmetros Nativos do Revit (Blindado contra idiomas)
        parametros_nativos = {
            BuiltInParameter.SHEET_DRAWN_BY: d_desenhado,
            BuiltInParameter.SHEET_CHECKED_BY: d_verificado,
            BuiltInParameter.SHEET_DESIGNED_BY: d_projetado,
            BuiltInParameter.SHEET_ISSUE_DATE: d_data_emissao
        }

        # 2. Parâmetros Customizados
        parametros_customizados = {
            "Revisão": d_revisao,
            "Data da revisão atual": d_data_revisao,
            "Disciplina da Prancha": d_disciplina,
            "Fase do Projeto": d_fase
        }

        # 🚨 PASSO CRÍTICO: Fecha a janela ANTES de começar o trabalho para o Revit não travar!
        self.Close()

        # Inicia o motor do Revit com Barra de Progresso
        with revit.Transaction("NnBim: Atualização Completa"):
            with forms.ProgressBar(title='Processando Pranchas... ({value} de {max_value})', cancellable=True) as pb:
                
                for index, item in enumerate(selecionados):
                    # Se o usuário clicar em Cancelar, o script para de forma segura
                    if pb.cancelled:
                        break

                    sheet_num = item.split(" - ")[0]
                    sheet = next(s for s in FilteredElementCollector(revit.doc).OfClass(ViewSheet) if s.SheetNumber == sheet_num)
                    
                    # A. Preenche os Parâmetros Nativos
                    for bip, valor in parametros_nativos.items():
                        if valor:
                            param_nativo = sheet.get_Parameter(bip)
                            if param_nativo and not param_nativo.IsReadOnly:
                                param_nativo.Set(valor)

                    # B. Preenche os Parâmetros Customizados
                    for param_nome, valor in parametros_customizados.items():
                        if valor:
                            param_custom = sheet.LookupParameter(param_nome)
                            if param_custom and not param_custom.IsReadOnly:
                                param_custom.Set(valor)

                    # C. Mapeia e Sincroniza os Desenhos
                    viewport_ids = sheet.GetAllViewports()
                    
                    # 🚨 CORREÇÃO DO BUG: A lista precisa ser zerada para CADA prancha!
                    lista_desenhos = [] 

                    for vp_id in viewport_ids:
                        vp = revit.doc.GetElement(vp_id)
                        view = revit.doc.GetElement(vp.ViewId)
                        
                        # Ignora tabelas e legendas
                        if view.ViewType in [ViewType.Legend, ViewType.Schedule]:
                            continue

                        detalhe_num = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
                        titulo_na_folha = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString()
                        nome_final = titulo_na_folha if titulo_na_folha else view.Name
                        
                        lista_desenhos.append({
                            'ordem': int(detalhe_num) if detalhe_num and detalhe_num.isdigit() else 99,
                            'texto': nome_final.upper()
                        })

                    # Organiza pela numeração da bolinha na prancha
                    lista_desenhos.sort(key=lambda x: x['ordem'])

                    # Escreve nos parâmetros DESENHO 01 a 17
                    for i in range(1, 18):
                        p_desenho = sheet.LookupParameter("DESENHO {:02d}".format(i))
                        if p_desenho and not p_desenho