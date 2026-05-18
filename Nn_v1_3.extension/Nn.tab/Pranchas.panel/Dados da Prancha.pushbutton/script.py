# -*- coding: utf-8 -*-
"""
NnBim: GESTOR DE DADOS DO CARIMBO
DESCRICAO:
Atualiza dados de responsabilidade, datas, revisoes e fases.
Injeta dados nas Folhas e nas Informacoes do Projeto para garantir
a leitura correta do carimbo.

COMO USAR:
1. Execute o script.
2. Selecione as pranchas na lista.
3. Preencha os campos desejados e clique em Atualizar.
"""
__title__ = 'Dados do\nCarimbo'
__author__ = 'NnBim Dev'

import sys
import os
import re
from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script

# --- LOCALIZACAO DA BIBLIOTECA NNBIM ---
cur_dir  = os.path.dirname(__file__)
ext_path = os.path.dirname(os.path.dirname(os.path.dirname(cur_dir)))
lib_path = os.path.join(ext_path, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    from GUI.WPF_Base import my_WPF
except ImportError as e:
    forms.alert("Erro ao importar a base WPF:\n{}".format(e))
    script.exit()

import wpf


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
        sheets = FilteredElementCollector(revit.doc).OfClass(ViewSheet).ToElements()

        def natural_sort(s):
            return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', s.SheetNumber)]

        sorted_sheets = sorted(sheets, key=natural_sort)
        self.lbPranchas.ItemsSource = [s.SheetNumber + " - " + s.Name for s in sorted_sheets]

        # CORRECAO: dict pre-calculado para evitar collector dentro do loop
        self._sheet_dict = {s.SheetNumber: s for s in sorted_sheets}

    def executar_atualizacao(self, sender, e):
        selecionados = self.lbPranchas.SelectedItems
        if not selecionados:
            forms.alert("Selecione as pranchas na lista.")
            return

        total_pranchas  = len(selecionados)
        d_desenhado     = self.txtDesenhado.Text
        d_verificado    = self.txtVerificado.Text
        d_projetado     = self.txtProjetado.Text
        d_data_emissao  = self.txtDataEmissao.Text
        d_revisao       = self.txtRevisao.Text
        d_data_revisao  = self.txtDataRevisao.Text
        d_disciplina    = self.txtDisciplina.Text
        d_fase          = self.txtFase.Text

        self.Close()

        with revit.Transaction("NnBim: Atualizacao de Carimbo"):

            # Ataque global: Informacoes do Projeto
            doc_info = revit.doc.ProjectInformation

            if d_fase:
                p_status = doc_info.get_Parameter(BuiltInParameter.PROJECT_STATUS)
                if p_status and not p_status.IsReadOnly:
                    p_status.Set(d_fase)

            if d_data_emissao:
                p_proj_date = doc_info.get_Parameter(BuiltInParameter.PROJECT_ISSUE_DATE)
                if p_proj_date and not p_proj_date.IsReadOnly:
                    p_proj_date.Set(d_data_emissao)

            # Ataque local: prancha por prancha
            parametros_nativos = {
                BuiltInParameter.SHEET_DRAWN_BY:    d_desenhado,
                BuiltInParameter.SHEET_CHECKED_BY:  d_verificado,
                BuiltInParameter.SHEET_DESIGNED_BY: d_projetado,
                BuiltInParameter.SHEET_ISSUE_DATE:  d_data_emissao
            }
            parametros_customizados = {
                "Disciplina":            d_disciplina,
                "Revisao":               d_revisao,
                "Data da revisao atual": d_data_revisao,
                "Status do Projeto":     d_fase
            }

            with forms.ProgressBar(title='Atualizando Carimbos... ({value} de {max_value})', cancellable=True) as pb:
                for index, item in enumerate(selecionados):
                    if pb.cancelled:
                        break

                    sheet_num = item.split(" - ")[0]
                    # CORRECAO: lookup no dict pre-calculado
                    sheet = self._sheet_dict.get(sheet_num)
                    if sheet is None:
                        continue

                    for bip, valor in parametros_nativos.items():
                        if valor:
                            param_nativo = sheet.get_Parameter(bip)
                            if param_nativo and not param_nativo.IsReadOnly:
                                param_nativo.Set(valor)

                    for param_nome, valor in parametros_customizados.items():
                        if valor:
                            param_custom = sheet.LookupParameter(param_nome)
                            if param_custom and not param_custom.IsReadOnly:
                                param_custom.Set(valor)

                    viewport_ids   = sheet.GetAllViewports()
                    lista_desenhos = []

                    for vp_id in viewport_ids:
                        vp   = revit.doc.GetElement(vp_id)
                        view = revit.doc.GetElement(vp.ViewId)

                        if view.ViewType in [ViewType.Legend, ViewType.Schedule]:
                            continue

                        detalhe_num     = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
                        titulo_na_folha = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString()
                        nome_final      = titulo_na_folha if titulo_na_folha else view.Name

                        lista_desenhos.append({
                            'ordem': int(detalhe_num) if detalhe_num and detalhe_num.isdigit() else 99,
                            'texto': nome_final.upper()
                        })

                    lista_desenhos.sort(key=lambda x: x['ordem'])

                    for i in range(1, 18):
                        p_desenho = sheet.LookupParameter("DESENHO {:02d}".format(i))
                        if p_desenho and not p_desenho.IsReadOnly:
                            if i <= len(lista_desenhos):
                                p_desenho.Set(lista_desenhos[i-1]['texto'])
                            else:
                                p_desenho.Set("")

                    pb.update_progress(index + 1, total_pranchas)

        if not pb.cancelled:
            forms.toast("Carimbos e Listas Atualizados!")


if __name__ == '__main__':
    janela = JanelaPranchas("GestorPranchas.xaml")
    janela.ShowDialog()