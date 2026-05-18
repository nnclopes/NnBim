# -*- coding: utf-8 -*-
"""
NnBim: GESTOR DE TITULOS E ORDEM DE DESENHO
------------------------------------------------------------
Sincroniza nomes de vistas e a ordem dos detalhes
com as gavetas do carimbo (DESENHO 01-17).
VERSAO: 8.0 (Busca na Lib da Extensao)
"""
__title__ = 'Titulo das\nVistas'
__author__ = 'NnBim Dev'

import os
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import *

doc = revit.doc

# --- LOCALIZACAO DINAMICA NA PASTA LIB ---
try:
    lib_path  = script.get_bundle_lib_path()
    xaml_path = os.path.join(lib_path, "GerirTitulos.xaml")
except Exception:
    xaml_path = os.path.join(os.path.dirname(__file__), "GerirTitulos.xaml")


# --- CLASSE DE DADOS PARA A TABELA ---
class VistaData(object):
    def __init__(self, view_id, vp_id, sheet_id, prancha_num, num_detalhe, nome_vista, titulo_atual):
        self.ViewId        = view_id
        self.ViewportId    = vp_id
        self.SheetId       = sheet_id
        self.Prancha       = prancha_num
        self.NumeroDetalhe = num_detalhe
        self.NomeVista     = nome_vista
        self.TituloFolha   = titulo_atual if titulo_atual else ""


# --- MOTOR DE ATUALIZACAO DO CARIMBO ---
def atualizar_carimbo_desenhos(sheet):
    viewports    = [doc.GetElement(vpid) for vpid in sheet.GetAllViewports()]
    dados_vistas = []

    for vp in viewports:
        view = doc.GetElement(vp.ViewId)
        if view.ViewType in [ViewType.Legend, ViewType.Schedule]:
            continue

        num       = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
        tit_param = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString()
        nome_final = tit_param if (tit_param and tit_param.strip()) else view.Name

        dados_vistas.append({
            'ordem': int(num) if (num and num.isdigit()) else 99,
            'texto': nome_final.upper()
        })

    dados_vistas.sort(key=lambda x: x['ordem'])

    for i in range(1, 18):
        param = sheet.LookupParameter("DESENHO {:02d}".format(i))
        if param and not param.IsReadOnly:
            param.Set(dados_vistas[i-1]['texto'] if i <= len(dados_vistas) else "")


# --- INTERFACE WPF ---
class JanelaTitulos(forms.WPFWindow):
    def __init__(self, xaml_file, lista_dados):
        if not os.path.exists(xaml_file):
            forms.alert("ERRO: O arquivo 'GerirTitulos.xaml' nao foi encontrado na pasta LIB!")
            script.exit()

        forms.WPFWindow.__init__(self, xaml_file)
        self.dgVistas.ItemsSource = lista_dados

    def salvar_titulos(self, sender, e):
        self.Close()
        lista          = self.dgVistas.ItemsSource
        sheets_afetadas = list(set([item.SheetId for item in lista]))

        with revit.Transaction("NnBim: Sincronizar Vistas"):
            # 1. Nomes e titulos (com prefixo temp para evitar conflitos de nome)
            for item in lista:
                view = doc.GetElement(item.ViewId)
                vp   = doc.GetElement(item.ViewportId)

                try:
                    if view.Name != item.NomeVista:
                        view.Name = item.NomeVista
                except:
                    pass

                val_tit = item.TituloFolha.strip() if item.TituloFolha else ""
                view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).Set(val_tit.upper())

                if item.NumeroDetalhe:
                    vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(
                        "temp_" + str(item.NumeroDetalhe)
                    )

            # 2. Numero de detalhe definitivo
            for item in lista:
                if item.NumeroDetalhe:
                    vp = doc.GetElement(item.ViewportId)
                    vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(
                        str(item.NumeroDetalhe)
                    )

        # 3. Atualizacao dos carimbos
        with revit.Transaction("NnBim: Atualizar Dados Carimbo"):
            for s_id in sheets_afetadas:
                atualizar_carimbo_desenhos(doc.GetElement(s_id))

        forms.toast("Nomes, Titulos e Carimbos Atualizados!")


# --- INICIALIZACAO ---
def main():
    sheets     = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    sheet_dict = {s.SheetNumber + " - " + s.Name: s for s in sheets}

    sel = forms.SelectFromList.show(
        sorted(sheet_dict.keys()),
        multiselect=True,
        title="Gestao de Titulos: Selecione as Pranchas",
        button_name="Analisar Vistas"
    )
    if not sel:
        script.exit()

    vistas_dados = []
    for p_nome in sel:
        sheet = sheet_dict[p_nome]
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            v  = doc.GetElement(vp.ViewId)
            if v.ViewType in [ViewType.Legend, ViewType.Schedule]:
                continue

            num = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
            tit = v.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION).AsString()
            vistas_dados.append(
                VistaData(v.Id, vp.Id, sheet.Id, sheet.SheetNumber, num, v.Name, tit)
            )

    if vistas_dados:
        JanelaTitulos(xaml_path, vistas_dados).ShowDialog()
    else:
        forms.alert("Nenhuma viewport encontrada nas pranchas selecionadas.")


if __name__ == '__main__':
    main()