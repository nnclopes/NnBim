# -*- coding: utf-8 -*-
"""
NnBim: IMPORTAR E CATALOGAR RFAs
DESCRICAO:
Carrega familias de uma pasta e as organiza em colunas por categoria,
adicionando etiquetas de identificacao.

COMO USAR:
1. Esteja em uma vista que aceite insercao de familias (Planta de Piso).
2. Execute o script e selecione a pasta com os arquivos .rfa.
3. Selecione as familias desejadas na lista.
"""
__title__ = 'Catalogar Familia\nImportar-RFA'
__author__ = 'Nn_Dev'

import clr
import os
from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc


def obter_tipo_texto():
    return FilteredElementCollector(doc).OfClass(TextNoteType).FirstElementId()


def executar_importacao():
    # CORRECAO: view capturada aqui dentro, no momento da execucao
    view = doc.ActiveView

    # 1. SELECAO DE PASTA
    pasta = forms.pick_folder(title="Selecionar Pasta de Familias")
    if not pasta:
        return

    arquivos = [f for f in os.listdir(pasta) if f.endswith('.rfa')]
    if not arquivos:
        forms.alert("Nenhuma familia .rfa encontrada.")
        return

    # 2. SELECAO INTERATIVA
    selecionados = forms.SelectFromList.show(
        arquivos,
        multiselect=True,
        title="NnBim: Selecione as familias para o catalogo"
    )
    if not selecionados:
        return

    # 3. CARREGAMENTO E ORGANIZACAO
    txt_type_id   = obter_tipo_texto()
    catalogo_dict = {}

    with revit.Transaction("NnBim: Catalogar Familias"):
        for nome_arq in selecionados:
            caminho_completo = os.path.join(pasta, nome_arq)
            try:
                ref_fam = clr.Reference[Family]()
                if doc.LoadFamily(caminho_completo, ref_fam):
                    # CORRECAO: .Value para desreferenciar o objeto .NET
                    fam = ref_fam.Value
                    if fam is None:
                        continue

                    cat = fam.FamilyCategory.Name

                    # CORRECAO: validacao antes de acessar index [0]
                    ids = list(fam.GetFamilySymbolIds())
                    if not ids:
                        continue

                    simbolo = doc.GetElement(ids[0])
                    if not simbolo.IsActive:
                        simbolo.Activate()

                    if cat not in catalogo_dict:
                        catalogo_dict[cat] = []
                    catalogo_dict[cat].append(simbolo)
            except:
                continue

        # 4. POSICIONAMENTO EM GRADE
        doc.Regenerate()

        espaco_x = 12.0
        espaco_y = 10.0
        offset_x = 0

        for nome_cat, simbolos in catalogo_dict.items():
            TextNote.Create(doc, view.Id, XYZ(offset_x, 15, 0), nome_cat.upper(), txt_type_id)

            offset_y = 0
            for sym in simbolos:
                ponto = XYZ(offset_x, offset_y, 0)
                try:
                    doc.Create.NewFamilyInstance(ponto, sym, Structure.StructuralType.NonStructural)
                    info = "{}\nTipo: {}".format(sym.Family.Name, sym.Name)
                    TextNote.Create(doc, view.Id, ponto + XYZ(0, -3, 0), info, txt_type_id)
                except:
                    pass
                offset_y -= espaco_y

            offset_x += espaco_x

    forms.alert("Catalogo gerado com sucesso!", title="NnBim")


if __name__ == "__main__":
    executar_importacao()