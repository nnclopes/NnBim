# -*- coding: utf-8 -*-
'''
NnBim: Dividir Curva (Universal)
DESCRICAO:
Divide Linhas de Modelo, Detalhe, Simbolicas e Croquis em X partes iguais.
Mantem o estilo original independentemente da categoria da linha.

COMO USAR:
1. Selecione um segmento de linha ou curva.
2. Execute o botao.
3. Digite a quantidade de divisoes.
'''

__title__ = 'Dividir\nLinha'
__author__ = 'NnBim Dev'

from Autodesk.Revit.DB import *
from pyrevit import revit, forms
from Snippets import _lines 

doc = __revit__.ActiveUIDocument.Document

def get_style_id(el):
    """Garimpo API: Busca o ID do estilo de forma universal via parametro interno."""
    p = el.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
    if p:
        return p.AsElementId()
    return None

def set_style_id(el, style_id):
    """Garimpo API: Aplica o ID do estilo de forma universal."""
    if style_id:
        p = el.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
        if p:
            p.Set(style_id)

def main():
    selection = revit.get_selection()
    if not selection:
        forms.alert("Selecione uma linha antes de iniciar.", title="NnBim")
        return
    
    el = selection[0]
    
    # Validacao robusta de geometria
    if not hasattr(el, "Location") or not hasattr(el.Location, "Curve"):
        forms.alert("O elemento selecionado nao possui uma curva valida.", title="Erro")
        return
    
    curva_geom = el.Location.Curve

    # UI de Entrada
    res = forms.ask_for_string(default="2", prompt="Quantidade de partes:", title="NnBim")
    if not res or not res.isdigit(): return
    n = int(res)

    with Transaction(doc, "NnBim: Dividir Curva") as t:
        t.Start()
        try:
            # Captura o estilo original antes de deletar
            original_style = get_style_id(el)
            
            # Chama a logica da sua LIB
            segmentos_geom = _lines.dividir_curva_em_partes(curva_geom, n)
            
            for g in segmentos_geom:
                nova_linha = None
                # Criacao baseada na classe do objeto original
                if isinstance(el, DetailCurve):
                    nova_linha = doc.Create.NewDetailCurve(doc.ActiveView, g)
                elif isinstance(el, ModelCurve):
                    nova_linha = doc.Create.NewModelCurve(g, el.SketchPlane)
                elif isinstance(el, SymbolicCurve):
                    nova_linha = doc.Create.NewSymbolicCurve(doc.ActiveView, g)
                
                # Aplica o estilo na nova linha
                if nova_linha and original_style:
                    set_style_id(nova_linha, original_style)
            
            # Deleta a original para completar a 'divisao'
            doc.Delete(el.Id)
            t.Commit()
            forms.toast("Dividido em {} partes!".format(n))
            
        except Exception as e:
            t.RollBack()
            forms.alert("Erro critico: {}".format(e))

if __name__ == "__main__":
    main()