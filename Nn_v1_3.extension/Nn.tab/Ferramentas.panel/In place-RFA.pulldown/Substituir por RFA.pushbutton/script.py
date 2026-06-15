# -*- coding: utf-8 -*-
"""
NnBim Ferramentas: InPlace para RFA
DESCRICAO:
Transforma geometrias In-Loco isoladas em Familias Carregaveis (.rfa) ja existentes no projeto.
O script calcula a altura da peca, cria tipos novos automaticamente se necessario
e alinha a peca no local exato.

PASSO A PASSO:
1. Exploda a modelagem original para individualizar as pecas (use Explodir InPlace).
2. Selecione as pecas soltas na tela.
3. Clique neste botao.
4. Escolha a familia carregavel de destino e o parametro que controla a altura.

Nota: As pecas originais nao sao apagadas, apenas ocultadas na vista atual.
"""
__title__ = 'InPlace\npara\nRFA'
__author__ = 'NnBim Dev'

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
import math
import System.Collections.Generic

doc = __revit__.ActiveUIDocument.Document


def get_single_solid(element):
    opt = Options()
    opt.ComputeReferences = True
    geometry = element.get_Geometry(opt)
    if geometry:
        for obj in geometry:
            if isinstance(obj, Solid) and obj.Volume > 0:
                return obj
            elif isinstance(obj, GeometryInstance):
                for inst_geom in obj.GetInstanceGeometry():
                    if isinstance(inst_geom, Solid) and inst_geom.Volume > 0:
                        return inst_geom
    return None


def calcular_alvo_exato(el, solido):
    bbox = el.get_BoundingBox(None)
    if not bbox:
        return None, 0, 0

    pt_alvo = XYZ(
        (bbox.Max.X + bbox.Min.X) / 2.0,
        (bbox.Max.Y + bbox.Min.Y) / 2.0,
        bbox.Min.Z
    )

    angulo = 0.0
    maior_comp = 0
    for edge in solido.Edges:
        try:
            curve = edge.AsCurve()
            if isinstance(curve, Line) and abs(curve.Direction.Z) < 0.01:
                if curve.Length > maior_comp:
                    maior_comp = curve.Length
                    angulo = math.atan2(curve.Direction.Y, curve.Direction.X)
        except:
            pass

    h_feet = 0
    for edge in solido.Edges:
        try:
            curve = edge.AsCurve()
            if isinstance(curve, Line) and abs(curve.Direction.Z) > 0.99:
                if curve.Length > h_feet:
                    h_feet = curve.Length
        except:
            pass

    if h_feet == 0:
        h_feet = bbox.Max.Z - bbox.Min.Z

    return pt_alvo, angulo, h_feet


# 1. VALIDACAO DA SELECAO
selection = revit.get_selection()
if not selection:
    forms.alert(
        "Selecione os elementos individuais antes de rodar o script.",
        title="NnBim: Selecao Necessaria"
    )
    script.exit()

cat_id         = selection[0].Category.Id
nome_categoria = selection[0].Category.Name

# 2. ESCOLHA DA FAMILIA
familias_projeto = FilteredElementCollector(doc).OfClass(FamilySymbol)
opcoes_lista = {
    "{} : {}".format(s.Family.Name, Element.Name.__get__(s)): s
    for s in familias_projeto
    if s.Category and s.Category.Id == cat_id
}

if not opcoes_lista:
    forms.alert(
        "Nenhuma familia carregavel da categoria '{}' encontrada.\n"
        "Carregue a familia primeiro.".format(nome_categoria),
        title="Aviso NnBim"
    )
    script.exit()

escolha = forms.SelectFromList.show(
    sorted(opcoes_lista.keys()),
    title="Selecione a Familia de Destino",
    multiselect=False
)
if not escolha:
    script.exit()
symb_base = opcoes_lista[escolha]

# 3. CONFIGURACAO DE PARAMETROS
param_altura = forms.ask_for_string(
    default="Altura",
    prompt="Nome do parametro de TIPO que controla a altura:",
    title="Configuracao NnBim"
)
if not param_altura:
    param_altura = "Altura"

prefixo_marca = forms.ask_for_string(
    default="B",
    prompt="Prefixo para a Marca (Ex: B001, B002...):",
    title="Numeracao Automatica"
)
if not prefixo_marca:
    prefixo_marca = "B"

# 4. EXECUCAO
t = Transaction(doc, "NnBim: InPlace para RFA")
t.Start()

try:
    if not symb_base.IsActive:
        symb_base.Activate()

    total            = 0
    tipos_em_cache   = {}
    ids_gerados      = []
    elementos_ocultar = System.Collections.Generic.List[ElementId]()

    with forms.ProgressBar(total=len(selection), title="Alinhando Familias Carregaveis...") as pb:
        for el in selection:
            solido = get_single_solid(el)
            if not solido:
                pb.update_progress(total + 1)
                continue

            pt_alvo, angulo, h_feet = calcular_alvo_exato(el, solido)
            if not pt_alvo:
                continue

            h_mm      = int(round(h_feet * 304.8, 0))
            nome_tipo = "H_{}mm".format(h_mm)

            if h_mm not in tipos_em_cache:
                tipo_existente = next(
                    (doc.GetElement(tid) for tid in symb_base.Family.GetFamilySymbolIds()
                     if Element.Name.__get__(doc.GetElement(tid)) == nome_tipo),
                    None
                )
                if not tipo_existente:
                    novo_tipo = symb_base.Duplicate(nome_tipo)
                    p_tipo    = novo_tipo.LookupParameter(param_altura)
                    if p_tipo and not p_tipo.IsReadOnly:
                        p_tipo.Set(h_feet)
                    tipos_em_cache[h_mm] = novo_tipo
                else:
                    tipos_em_cache[h_mm] = tipo_existente

            tipo_atual = tipos_em_cache[h_mm]
            if not tipo_atual.IsActive:
                tipo_atual.Activate()

            inst = doc.Create.NewFamilyInstance(pt_alvo, tipo_atual, Structure.StructuralType.NonStructural)
            ids_gerados.append(inst.Id)

            if angulo != 0.0:
                eixo_z = Line.CreateBound(pt_alvo, pt_alvo + XYZ(0, 0, 1))
                ElementTransformUtils.RotateElement(doc, inst.Id, eixo_z, angulo)

            doc.Regenerate()

            bbox_novo = inst.get_BoundingBox(None)
            if bbox_novo:
                centro_novo_x = (bbox_novo.Max.X + bbox_novo.Min.X) / 2.0
                centro_novo_y = (bbox_novo.Max.Y + bbox_novo.Min.Y) / 2.0
                deslocamento  = XYZ(
                    pt_alvo.X - centro_novo_x,
                    pt_alvo.Y - centro_novo_y,
                    pt_alvo.Z - bbox_novo.Min.Z
                )
                if deslocamento.GetLength() > 0.003:
                    ElementTransformUtils.MoveElement(doc, inst.Id, deslocamento)

            p_mark = inst.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p_mark:
                p_mark.Set("{}{:03d}".format(prefixo_marca, total + 1))

            elementos_ocultar.Add(el.Id)
            total += 1
            pb.update_progress(total)

    if elementos_ocultar.Count > 0:
        doc.ActiveView.HideElements(elementos_ocultar)

    t.Commit()
    revit.get_selection().set_to(ids_gerados)
    forms.alert(
        "Substituicao concluida!\n\n{} elementos processados.".format(total),
        title="Relatorio NnBim"
    )

except Exception as e:
    if t.GetStatus() == TransactionStatus.Started:
        t.RollBack()
    forms.alert("Erro critico:\n\n" + str(e), title="Erro NnBim")