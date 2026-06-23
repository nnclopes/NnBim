"""
Título: Nn_Dimensoes - Dimensionar Elementos
Descrição:
    Calcula e grava as dimensões reais (Comprimento e Largura) de elementos
    selecionados diretamente nos parâmetros compartilhados Nn_Comprimento e
    Nn_Largura. Utiliza o CurveLoop do sketch do elemento para obter medidas
    precisas, independente de rotação. Compatível com Pisos, Tetos e Paredes.

Instruções de Uso:
    1. Clique no botão 'Dimensões Elem'.
    2. Selecione um ou mais elementos (Piso, Teto ou Parede) no modelo.
    3. Confirme a seleção com Enter ou clique duplo.
    4. Os parâmetros Nn_Comprimento e Nn_Largura serão gravados automaticamente.
    Obs: Na primeira execução, o botão verifica e cria os parâmetros
    compartilhados automaticamente no arquivo .txt carregado no Revit.
"""
__title__ = "Dimensões\nElem"
__author__ = "NnBim Dev"

# -*- coding: utf-8 -*-

import os
import math

import clr
clr.AddReference("System")
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    CategorySet, InstanceBinding, ExternalDefinitionCreationOptions,
    Transaction, ElementId
)

try:
    from Autodesk.Revit.DB import SpecTypeId
    _USA_SPEC_TYPE_ID = True
except ImportError:
    _USA_SPEC_TYPE_ID = False

try:
    from Autodesk.Revit.DB import GroupTypeId
    _USA_GROUP_TYPE_ID = True
except ImportError:
    _USA_GROUP_TYPE_ID = False

from Autodesk.Revit.DB import ParameterType, BuiltInParameterGroup
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from pyrevit import revit, forms, script

# ─────────────────────────────────────────────────────────────────────────────
# Constantes e contexto
# ─────────────────────────────────────────────────────────────────────────────
doc   = revit.doc
uidoc = revit.uidoc
app   = doc.Application
out   = script.get_output()

# Nome do grupo no arquivo de parâmetros compartilhados
GRUPO_SP      = "NnBim_Dimensoes"
NOME_COMP     = "Nn_Comprimento"
NOME_LARG     = "Nn_Largura"

# Categorias suportadas pelo botão
CATEGORIAS_SUPORTADAS = [
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Walls,
]

# ─────────────────────────────────────────────────────────────────────────────
# Filtro de seleção interativa
# ─────────────────────────────────────────────────────────────────────────────
class FiltroDimensionamento(ISelectionFilter):
    """Aceita apenas Pisos, Tetos e Paredes."""

    def AllowElement(self, elemento):
        if elemento is None:
            return False
        cat = elemento.Category
        if cat is None:
            return False
        for bic in CATEGORIAS_SUPORTADAS:
            if cat.Id.IntegerValue == int(bic):
                return True
        return False

    def AllowReference(self, referencia, ponto):
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de compatibilidade Revit 2022+ / 2025+
# ─────────────────────────────────────────────────────────────────────────────
def _obter_spec_length():
    """Retorna o SpecTypeId ou ParameterType correto para comprimento."""
    if _USA_SPEC_TYPE_ID:
        return SpecTypeId.Length
    return ParameterType.Length


def _obter_group_id():
    """Retorna o GroupTypeId ou BuiltInParameterGroup correto para Data."""
    if _USA_GROUP_TYPE_ID:
        return GroupTypeId.Data
    return BuiltInParameterGroup.PG_DATA


# ─────────────────────────────────────────────────────────────────────────────
# Setup de parâmetros compartilhados
# ─────────────────────────────────────────────────────────────────────────────
def _garantir_arquivo_sp():
    """
    Verifica se há um arquivo de parâmetros compartilhados carregado.
    Se não houver, abre diálogo para o usuário selecionar ou criar um.
    Retorna o arquivo aberto ou None em caso de cancelamento.
    """
    sp_file = app.OpenSharedParameterFile()
    if sp_file:
        return sp_file

    # Nenhum arquivo carregado — pedir ao usuário
    forms.alert(
        "Nenhum arquivo de Parâmetros Compartilhados está carregado no Revit.\n\n"
        "Na próxima janela, selecione o arquivo '.txt' de parâmetros "
        "compartilhados do NnBim (ou crie um novo arquivo de texto vazio).",
        title="NnBim — Parâmetros Compartilhados",
        warn_icon=True
    )

    caminho = forms.pick_file(
        file_ext="txt",
        title="Selecionar arquivo de Parâmetros Compartilhados (.txt)"
    )

    if not caminho:
        forms.alert(
            "Operação cancelada. Nenhum arquivo selecionado.",
            title="NnBim — Cancelado"
        )
        return None

    # Se o arquivo estiver vazio, escreve o cabeçalho obrigatório do Revit
    if os.path.getsize(caminho) == 0:
        try:
            with open(caminho, "w") as f:
                f.write(
                    "# This is a Revit shared parameter file.\n"
                    "# Do not edit manually.\n"
                    "*META\tVERSION\tMINVERSION\n"
                    "META\t2\t1\n"
                    "*GROUP\tID\tNAME\n"
                    "*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\t"
                    "VISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n"
                )
        except Exception as e:
            forms.alert(
                "Erro ao inicializar o arquivo de parâmetros:\n" + str(e),
                title="Erro NnBim"
            )
            return None

    try:
        app.SharedParametersFilename = caminho
        sp_file = app.OpenSharedParameterFile()
    except Exception as e:
        forms.alert(
            "Erro ao carregar o arquivo de parâmetros:\n" + str(e),
            title="Erro NnBim"
        )
        return None

    return sp_file


def setup_parametro(nome_param, categorias_bic):
    """
    Verifica se o parâmetro já existe e está vinculado às categorias.
    Se não existir, cria no arquivo de shared params carregado.
    Retorna: "existe", "criado", "atualizado" ou lança exceção.
    """
    # ── 1. Verificar se já existe no documento ────────────────────────────────
    definicao_existente  = None
    vinculo_existente    = None
    iterador = doc.ParameterBindings.ForwardIterator()
    while iterador.MoveNext():
        if iterador.Key.Name == nome_param:
            definicao_existente = iterador.Key
            vinculo_existente   = iterador.Current
            break

    if definicao_existente and vinculo_existente:
        # Verificar se todas as categorias necessárias estão vinculadas
        cat_set      = vinculo_existente.Categories
        precisa_atualizar = False
        for bic in categorias_bic:
            cat = doc.Settings.Categories.get_Item(bic)
            if cat and cat.AllowsBoundParameters and not cat_set.Contains(cat):
                cat_set.Insert(cat)
                precisa_atualizar = True

        if precisa_atualizar:
            t_rebind = Transaction(doc, "NnBim: Atualizar categorias " + nome_param)
            t_rebind.Start()
            try:
                novo_vinculo = app.Create.NewInstanceBinding(cat_set)
                doc.ParameterBindings.ReInsert(definicao_existente, novo_vinculo)
                t_rebind.Commit()
                return "atualizado"
            except Exception as e:
                t_rebind.RollBack()
                raise Exception("Falha ao atualizar categorias: " + str(e))
        return "existe"

    # ── 2. Parâmetro não existe — garantir arquivo SP carregado ───────────────
    sp_file = _garantir_arquivo_sp()
    if not sp_file:
        raise Exception("Arquivo de parâmetros compartilhados não disponível.")

    # ── 3. Localizar ou criar a definição no grupo NnBim_Dimensoes ────────────
    grupo = sp_file.Groups.get_Item(GRUPO_SP)
    if not grupo:
        grupo = sp_file.Groups.Create(GRUPO_SP)

    definicao_alvo = grupo.Definitions.get_Item(nome_param)
    if not definicao_alvo:
        try:
            opcoes = ExternalDefinitionCreationOptions(nome_param, _obter_spec_length())
        except Exception:
            # Fallback para API mais antiga
            opcoes = ExternalDefinitionCreationOptions(nome_param, ParameterType.Length)
        definicao_alvo = grupo.Definitions.Create(opcoes)

    if not definicao_alvo:
        raise Exception("Não foi possível criar a definição: " + nome_param)

    # ── 4. Montar CategorySet com as categorias suportadas ───────────────────
    cat_set = app.Create.NewCategorySet()
    for bic in categorias_bic:
        try:
            cat = doc.Settings.Categories.get_Item(bic)
            if cat and cat.AllowsBoundParameters:
                cat_set.Insert(cat)
        except Exception:
            pass

    if cat_set.IsEmpty:
        raise Exception("Nenhuma categoria válida encontrada para vincular.")

    # ── 5. Criar o vínculo de instância ──────────────────────────────────────
    vinculo   = app.Create.NewInstanceBinding(cat_set)
    t_param   = Transaction(doc, "NnBim: Criar parâmetro " + nome_param)
    t_param.Start()
    try:
        try:
            doc.ParameterBindings.Insert(definicao_alvo, vinculo, _obter_group_id())
        except Exception:
            doc.ParameterBindings.Insert(definicao_alvo, vinculo,
                                         BuiltInParameterGroup.PG_DATA)
        t_param.Commit()
        return "criado"
    except Exception as e:
        t_param.RollBack()
        raise Exception("Falha ao vincular parâmetro ao documento: " + str(e))


# ─────────────────────────────────────────────────────────────────────────────
# Extração de dimensões via CurveLoop do sketch
# ─────────────────────────────────────────────────────────────────────────────
def _comprimento_curva(curva):
    """Retorna o comprimento de uma curva da API do Revit."""
    try:
        return curva.Length
    except Exception:
        # Fallback: distância euclidiana entre os pontos extremos
        p0 = curva.GetEndPoint(0)
        p1 = curva.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        dz = p1.Z - p0.Z
        return math.sqrt(dx * dx + dy * dy + dz * dz)


def _obter_sketch_elemento(elemento):
    """
    Tenta acessar o Sketch de um Floor, Ceiling ou Wall.
    Retorna o objeto Sketch ou None.
    """
    try:
        # Floor e Ceiling expõem SketchId diretamente
        sketch_id = elemento.SketchId
        if sketch_id and sketch_id != ElementId.InvalidElementId:
            return doc.GetElement(sketch_id)
    except AttributeError:
        pass

    # Fallback: percorrer sub-elementos (Walls e casos especiais)
    try:
        sketch_ids = elemento.GetDependentElements(None)
        for sid in sketch_ids:
            sub = doc.GetElement(sid)
            if sub and sub.GetType().Name == "Sketch":
                return sub
    except Exception:
        pass

    return None


def calcular_dimensoes(elemento):
    """
    Extrai Comprimento (maior aresta) e Largura (menor aresta) do CurveLoop
    do sketch do elemento.
    Retorna (comprimento_pes, largura_pes) ou lança exceção descritiva.
    """
    sketch = _obter_sketch_elemento(elemento)

    if sketch is None:
        raise Exception(
            "Sketch não acessível. Verifique se o elemento é um Piso, "
            "Teto ou Parede modelada (não elementos de sistema)."
        )

    # Coletar todos os comprimentos de arestas do perfil externo
    comprimentos_arestas = []
    try:
        curve_loops = sketch.Profile
        # Iterar os CurveLoops do perfil
        enum = curve_loops.GetEnumerator()
        while enum.MoveNext():
            loop = enum.Current
            loop_enum = loop.GetEnumerator()
            while loop_enum.MoveNext():
                curva = loop_enum.Current
                comp  = _comprimento_curva(curva)
                if comp > 1e-6:   # ignorar arestas degeneradas
                    comprimentos_arestas.append(comp)
    except Exception as e:
        raise Exception("Falha ao ler o CurveLoop do sketch: " + str(e))

    if not comprimentos_arestas:
        raise Exception("Nenhuma aresta encontrada no sketch do elemento.")

    comprimento = max(comprimentos_arestas)  # aresta mais longa
    largura     = min(comprimentos_arestas)  # aresta mais curta

    return comprimento, largura


# ─────────────────────────────────────────────────────────────────────────────
# Função principal
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # ── Fase 1: Garantir que os parâmetros existem ───────────────────────────
    out.print_md("## NnBim — Dimensões de Elementos")
    out.print_md("### Verificando parâmetros compartilhados...")

    try:
        status_comp = setup_parametro(NOME_COMP, CATEGORIAS_SUPORTADAS)
        status_larg = setup_parametro(NOME_LARG,  CATEGORIAS_SUPORTADAS)
    except Exception as e:
        forms.alert(
            "Erro ao configurar parâmetros compartilhados:\n\n" + str(e),
            title="Erro NnBim"
        )
        return

    # Feedback de setup
    icones = {"existe": "✅", "criado": "🆕", "atualizado": "🔄"}
    out.print_md("- {} **{}**: {}".format(
        icones.get(status_comp, "❓"), NOME_COMP, status_comp))
    out.print_md("- {} **{}**: {}".format(
        icones.get(status_larg, "❓"), NOME_LARG, status_larg))

    # ── Fase 2: Seleção interativa ────────────────────────────────────────────
    out.print_md("### Aguardando seleção de elementos...")

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            FiltroDimensionamento(),
            "Selecione Pisos, Tetos ou Paredes — confirme com Enter"
        )
    except Exception:
        # Usuário cancelou com Esc
        out.print_md("⚠️ Seleção cancelada pelo usuário.")
        return

    if not refs:
        forms.alert("Nenhum elemento selecionado.", title="NnBim — Dimensões")
        return

    elementos = [doc.GetElement(ref.ElementId) for ref in refs]
    out.print_md("**{} elemento(s) selecionado(s).**".format(len(elementos)))

    # ── Fase 3: Calcular e gravar ─────────────────────────────────────────────
    out.print_md("### Calculando e gravando dimensões...")

    erros      = []
    sucessos   = 0

    # Importar ef_Transaction somente aqui para manter compatibilidade
    from Snippets._context_manager import ef_Transaction

    with ef_Transaction(doc, __title__, debug=False):
        for elemento in elementos:
            nome_elem = "ID {}".format(elemento.Id.IntegerValue)
            try:
                # Tentar obter nome mais descritivo
                nome_tipo = doc.GetElement(elemento.GetTypeId())
                if nome_tipo:
                    nome_elem = "{} ({})".format(nome_tipo.Name, nome_elem)
            except Exception:
                pass

            try:
                comp_pes, larg_pes = calcular_dimensoes(elemento)

                # Gravar no parâmetro Nn_Comprimento
                p_comp = elemento.LookupParameter(NOME_COMP)
                if not p_comp:
                    raise Exception("Parâmetro '{}' não encontrado no elemento. "
                                    "Recarregue o pyRevit.".format(NOME_COMP))
                if p_comp.IsReadOnly:
                    raise Exception("Parâmetro '{}' está somente leitura.".format(NOME_COMP))
                p_comp.Set(comp_pes)

                # Gravar no parâmetro Nn_Largura
                p_larg = elemento.LookupParameter(NOME_LARG)
                if not p_larg:
                    raise Exception("Parâmetro '{}' não encontrado no elemento. "
                                    "Recarregue o pyRevit.".format(NOME_LARG))
                if p_larg.IsReadOnly:
                    raise Exception("Parâmetro '{}' está somente leitura.".format(NOME_LARG))
                p_larg.Set(larg_pes)

                # Converter para metros para exibição (1 pé = 0.3048 m)
                comp_m = comp_pes * 0.3048
                larg_m = larg_pes * 0.3048
                out.print_md(
                    "- ✅ **{}** — Comprimento: `{:.3f} m` | Largura: `{:.3f} m`".format(
                        nome_elem, comp_m, larg_m
                    )
                )
                sucessos += 1

            except Exception as e:
                erros.append((nome_elem, str(e)))

    # ── Fase 4: Relatório de erros ────────────────────────────────────────────
    if erros:
        out.print_md("### ⚠️ Elementos com erro ({})".format(len(erros)))
        for nome, msg in erros:
            out.print_md("- ❌ **{}**: {}".format(nome, msg))

    # ── Fase 5: Toast de resumo ───────────────────────────────────────────────
    msg_toast = "{} elemento(s) dimensionado(s) com sucesso.".format(sucessos)
    if erros:
        msg_toast += " {} com erro (ver output).".format(len(erros))

    forms.toast(msg_toast, title="NnBim — Dimensões")
    out.print_md("---\n**Concluído.** {}".format(msg_toast))


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()