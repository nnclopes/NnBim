# -*- coding: utf-8 -*-
"""
Título: Nn_GravarAmbiente
Descrição: Identifica o ambiente (Room) de cada elemento do modelo e grava
           o nome e número do ambiente nos parâmetros compartilhados
           Nn_Ambiente e Nn_Ambiente_Numero. Trata portas/janelas via
           FromRoom/ToRoom, pisos/forros via BoundingBox com fallback Z,
           e demais elementos via LocationPoint com fallback triplo.
Instruções de Uso: Execute o script pelo painel NnBim. Informe o caminho do
                   arquivo de parâmetros compartilhados (.txt) quando solicitado.
                   Escolha se deseja sobrescrever valores ja preenchidos ou
                   apenas preencher campos vazios.
"""
__title__ = 'Nn\nGravar\nAmbiente'
__author__ = 'NnBim Dev'
import os
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    ElementMulticategoryFilter,
    BuiltInCategory,
    BuiltInParameter,
    LocationPoint,
    SharedParameterElement,
    ExternalDefinitionCreationOptions,
    BuiltInParameterGroup,
    Transaction,
    XYZ,
)

try:
    from Autodesk.Revit.DB import ParameterType
    _REVIT_2024 = True
except ImportError:
    _REVIT_2024 = False

from pyrevit import forms, script

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

output = script.get_output()

# ──────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────
PARAM_AMBIENTE    = "Nn_Ambiente"
PARAM_NUMERO      = "Nn_Ambiente_Numero"
PARAM_GRUPO       = BuiltInParameterGroup.PG_IDENTITY_DATA
SEM_AMBIENTE_NOME = "Sem Ambiente"
SEM_AMBIENTE_NUM  = "--"
OFFSET_Z          = 0.1  # pes

CAT_ROOM_AWARE = {
    int(BuiltInCategory.OST_Doors),
    int(BuiltInCategory.OST_Windows),
}

CAT_BBOX = {
    int(BuiltInCategory.OST_Floors),
    int(BuiltInCategory.OST_Ceilings),
}

# Categorias "Model" que nao representam elementos fisicos posicionados
# (ex.: Materiais) e por isso devem ficar fora do binding e da varredura.
CAT_EXCLUIR = {
    int(BuiltInCategory.OST_Materials),
}

# ──────────────────────────────────────────────
# UTILITÁRIOS GERAIS
# ──────────────────────────────────────────────

def get_last_phase(doc):
    phases = doc.Phases
    if phases.Size == 0:
        return None
    return phases.get_Item(phases.Size - 1)


def get_all_model_categories(doc):
    cats = app.Create.NewCategorySet()
    for cat in doc.Settings.Categories:
        if not cat.AllowsBoundParameters:
            continue
        if cat.CategoryType.ToString() != "Model":
            continue
        if int(cat.Id.IntegerValue) in CAT_EXCLUIR:
            continue
        try:
            cats.Insert(cat)
        except Exception:
            pass
    return cats


def open_shared_param_file(txt_path):
    app.SharedParametersFilename = txt_path
    return app.OpenSharedParameterFile()


def get_or_create_definition(spf, group_name, param_name):
    grp = None
    for g in spf.Groups:
        if g.Name == group_name:
            grp = g
            break
    if grp is None:
        grp = spf.Groups.Create(group_name)
    for defn in grp.Definitions:
        if defn.Name == param_name:
            return defn
    if _REVIT_2024:
        opts = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
    else:
        from Autodesk.Revit.DB import SpecTypeId
        opts = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
    opts.Visible = True
    return grp.Definitions.Create(opts)


def bind_parameter(doc, defn, cat_set, param_group):
    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    while it.MoveNext():
        if it.Key.Name == defn.Name:
            return False
    binding = app.Create.NewInstanceBinding(cat_set)
    binding_map.Insert(defn, binding, param_group)
    return True


def get_param_by_name(element, name):
    for p in element.Parameters:
        if p.Definition.Name == name:
            return p
    return None


def set_param_value(element, param_name, value):
    p = get_param_by_name(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    p.Set(value)
    return True


def room_nome(room):
    if room is None:
        return None
    p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    return p.AsString() if p else None


def room_numero(room):
    if room is None:
        return None
    p = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
    return p.AsString() if p else None


def get_room_at(ponto, phase):
    try:
        return doc.GetRoomAtPoint(ponto, phase)
    except Exception:
        return None


# ──────────────────────────────────────────────
# ESTRATÉGIAS DE IDENTIFICAÇÃO DE AMBIENTE
# ──────────────────────────────────────────────

def resolver_room_aware(elem, phase):
    """Portas e janelas — FromRoom / ToRoom nativos."""
    try:
        from_id = elem.get_Parameter(BuiltInParameter.ROOM_FROM_ID)
        to_id   = elem.get_Parameter(BuiltInParameter.ROOM_TO_ID)

        r_from = doc.GetElement(from_id.AsElementId()) if from_id and from_id.AsElementId() != ElementId.InvalidElementId else None
        r_to   = doc.GetElement(to_id.AsElementId())   if to_id   and to_id.AsElementId()   != ElementId.InvalidElementId else None

        nomes   = []
        numeros = []

        if r_from is not None:
            nomes.append(room_nome(r_from)   or SEM_AMBIENTE_NOME)
            numeros.append(room_numero(r_from) or SEM_AMBIENTE_NUM)
        if r_to is not None:
            nomes.append(room_nome(r_to)   or SEM_AMBIENTE_NOME)
            numeros.append(room_numero(r_to) or SEM_AMBIENTE_NUM)

        if nomes:
            return " / ".join(nomes), " / ".join(numeros)
    except Exception:
        pass
    return SEM_AMBIENTE_NOME, SEM_AMBIENTE_NUM


def resolver_bbox(elem, phase):
    """Pisos e forros — centroide do BoundingBox com fallback +Z e -Z."""
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return SEM_AMBIENTE_NOME, SEM_AMBIENTE_NUM

        cx = (bb.Min.X + bb.Max.X) / 2.0
        cy = (bb.Min.Y + bb.Max.Y) / 2.0
        cz = (bb.Min.Z + bb.Max.Z) / 2.0

        for ponto in [
            XYZ(cx, cy, cz),
            XYZ(cx, cy, cz + OFFSET_Z),
            XYZ(cx, cy, cz - OFFSET_Z),
        ]:
            room = get_room_at(ponto, phase)
            if room is not None:
                return room_nome(room) or SEM_AMBIENTE_NOME, room_numero(room) or SEM_AMBIENTE_NUM

    except Exception:
        pass
    return SEM_AMBIENTE_NOME, SEM_AMBIENTE_NUM


def resolver_location_point(elem, phase):
    """
    Elementos genéricos com LocationPoint.
    Fallback triplo: ponto original, +Z, -Z.
    Retorna None se o elemento nao tem LocationPoint.
    """
    try:
        loc = elem.Location
        if loc is None or not isinstance(loc, LocationPoint):
            return None
        p = loc.Point
        for ponto in [
            p,
            XYZ(p.X, p.Y, p.Z + OFFSET_Z),
            XYZ(p.X, p.Y, p.Z - OFFSET_Z),
        ]:
            room = get_room_at(ponto, phase)
            if room is not None:
                return room_nome(room) or SEM_AMBIENTE_NOME, room_numero(room) or SEM_AMBIENTE_NUM
    except Exception:
        pass
    return SEM_AMBIENTE_NOME, SEM_AMBIENTE_NUM


# ──────────────────────────────────────────────
# FLUXO PRINCIPAL
# ──────────────────────────────────────────────

def main():

    output.print_md("# 🏠 Nn_Ambientes")
    output.print_md("---")

    # 1. CAMINHO DO ARQUIVO .TXT
    txt_path = forms.pick_file(
        file_ext="txt",
        title="Selecione o arquivo de Parametros Compartilhados (.txt)"
    )
    if not txt_path:
        forms.alert("Nenhum arquivo selecionado. Script cancelado.", title="Nn_Ambientes")
        return
    if not os.path.exists(txt_path):
        forms.alert("Arquivo nao encontrado:\n{}".format(txt_path), title="Nn_Ambientes")
        return

    output.print_md("📄 **Arquivo:** `{}`".format(txt_path))

    # 2. MODO DE PREENCHIMENTO
    modo = forms.CommandSwitchWindow.show(
        ["Sobrescrever tudo", "Pular ja preenchidos"],
        message="Como deseja processar os elementos?"
    )
    if not modo:
        forms.alert("Script cancelado.", title="Nn_Ambientes")
        return
    sobrescrever = (modo == "Sobrescrever tudo")
    output.print_md("⚙️ **Modo:** `{}`".format(modo))

    # 3. ULTIMA FASE
    phase = get_last_phase(doc)
    if phase is None:
        forms.alert("Nenhuma fase encontrada.", title="Nn_Ambientes")
        return
    output.print_md("📅 **Fase:** `{}`".format(phase.Name))

    # 4. PARAMETROS COMPARTILHADOS
    original_spf_path = app.SharedParametersFilename
    try:
        spf = open_shared_param_file(txt_path)
        if spf is None:
            forms.alert("Nao foi possivel abrir o arquivo .txt.", title="Nn_Ambientes")
            return

        cat_set = get_all_model_categories(doc)

        with Transaction(doc, "Nn_Ambientes - Vincular Parametros") as t:
            t.Start()
            try:
                defn_amb = get_or_create_definition(spf, "NnBim", PARAM_AMBIENTE)
                defn_num = get_or_create_definition(spf, "NnBim", PARAM_NUMERO)
                vinc_amb = bind_parameter(doc, defn_amb, cat_set, PARAM_GRUPO)
                vinc_num = bind_parameter(doc, defn_num, cat_set, PARAM_GRUPO)
                t.Commit()
                output.print_md("🔗 **{}:** {}".format(
                    PARAM_AMBIENTE, "vinculado agora" if vinc_amb else "ja existia"))
                output.print_md("🔗 **{}:** {}".format(
                    PARAM_NUMERO, "vinculado agora" if vinc_num else "ja existia"))
            except Exception as ex:
                t.RollBack()
                forms.alert("Erro ao vincular parametros:\n{}".format(str(ex)), title="Nn_Ambientes")
                return
    finally:
        app.SharedParametersFilename = original_spf_path

    # 5. VARREDURA
    output.print_md("---")
    output.print_md("🔍 **Iniciando varredura...**")

    # Restringe a varredura as categorias vinculadas (cat_set), evitando
    # iterar elementos sem geometria util (ex.: Materiais, que podem somar
    # dezenas de milhares de itens "fantasma" no projeto).
    cat_ids = List[ElementId]()
    for cat in cat_set:
        cat_ids.Add(cat.Id)
    filtro_cat = ElementMulticategoryFilter(cat_ids)

    elementos    = FilteredElementCollector(doc).WherePasses(filtro_cat).WhereElementIsNotElementType().ToElements()
    total        = 0
    com_ambiente = 0
    sem_ambiente = 0
    ignorados    = 0
    pulados      = 0

    with Transaction(doc, "Nn_Ambientes - Preencher Parametros") as t:
        t.Start()
        try:
            for elem in elementos:

                if elem.Category is None:
                    ignorados += 1
                    continue
                if elem.Category.CategoryType.ToString() != "Model":
                    ignorados += 1
                    continue

                cat_id = int(elem.Category.Id.IntegerValue)

                if not sobrescrever:
                    p_ex = get_param_by_name(elem, PARAM_AMBIENTE)
                    if p_ex and p_ex.AsString() not in ("", None, SEM_AMBIENTE_NOME):
                        pulados += 1
                        continue

                # Roteamento por estratégia
                if cat_id in CAT_ROOM_AWARE:
                    total += 1
                    nome_amb, num_amb = resolver_room_aware(elem, phase)

                elif cat_id in CAT_BBOX:
                    total += 1
                    nome_amb, num_amb = resolver_bbox(elem, phase)

                else:
                    resultado = resolver_location_point(elem, phase)
                    if resultado is None:
                        ignorados += 1
                        continue
                    total += 1
                    nome_amb, num_amb = resultado

                if nome_amb != SEM_AMBIENTE_NOME:
                    com_ambiente += 1
                else:
                    sem_ambiente += 1

                set_param_value(elem, PARAM_AMBIENTE, nome_amb)
                set_param_value(elem, PARAM_NUMERO,   num_amb)

            t.Commit()

        except Exception as ex:
            t.RollBack()
            forms.alert("Erro durante a varredura:\n{}".format(str(ex)), title="Nn_Ambientes")
            return

    # 6. LOG
    output.print_md("---")
    output.print_md("## 📊 Resultado")
    output.print_md("| | |")
    output.print_md("|---|---|")
    output.print_md("| **Total processado**              | {} |".format(total))
    output.print_md("| **Com ambiente**                  | {} |".format(com_ambiente))
    output.print_md("| **Sem ambiente**                  | {} |".format(sem_ambiente))
    output.print_md("| **Pulados (ja preenchidos)**       | {} |".format(pulados))
    output.print_md("| **Ignorados (sem estrategia)**    | {} |".format(ignorados))

    output.print_md("---")
    output.print_md("✅ **Nn_Ambientes concluido.**")


main()