# -*- coding: utf-8 -*-
"""
Titulo: LUM 01 - Iniciar Luminotecnico
Descricao: Cria parametros compartilhados nos Rooms e importa
           os resultados do calculo luminotecnico a partir do
           arquivo Excel (ESTUDO-CARGAS_LUMINOTECNICO).
           Cruza por nome do ambiente. Preparado para integracao
           com DiRoots.
Instrucoes de Uso: 1. Selecione os Rooms na vista ativa.
                   2. Rode este botao.
                   3. Selecione o arquivo .txt de parametros.
                   4. Selecione o Excel de calculo.
                   5. Confirme o preview e grave.
"""
__title__ = 'Exportar\nLumino_teste'
__author__ = 'NnBim Dev'
import os
import uuid
import re
import clr
import unicodedata

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, BuiltInParameterGroup,
    ExternalDefinitionCreationOptions,
    CategorySet, InstanceBinding, SpecTypeId,
    BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

import System
import System.Collections.ObjectModel as SCO
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Windows import Markup

from pyrevit import revit, script

doc   = revit.doc
app   = doc.Application
uidoc = revit.uidoc

# =============================================================================
# CONFIGURACAO DOS PARAMETROS
# =============================================================================
GRUPO_PARAMS = "NnBim_Luminotecnico"

PARAMS = [
    ("LUM_Altura_Luminaria",           SpecTypeId.Length,      BuiltInParameterGroup.PG_GEOMETRY),
    ("LUM_Iluminancia_Em",             SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Fluxo_Luminoso",             SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_FU",                         SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_FM",                         SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Indice_K",                   SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Quant_Luminaria_Calculado",  SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Quant_Luminaria_Adotado",    SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Lux_Real",                   SpecTypeId.Number,      BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Status_NBR",                 SpecTypeId.String.Text, BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Atividade_NBR",              SpecTypeId.String.Text, BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Tipo_Luminaria",             SpecTypeId.String.Text, BuiltInParameterGroup.PG_IDENTITY_DATA),
    ("LUM_Potencia_Unitaria",          SpecTypeId.Number,      BuiltInParameterGroup.PG_ELECTRICAL),
    ("LUM_Carga_Total",                SpecTypeId.Number,      BuiltInParameterGroup.PG_ELECTRICAL),
]

COL_MAP = [
    (2,  "LUM_Atividade_NBR",            "txt"),
    (6,  "LUM_Altura_Luminaria",          "len"),
    (7,  "LUM_Iluminancia_Em",            "num"),
    (8,  "LUM_Indice_K",                  "num"),
    (9,  "LUM_Tipo_Luminaria",            "txt"),
    (10, "LUM_Fluxo_Luminoso",            "num"),
    (11, "LUM_FU",                        "num"),
    (12, "LUM_FM",                        "num"),
    (13, "LUM_Quant_Luminaria_Calculado", "num"),
    (14, "LUM_Quant_Luminaria_Adotado",   "num"),
    (15, "LUM_Potencia_Unitaria",         "num"),
    (16, "LUM_Carga_Total",               "num"),
    (17, "LUM_Status_NBR",                "txt"),
    (19, "LUM_Lux_Real",                  "num"),
]

FT_TO_M = 0.3048


# =============================================================================
# UTILITARIOS
# =============================================================================
def normalizar_nome(nome):
    """
    Remove acentos via NFD, converte para uppercase,
    remove pontuacao que nao seja alfanumerico/espaco/barra/hifen.
    Usa apenas [A-Z0-9] apos remocao de acentos — seguro no IronPython 2.7.
    """
    nfd = unicodedata.normalize("NFD", unicode(nome))
    sem_acento = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    upper = sem_acento.upper()
    limpo = re.sub(r"[^A-Z0-9\s/\-]", "", upper)
    limpo = re.sub(r"\s+", " ", limpo).strip()
    return limpo


def get_room_name(room):
    """
    Le o nome do Room via BuiltInParameter.ROOM_NAME.
    Funciona independente do idioma do Revit (PT-BR, EN, etc).
    """
    p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    if p:
        val = p.AsString()
        if val and val.strip():
            return val.strip()
    return ""


def safe_float(valor):
    """Converte valor do Excel para float de forma segura."""
    try:
        return float(str(valor).replace(",", ".").strip())
    except Exception:
        return None


# =============================================================================
# LEITURA DO EXCEL VIA COM (Microsoft.Office.Interop.Excel)
# Retorna (lista_abas, dict{nome_aba: list_of_rows})
# Value2 ja retorna valores calculados para formulas.
# =============================================================================
def ler_excel_com(caminho):
    clr.AddReference("Microsoft.Office.Interop.Excel")
    import Microsoft.Office.Interop.Excel as Excel

    xl_app = Excel.ApplicationClass()
    xl_app.Visible       = False
    xl_app.DisplayAlerts = False
    wb = None

    try:
        wb = xl_app.Workbooks.Open(caminho)
        abas = [wb.Sheets[i].Name for i in range(1, wb.Sheets.Count + 1)]
        dados = {}
        for nome_aba in abas:
            ws      = wb.Sheets[nome_aba]
            usado   = ws.UsedRange
            max_row = usado.Rows.Count + usado.Row - 1
            max_col = usado.Columns.Count + usado.Column - 1
            linhas  = []
            for r in range(1, max_row + 1):
                row = []
                for c in range(1, max_col + 1):
                    v = ws.Cells[r, c].Value2
                    row.append(v)
                linhas.append(row)
            dados[nome_aba] = linhas          # ← FIX: atribuicao dentro do loop
        return abas, dados
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            xl_app.Quit()
        except Exception:
            pass
        try:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xl_app)
        except Exception:
            pass


# =============================================================================
# PASSO 1 — ROOMS SELECIONADOS
# =============================================================================
sel   = uidoc.Selection.GetElementIds()
rooms = []
for eid in sel:
    el = doc.GetElement(eid)
    if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
        rooms.append(el)

if not rooms:
    TaskDialog.Show("NnBim - Iniciar Luminotecnico",
                    "Nenhum Room selecionado.\n\n"
                    "Selecione os Rooms na vista ativa e rode o botao novamente.")
    script.exit()


# =============================================================================
# PASSO 2 — ARQUIVO .TXT DE PARAMETROS COMPARTILHADOS
# =============================================================================
def gerar_txt():
    tipos_txt = [
        "LENGTH", "NUMBER", "NUMBER", "NUMBER", "NUMBER",
        "NUMBER", "NUMBER", "NUMBER", "NUMBER", "TEXT",
        "TEXT",   "TEXT",   "NUMBER", "NUMBER"
    ]
    linhas = [
        "# Parametros Compartilhados - NnBim Luminotecnico",
        "*META\tVERSION\t2",
        "*META\tMINVERSION\t1",
        "*GROUP\tID\tNAME",
        "*GROUP\t1\t{}".format(GRUPO_PARAMS),
        "*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE",
    ]
    for (nome, _, _), dtype in zip(PARAMS, tipos_txt):
        linhas.append("*PARAM\t{}\t{}\t{}\t\t1\t1\t\t1".format(
            str(uuid.uuid4()).upper(), nome, dtype))
    return "\n".join(linhas)

dlg_open        = OpenFileDialog()
dlg_open.Title  = "Selecione o arquivo de Parametros Compartilhados (.txt)"
dlg_open.Filter = "Arquivo de texto (*.txt)|*.txt"
caminho_txt     = None

if dlg_open.ShowDialog() == DialogResult.OK:
    caminho_txt = dlg_open.FileName
else:
    td = TaskDialog("NnBim")
    td.MainInstruction = "Nenhum arquivo .txt selecionado."
    td.MainContent     = "Deseja criar um novo arquivo de parametros?"
    td.CommonButtons   = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    if td.Show() == TaskDialogResult.Yes:
        dlg_save          = SaveFileDialog()
        dlg_save.Title    = "Salvar arquivo de Parametros Compartilhados"
        dlg_save.Filter   = "Arquivo de texto (*.txt)|*.txt"
        dlg_save.FileName = "NnBim_Luminotecnico_Params.txt"
        if dlg_save.ShowDialog() == DialogResult.OK:
            caminho_txt = dlg_save.FileName
            with open(caminho_txt, "w", encoding="utf-8") as f:
                f.write(gerar_txt())

if not caminho_txt:
    script.exit()


# =============================================================================
# PASSO 3 — CRIAR PARAMETROS NO PROJETO
# =============================================================================
app.SharedParametersFilename = caminho_txt
def_file = app.OpenSharedParameterFile()
if def_file is None:
    TaskDialog.Show("NnBim - Erro", "Nao foi possivel abrir o arquivo de parametros.")
    script.exit()

grupo_def = None
for g in def_file.Groups:
    if g.Name == GRUPO_PARAMS:
        grupo_def = g
        break
if grupo_def is None:
    grupo_def = def_file.Groups.Create(GRUPO_PARAMS)

existentes_arq = set(d.Name for d in grupo_def.Definitions)
for nome, tipo, _ in PARAMS:
    if nome not in existentes_arq:
        opts         = ExternalDefinitionCreationOptions(nome, tipo)
        opts.Visible = True
        grupo_def.Definitions.Create(opts)

existentes_proj = set()
it = doc.ParameterBindings.ForwardIterator()
it.Reset()
while it.MoveNext():
    if it.Key:
        existentes_proj.add(it.Key.Name)

cats     = app.Create.NewCategorySet()
cat_room = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms)
cats.Insert(cat_room)
binding  = app.Create.NewInstanceBinding(cats)

criados = []
pulados = []

with Transaction(doc, "NnBim - Criar Parametros") as t:
    t.Start()
    for nome, _, grp_prop in PARAMS:
        if nome in existentes_proj:
            pulados.append(nome)
            continue
        defn = grupo_def.Definitions.get_Item(nome)
        if defn:
            doc.ParameterBindings.Insert(defn, binding, grp_prop)
            criados.append(nome)
    t.Commit()


# =============================================================================
# PASSO 4 — SELECIONAR EXCEL
# =============================================================================
dlg_xl        = OpenFileDialog()
dlg_xl.Title  = "Selecione o arquivo Excel de calculo luminotecnico"
dlg_xl.Filter = "Excel (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|Todos (*.*)|*.*"
if dlg_xl.ShowDialog() != DialogResult.OK:
    TaskDialog.Show("NnBim", "Nenhum arquivo Excel selecionado. Operacao cancelada.")
    script.exit()

caminho_xl = dlg_xl.FileName

try:
    abas_disponiveis, dados_excel = ler_excel_com(caminho_xl)
except Exception as ex:
    TaskDialog.Show("NnBim - Erro ao ler Excel",
                    "Nao foi possivel ler o arquivo Excel.\n\n"
                    "Certifique-se que o Excel esta instalado no computador.\n\n"
                    "Detalhe: {}".format(str(ex)))
    script.exit()


# =============================================================================
# PASSO 5 — SELECIONAR ABA DO EXCEL
# =============================================================================
XAML_ABA = (
    "<Window "
    "xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" "
    "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\" "
    "Title=\"NnBim - Selecionar Aba\" Width=\"400\" Height=\"300\" "
    "WindowStartupLocation=\"CenterScreen\" ResizeMode=\"NoResize\" "
    "Background=\"#1E1E2E\">"
    "<Grid Margin=\"24\">"
    "<Grid.RowDefinitions>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"*\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "</Grid.RowDefinitions>"
    "<TextBlock Grid.Row=\"0\" Text=\"SELECIONE A ABA DO PROJETO\" "
    "FontFamily=\"Segoe UI\" FontSize=\"13\" FontWeight=\"Bold\" "
    "Foreground=\"#89DCEB\" Margin=\"0,0,0,6\"/>"
    "<TextBlock Grid.Row=\"1\" Text=\"Selecione a aba correspondente ao projeto:\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" Foreground=\"#CDD6F4\" Margin=\"0,0,0,12\"/>"
    "<ListBox Grid.Row=\"2\" x:Name=\"lstAbas\" "
    "Background=\"#313244\" BorderThickness=\"0\" "
    "Foreground=\"#CDD6F4\" FontFamily=\"Segoe UI\" FontSize=\"11\" "
    "Margin=\"0,0,0,16\"/>"
    "<StackPanel Grid.Row=\"3\" Orientation=\"Horizontal\" HorizontalAlignment=\"Right\">"
    "<Button x:Name=\"btnCancelar\" Content=\"Cancelar\" "
    "Width=\"100\" Height=\"32\" Margin=\"0,0,10,0\" "
    "Background=\"#45475A\" Foreground=\"#CDD6F4\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" BorderThickness=\"0\"/>"
    "<Button x:Name=\"btnOk\" Content=\"Confirmar\" "
    "Width=\"110\" Height=\"32\" "
    "Background=\"#A6E3A1\" Foreground=\"#1E1E2E\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" FontWeight=\"Bold\" "
    "BorderThickness=\"0\"/>"
    "</StackPanel>"
    "</Grid>"
    "</Window>"
)

jAba     = Markup.XamlReader.Parse(XAML_ABA)
lst_abas = jAba.FindName("lstAbas")
for aba in abas_disponiveis:
    if not aba.startswith("_"):
        lst_abas.Items.Add(aba)
if lst_abas.Items.Count > 0:
    lst_abas.SelectedIndex = 0

res_aba = [None]

def on_ok_aba(s, e):
    if lst_abas.SelectedItem:
        res_aba[0] = str(lst_abas.SelectedItem)
        jAba.Close()
    else:
        TaskDialog.Show("NnBim", "Selecione uma aba.")

def on_cancel_aba(s, e):
    jAba.Close()

jAba.FindName("btnOk").Click       += on_ok_aba
jAba.FindName("btnCancelar").Click += on_cancel_aba
jAba.ShowDialog()

if not res_aba[0]:
    script.exit()

aba_sel = res_aba[0]

if aba_sel not in dados_excel:
    TaskDialog.Show("NnBim - Erro", "Aba '{}' nao encontrada nos dados lidos.".format(aba_sel))
    script.exit()

linhas_xl = dados_excel[aba_sel]


# =============================================================================
# PROCESSAR LINHAS DO EXCEL
# Estrutura esperada:
#   L1 = titulo geral
#   L2 = cabecalho ("AMBIENTE", "ATIVIDADE NBR", ...)
#   L3 = linha de unidades (None, "8995-1 / 5413", ...)
#   L4 = primeira linha de secao (col A preenchida, col B vazia) — ignorada
#   L5+ = ambientes reais (col A = nome, col B = atividade NBR)
# =============================================================================
data_start = None
for idx, row in enumerate(linhas_xl):
    if row and row[0] is not None:
        if str(row[0]).strip().upper() == "AMBIENTE":
            data_start = idx + 2   # pula cabecalho + linha de unidades
            break

if data_start is None:
    data_start = 3                 # fallback seguro

excel_map       = {}   # chave normalizada -> dict de parametros
excel_map_debug = {}   # chave normalizada -> nome original (para log)

for row in linhas_xl[data_start:]:
    if not row or row[0] is None:
        continue
    nome_amb = str(row[0]).strip()
    if not nome_amb:
        continue
    # Ignora linhas de secao: col B vazia = sem atividade NBR
    col_b = row[1] if len(row) > 1 else None
    if col_b is None or str(col_b).strip() == "":
        continue

    chave    = normalizar_nome(nome_amb)
    if not chave:
        continue

    row_data = {}
    for col_1based, param_nome, tipo in COL_MAP:
        idx_0 = col_1based - 1
        if idx_0 < len(row) and row[idx_0] is not None:
            val = row[idx_0]
            if tipo == "txt" and str(val).strip() == "":
                continue
            row_data[param_nome] = (val, tipo)

    excel_map[chave]       = row_data
    excel_map_debug[chave] = nome_amb


# =============================================================================
# PASSO 6 — CRUZAMENTO REVIT x EXCEL
# Prioridade: match exato > match parcial
# Usa BuiltInParameter.ROOM_NAME — funciona em PT-BR
# =============================================================================
encontrados     = []
nao_encontrados = []

for room in rooms:
    nome_room = get_room_name(room)
    if not nome_room:
        nao_encontrados.append("<Sem nome - Id {}>".format(room.Id.IntegerValue))
        continue

    chave_room = normalizar_nome(nome_room)

    if chave_room in excel_map:
        # Match exato
        encontrados.append((room, nome_room, excel_map[chave_room]))
    else:
        # Match parcial — busca mais especifica primeiro
        # Ordena chaves do Excel do maior para menor (evita "AREA ADMIN"
        # casar antes de "AREA ADMIN POSTO DE TRABALHO")
        parcial = None
        for chave_xl in sorted(excel_map.keys(), key=len, reverse=True):
            if chave_xl.startswith(chave_room) or chave_room.startswith(chave_xl):
                parcial = (chave_xl, excel_map[chave_xl])
                break
        if parcial:
            label = "{} [~{}]".format(nome_room, excel_map_debug[parcial[0]])
            encontrados.append((room, label, parcial[1]))
        else:
            nao_encontrados.append(nome_room)


# =============================================================================
# PASSO 6b — PREVIEW DO CRUZAMENTO
# =============================================================================
XAML_PREV = (
    "<Window "
    "xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" "
    "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\" "
    "Title=\"NnBim - Preview do Cruzamento\" Width=\"680\" Height=\"520\" "
    "WindowStartupLocation=\"CenterScreen\" "
    "Background=\"#1E1E2E\">"
    "<Grid Margin=\"20\">"
    "<Grid.RowDefinitions>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"*\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "</Grid.RowDefinitions>"
    "<TextBlock Grid.Row=\"0\" Text=\"PREVIEW - CRUZAMENTO REVIT x EXCEL\" "
    "FontFamily=\"Segoe UI\" FontSize=\"13\" FontWeight=\"Bold\" "
    "Foreground=\"#89DCEB\" Margin=\"0,0,0,4\"/>"
    "<TextBlock Grid.Row=\"1\" x:Name=\"txbResumo\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" "
    "Foreground=\"#A6E3A1\" Margin=\"0,0,0,10\"/>"
    "<DataGrid Grid.Row=\"2\" x:Name=\"dgPrev\" "
    "AutoGenerateColumns=\"False\" IsReadOnly=\"True\" "
    "Background=\"#313244\" Foreground=\"#CDD6F4\" "
    "FontFamily=\"Segoe UI\" FontSize=\"10\" "
    "GridLinesVisibility=\"Horizontal\" "
    "HeadersVisibility=\"Column\" "
    "RowBackground=\"#313244\" "
    "AlternatingRowBackground=\"#2A2A3E\" "
    "BorderThickness=\"0\" CanUserAddRows=\"False\" "
    "Margin=\"0,0,0,8\">"
    "<DataGrid.Columns>"
    "<DataGridTextColumn Header=\"Ambiente\" Binding=\"{Binding Nome}\" Width=\"160\"/>"
    "<DataGridTextColumn Header=\"Status\" Binding=\"{Binding Status}\" Width=\"100\"/>"
    "<DataGridTextColumn Header=\"Luminaria\" Binding=\"{Binding Lum}\" Width=\"160\"/>"
    "<DataGridTextColumn Header=\"Em lux\" Binding=\"{Binding Em}\" Width=\"70\"/>"
    "<DataGridTextColumn Header=\"N Adot.\" Binding=\"{Binding NAdot}\" Width=\"60\"/>"
    "<DataGridTextColumn Header=\"Analise\" Binding=\"{Binding Analise}\" Width=\"*\"/>"
    "</DataGrid.Columns>"
    "</DataGrid>"
    "<TextBlock Grid.Row=\"3\" x:Name=\"txbNaoEnc\" "
    "FontFamily=\"Segoe UI\" FontSize=\"10\" "
    "Foreground=\"#FAB387\" Margin=\"0,0,0,12\" TextWrapping=\"Wrap\"/>"
    "<StackPanel Grid.Row=\"4\" Orientation=\"Horizontal\" HorizontalAlignment=\"Right\">"
    "<Button x:Name=\"btnCancelar\" Content=\"Cancelar\" "
    "Width=\"100\" Height=\"32\" Margin=\"0,0,10,0\" "
    "Background=\"#45475A\" Foreground=\"#CDD6F4\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" BorderThickness=\"0\"/>"
    "<Button x:Name=\"btnGravar\" Content=\"Gravar no Revit\" "
    "Width=\"140\" Height=\"32\" "
    "Background=\"#A6E3A1\" Foreground=\"#1E1E2E\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" FontWeight=\"Bold\" "
    "BorderThickness=\"0\"/>"
    "</StackPanel>"
    "</Grid>"
    "</Window>"
)

jPrev = Markup.XamlReader.Parse(XAML_PREV)
dg    = jPrev.FindName("dgPrev")

jPrev.FindName("txbResumo").Text = (
    "Encontrados: {}  |  Nao encontrados: {}  |  Aba: {}".format(
        len(encontrados), len(nao_encontrados), aba_sel))

class ItemPrev(object):
    def __init__(self, nome, status, lum, em, nadot, analise):
        self.Nome    = nome
        self.Status  = status
        self.Lum     = lum
        self.Em      = em
        self.NAdot   = nadot
        self.Analise = analise

itens_prev = SCO.ObservableCollection[object]()
for room, nome, dados in encontrados:
    lum     = str(dados.get("LUM_Tipo_Luminaria",          ("", ""))[0])
    em      = str(dados.get("LUM_Iluminancia_Em",          ("", ""))[0])
    nadot   = str(dados.get("LUM_Quant_Luminaria_Adotado", ("", ""))[0])
    analise = str(dados.get("LUM_Status_NBR",              ("", ""))[0])
    status  = "~APROX" if "~" in nome else "OK"
    itens_prev.Add(ItemPrev(nome, status, lum, em, nadot, analise))

for nome in nao_encontrados:
    itens_prev.Add(ItemPrev(nome, "NAO ENCONTRADO", "-", "-", "-", "-"))

dg.ItemsSource = itens_prev

if nao_encontrados:
    jPrev.FindName("txbNaoEnc").Text = (
        "Nao encontrados: " +
        ", ".join(nao_encontrados[:10]) +
        (" ..." if len(nao_encontrados) > 10 else ""))

ok_grav = [False]

def on_grav(s, e):
    ok_grav[0] = True
    jPrev.Close()

def on_cancel_prev(s, e):
    jPrev.Close()

jPrev.FindName("btnGravar").Click   += on_grav
jPrev.FindName("btnCancelar").Click += on_cancel_prev
jPrev.ShowDialog()

if not ok_grav[0]:
    script.exit()


# =============================================================================
# PASSO 7 — GRAVAR NOS ROOMS
# =============================================================================
gravados = 0
erros    = []

with Transaction(doc, "NnBim - Importar Calculo Luminotecnico") as t:
    t.Start()
    for room, nome, dados in encontrados:
        for param_nome, (valor, tipo) in dados.items():
            try:
                p = room.LookupParameter(param_nome)
                if not p or p.IsReadOnly:
                    continue
                if tipo == "num":
                    v = safe_float(valor)
                    if v is not None:
                        p.Set(v)
                elif tipo == "len":
                    v = safe_float(valor)
                    if v is not None:
                        p.Set(v / FT_TO_M)
                elif tipo == "txt":
                    p.Set(str(valor))
            except Exception as ex:
                erros.append("{} / {}: {}".format(nome, param_nome, str(ex)))
        gravados += 1
    t.Commit()


# =============================================================================
# PASSO 8 — JANELA FINAL
# =============================================================================
XAML_FIN = (
    "<Window "
    "xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" "
    "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\" "
    "Title=\"NnBim - Importacao Concluida\" Width=\"460\" Height=\"340\" "
    "WindowStartupLocation=\"CenterScreen\" ResizeMode=\"NoResize\" "
    "Background=\"#1E1E2E\">"
    "<Grid Margin=\"24\">"
    "<Grid.RowDefinitions>"
    "<RowDefinition Height=\"Auto\"/>"
    "<RowDefinition Height=\"*\"/>"
    "<RowDefinition Height=\"Auto\"/>"
    "</Grid.RowDefinitions>"
    "<TextBlock Grid.Row=\"0\" Text=\"IMPORTACAO CONCLUIDA\" "
    "FontFamily=\"Segoe UI\" FontSize=\"16\" FontWeight=\"Bold\" "
    "Foreground=\"#A6E3A1\" Margin=\"0,0,0,20\"/>"
    "<Border Grid.Row=\"1\" Background=\"#313244\" CornerRadius=\"6\" "
    "Padding=\"16,12\" Margin=\"0,0,0,20\">"
    "<StackPanel>"
    "<TextBlock x:Name=\"txbExcel\"  FontFamily=\"Segoe UI\" FontSize=\"10\" Foreground=\"#89DCEB\" Margin=\"0,0,0,4\"/>"
    "<TextBlock x:Name=\"txbAba\"    FontFamily=\"Segoe UI\" FontSize=\"10\" Foreground=\"#89DCEB\" Margin=\"0,0,0,4\"/>"
    "<TextBlock x:Name=\"txbGrav\"   FontFamily=\"Segoe UI\" FontSize=\"11\" Foreground=\"#A6E3A1\" Margin=\"0,0,0,4\"/>"
    "<TextBlock x:Name=\"txbNEnc\"   FontFamily=\"Segoe UI\" FontSize=\"11\" Foreground=\"#FAB387\" Margin=\"0,0,0,4\"/>"
    "<TextBlock x:Name=\"txbParams\" FontFamily=\"Segoe UI\" FontSize=\"11\" Foreground=\"#CDD6F4\" Margin=\"0,0,0,4\"/>"
    "<TextBlock x:Name=\"txbErros\"  FontFamily=\"Segoe UI\" FontSize=\"10\" Foreground=\"#F38BA8\" Margin=\"0,0,0,12\" TextWrapping=\"Wrap\"/>"
    "<TextBlock Text=\"PROXIMO PASSO:\" FontFamily=\"Segoe UI\" FontSize=\"10\" "
    "FontWeight=\"Bold\" Foreground=\"#CBA6F7\" Margin=\"0,0,0,4\"/>"
    "<TextBlock Text=\"Exporte a Tabela de Ambientes via DiRoots para validacao.\" "
    "FontFamily=\"Segoe UI\" FontSize=\"11\" Foreground=\"#BAC2DE\"/>"
    "</StackPanel>"
    "</Border>"
    "<Button Grid.Row=\"2\" x:Name=\"btnOk\" Content=\"OK\" "
    "Height=\"36\" FontFamily=\"Segoe UI\" FontSize=\"12\" "
    "FontWeight=\"Bold\" BorderThickness=\"0\" "
    "Background=\"#A6E3A1\" Foreground=\"#1E1E2E\"/>"
    "</Grid>"
    "</Window>"
)

jFin = Markup.XamlReader.Parse(XAML_FIN)
jFin.FindName("txbExcel").Text  = "Excel: {}".format(os.path.basename(caminho_xl))
jFin.FindName("txbAba").Text    = "Aba: {}".format(aba_sel)
jFin.FindName("txbGrav").Text   = "Rooms gravados: {}".format(gravados)
jFin.FindName("txbNEnc").Text   = "Nao encontrados: {}".format(len(nao_encontrados))
jFin.FindName("txbParams").Text = "Parametros criados: {}  |  Ja existiam: {}".format(len(criados), len(pulados))
jFin.FindName("txbErros").Text  = ("Erros: " + " | ".join(erros[:5])) if erros else "Sem erros."

def on_ok_fin(s, e):
    jFin.Close()

jFin.FindName("btnOk").Click += on_ok_fin
jFin.ShowDialog()