# -*- coding: utf-8 -*-
"""
Titulo: Importar Luminotecnico
Descricao: Le o arquivo Excel de calculo luminotecnico (gerado pelos botoes
           Calculo por ambiente / LUM_CAL_PAVIMENTOS) e grava de volta nos
           Rooms do modelo os valores calculados: LUM_Atividade_NBR,
           LUM_Altura_Luminaria, LUM_Iluminancia_Em, LUM_Indice_K,
           LUM_Tipo_Luminaria, LUM_Fluxo_Luminoso, LUM_FU, LUM_FM,
           LUM_Quant_Luminaria_Calculado, LUM_Quant_Luminaria_Adotado,
           LUM_Potencia_Unitaria e LUM_Carga_Total.
           O cruzamento entre a planilha e os Rooms e feito pelo GUID
           (coluna A), o mesmo GUID gravado na exportacao.
Instrucoes de Uso:
    1. Clique em Importar Luminotecnico na aba NnBim
    2. Selecione o arquivo Excel de calculo ja preenchido
    3. Escolha a aba com os resultados (ex.: UOP)
    4. Clique em Importar
"""
__title__ = 'Importar\nLuminotecnico'
__author__ = 'NnBim Dev'

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SpatialElement,
    BuiltInCategory,
    BuiltInParameter,
    ExternalDefinitionCreationOptions,
    Transaction,
    ForgeTypeId,
)

import System
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Windows.Markup import XamlReader
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
import System.Windows.Media as Media

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# ── CONSTANTES ────────────────────────────────────────────────────────────────

GRUPO_PARAM = u"NnBim - LUMINOTECNICO"
FT_TO_M     = 0.3048

_NUMBER = ForgeTypeId(u"autodesk.spec.aec:number-2.0.0")
_TEXT   = ForgeTypeId(u"autodesk.spec:spec.string-2.0.0")
_LENGTH = ForgeTypeId(u"autodesk.spec.aec:length-2.0.0")

PARAMS_LUM = [
    (u"LUM_Atividade_NBR",             _TEXT,   u"Dados"),  # Excel D - ATIVIDADE NBR
    (u"LUM_Altura_Luminaria",          _LENGTH, u"Dados"),  # Excel I - ALTURA LUMINARIA
    (u"LUM_Iluminancia_Em",            _NUMBER, u"Dados"),  # Excel J - ILUMINANCIA
    (u"LUM_Indice_K",                  _NUMBER, u"Dados"),  # Excel K - INDICE DO LOCAL
    (u"LUM_Tipo_Luminaria",            _TEXT,   u"Dados"),  # Excel L - LUMINARIA
    (u"LUM_Fluxo_Luminoso",            _NUMBER, u"Dados"),  # Excel M - FLUXO LUMINOSO
    (u"LUM_FU",                        _NUMBER, u"Dados"),  # Excel N - FU
    (u"LUM_FM",                        _NUMBER, u"Dados"),  # Excel O - FM
    (u"LUM_Quant_Luminaria_Calculado", _NUMBER, u"Dados"),  # Excel P - N CALC
    (u"LUM_Quant_Luminaria_Adotado",   _NUMBER, u"Dados"),  # Excel Q - N ADOTADO
    (u"LUM_Potencia_Unitaria",         _NUMBER, u"Dados"),  # Excel R - POTENCIA
    (u"LUM_Carga_Total",               _NUMBER, u"Dados"),  # Excel S - CARGA ILUM
]

LINHA_INICIO = 5
COL_GUID     = 1
NUM_COLS     = 19

# (coluna na planilha, nome do parametro, tipo)
COL_PARAMS = [
    (4,  u"LUM_Atividade_NBR",             "txt"),
    (9,  u"LUM_Altura_Luminaria",          "len"),
    (10, u"LUM_Iluminancia_Em",            "num"),
    (11, u"LUM_Indice_K",                  "num"),
    (12, u"LUM_Tipo_Luminaria",            "txt"),
    (13, u"LUM_Fluxo_Luminoso",            "num"),
    (14, u"LUM_FU",                        "num"),
    (15, u"LUM_FM",                        "num"),
    (16, u"LUM_Quant_Luminaria_Calculado", "num"),
    (17, u"LUM_Quant_Luminaria_Adotado",   "num"),
    (18, u"LUM_Potencia_Unitaria",         "num"),
    (19, u"LUM_Carga_Total",               "num"),
]

# ── FUNÇÕES AUXILIARES ─────────────────────────────────────────────────────────

def get_rooms():
    rooms = []
    for r in FilteredElementCollector(doc).OfClass(SpatialElement).WhereElementIsNotElementType():
        try:
            if r.GetType().Name == "Room" and r.Area > 0:
                rooms.append(r)
        except Exception:
            pass
    return rooms


def get_rooms_by_guid():
    return dict((unicode(r.UniqueId), r) for r in get_rooms())


def safe_float(value):
    try:
        return float(unicode(value).strip().replace(u",", u"."))
    except Exception:
        return None


def garantir_shared_params():
    try:
        spf = app.OpenSharedParameterFile()
        if spf is not None:
            return True
    except Exception:
        pass

    res = MessageBox.Show(
        u"Este projeto nao tem arquivo de parametros compartilhados configurado.\n\n"
        u"SIM  — Localizar um arquivo .txt existente\n"
        u"NAO — Criar um novo arquivo .txt",
        u"Parametros Compartilhados — NnBim",
        MessageBoxButton.YesNoCancel,
        MessageBoxImage.Question
    )

    if res == System.Windows.MessageBoxResult.Cancel:
        return False

    if res == System.Windows.MessageBoxResult.Yes:
        dlg = OpenFileDialog()
        dlg.Title  = u"Selecione o arquivo de parametros compartilhados (.txt)"
        dlg.Filter = u"Arquivo de texto (*.txt)|*.txt"
        if dlg.ShowDialog() != DialogResult.OK:
            return False
        path = dlg.FileName
    else:
        dlg = SaveFileDialog()
        dlg.Title    = u"Salvar novo arquivo de parametros compartilhados"
        dlg.Filter   = u"Arquivo de texto (*.txt)|*.txt"
        dlg.FileName = u"NnBim_SharedParameters.txt"
        if dlg.ShowDialog() != DialogResult.OK:
            return False
        path = dlg.FileName
        with open(path, "w") as f:
            f.write("# This is a Revit shared parameter file.\n")
            f.write("# Do not edit manually.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")

    app.SharedParametersFilename = path
    try:
        spf = app.OpenSharedParameterFile()
        return spf is not None
    except Exception:
        return False


def criar_parametros_lum():
    if not garantir_shared_params():
        return False, u"Operacao cancelada pelo usuario."

    spf = app.OpenSharedParameterFile()
    if spf is None:
        return False, u"Nao foi possivel abrir o arquivo de parametros compartilhados."

    grupo = None
    for g in spf.Groups:
        if unicode(g.Name) == GRUPO_PARAM:
            grupo = g
            break
    if grupo is None:
        grupo = spf.Groups.Create(GRUPO_PARAM)

    cat_rooms = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms)
    cat_set   = app.Create.NewCategorySet()
    cat_set.Insert(cat_rooms)
    binding  = app.Create.NewInstanceBinding(cat_set)
    bind_map = doc.ParameterBindings

    nomes_existentes = set()
    it = bind_map.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        try:
            nomes_existentes.add(unicode(it.Key.Name))
        except Exception:
            pass

    criados = []
    with Transaction(doc, u"NnBim - Criar parametros LUM_*") as t:
        t.Start()
        try:
            for nome, forge_type, _ in PARAMS_LUM:
                if nome in nomes_existentes:
                    continue
                ext_def = None
                for d in grupo.Definitions:
                    if unicode(d.Name) == nome:
                        ext_def = d
                        break
                if ext_def is None:
                    opts = ExternalDefinitionCreationOptions(nome, forge_type)
                    opts.Visible = True
                    ext_def = grupo.Definitions.Create(opts)
                bind_map.Insert(ext_def, binding)
                criados.append(nome)
            t.Commit()
        except Exception as ex:
            t.RollBack()
            return False, u"Erro ao criar parametros: {}".format(unicode(ex))

    if criados:
        msg = u"{} parametro(s) criado(s):\n{}".format(len(criados), u"\n".join(criados))
    else:
        msg = u"Todos os parametros LUM_* ja existem no projeto."
    return True, msg


def get_abas_excel(path_xlsx):
    clr.AddReference("Microsoft.Office.Interop.Excel")
    import Microsoft.Office.Interop.Excel as xl
    import System.Runtime.InteropServices as interop

    excel = None
    wb    = None
    abas  = []
    try:
        excel = xl.ApplicationClass()
        excel.Visible       = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path_xlsx)
        for ws in wb.Sheets:
            abas.append(unicode(ws.Name))
    finally:
        if wb is not None:
            try:
                wb.Close(False)
                interop.Marshal.ReleaseComObject(wb)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
                interop.Marshal.ReleaseComObject(excel)
            except Exception:
                pass
    return abas


def ler_linhas_excel(path_xlsx, nome_aba):
    clr.AddReference("Microsoft.Office.Interop.Excel")
    import Microsoft.Office.Interop.Excel as xl
    import System.Runtime.InteropServices as interop

    xlUp  = -4162
    excel = None
    wb    = None
    try:
        excel = xl.ApplicationClass()
        excel.Visible       = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path_xlsx)

        ws = None
        for s in wb.Sheets:
            if unicode(s.Name) == nome_aba:
                ws = s
                break
        if ws is None:
            raise Exception(u"Aba '{}' nao encontrada.".format(nome_aba))

        last_row = ws.Cells(ws.Rows.Count, COL_GUID).End(xlUp).Row
        linhas = []
        for row in range(LINHA_INICIO, last_row + 1):
            valores = {}
            for col in range(1, NUM_COLS + 1):
                try:
                    valores[col] = ws.Cells(row, col).Value2
                except Exception:
                    valores[col] = None
            linhas.append(valores)
        return linhas
    finally:
        if wb is not None:
            try:
                wb.Close(False)
                interop.Marshal.ReleaseComObject(wb)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
                interop.Marshal.ReleaseComObject(excel)
            except Exception:
                pass


def aplicar_valores(linhas):
    rooms_by_guid   = get_rooms_by_guid()
    atualizados     = 0
    nao_encontrados = 0
    params_faltando = set()

    with Transaction(doc, u"NnBim - Importar Luminotecnico") as t:
        t.Start()
        for valores in linhas:
            guid = valores.get(COL_GUID)
            guid = unicode(guid).strip() if guid is not None else u""
            if not guid:
                continue

            room = rooms_by_guid.get(guid)
            if room is None:
                nao_encontrados += 1
                continue

            algo_gravado = False
            for col, nome_param, tipo in COL_PARAMS:
                val = valores.get(col)
                if val is None or unicode(val).strip() == u"":
                    continue

                p = room.LookupParameter(nome_param)
                if p is None or p.IsReadOnly:
                    params_faltando.add(nome_param)
                    continue

                if tipo == "txt":
                    p.Set(unicode(val))
                    algo_gravado = True
                else:
                    num = safe_float(val)
                    if num is None:
                        continue
                    if tipo == "len":
                        num = num / FT_TO_M
                    p.Set(num)
                    algo_gravado = True

            if algo_gravado:
                atualizados += 1

        t.Commit()

    return atualizados, nao_encontrados, params_faltando


# ── XAML ──────────────────────────────────────────────────────────────────────

XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Importar Luminotecnico"
    Width="560" Height="320"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    Background="#1B2736">
    <Window.Resources>
        <Style x:Key="FlatButton" TargetType="Button">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Opacity" Value="0.85"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,14">
            <TextBlock Text="Importar Luminotecnico"
                       Foreground="White" FontSize="15" FontWeight="Bold"/>
            <TextBlock Text="Le os valores calculados da planilha e grava nos Rooms do modelo (cruzamento por GUID)"
                       Foreground="#5A7A9A" FontSize="11" Margin="0,3,0,0" TextWrapping="Wrap"/>
        </StackPanel>

        <Grid Grid.Row="1" Margin="0,0,0,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Arquivo de calculo (.xlsx):"
                       Foreground="#90AABF" FontSize="11" Margin="0,0,0,5"/>
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="txtXlsx" Grid.Column="0"
                         Background="#243545" Foreground="#D0E0F0"
                         BorderBrush="#2E4A6A" BorderThickness="1"
                         Padding="6,4" Height="30" IsReadOnly="True"
                         Text="Nenhum arquivo selecionado" Margin="0,0,8,0"/>
                <Button x:Name="btnXlsx" Grid.Column="1"
                        Style="{StaticResource FlatButton}"
                        Content="Selecionar xlsx"
                        Background="#2E4A6A" Foreground="White"
                        Height="30" Width="130"/>
            </Grid>
        </Grid>

        <Grid Grid.Row="2" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Ler resultados da aba:"
                       Foreground="#90AABF" FontSize="11"
                       VerticalAlignment="Center" Margin="0,0,10,0"/>
            <ComboBox x:Name="cmbAbas" Grid.Column="1"
                      Background="#243545" Foreground="Black"
                      BorderBrush="#2E4A6A" Height="30"
                      IsEnabled="False">
                <ComboBox.ItemContainerStyle>
                    <Style TargetType="ComboBoxItem">
                        <Setter Property="Background" Value="#243545"/>
                        <Setter Property="Foreground" Value="#D0E0F0"/>
                        <Setter Property="Padding" Value="6,4"/>
                        <Style.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter Property="Background" Value="#2E4A6A"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </ComboBox.ItemContainerStyle>
            </ComboBox>
        </Grid>

        <TextBlock x:Name="txtStatus" Grid.Row="3"
                   Foreground="#6AF0A0" FontSize="11"
                   Margin="0,0,0,12" TextWrapping="Wrap"
                   VerticalAlignment="Top"
                   Text="" Visibility="Collapsed"/>

        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnCancelar" Grid.Column="1"
                    Style="{StaticResource FlatButton}"
                    Content="Cancelar"
                    Background="#3A4A5A" Foreground="White"
                    Height="32" Width="100"
                    Margin="0,0,10,0"/>
            <Button x:Name="btnImportar" Grid.Column="2"
                    Style="{StaticResource FlatButton}"
                    Content="Importar"
                    Background="#1A6AB0" Foreground="White"
                    FontWeight="Bold"
                    Height="32" Width="120"
                    IsEnabled="False"/>
        </Grid>
    </Grid>
</Window>"""


# ── CONTROLLER ────────────────────────────────────────────────────────────────

class ImportarLuminotecnicoWindow(object):

    def __init__(self):
        self.window    = XamlReader.Parse(XAML)
        self.path_xlsx = None

        self.txt_xlsx    = self.window.FindName("txtXlsx")
        self.btn_xlsx    = self.window.FindName("btnXlsx")
        self.cmb_abas    = self.window.FindName("cmbAbas")
        self.txt_status  = self.window.FindName("txtStatus")
        self.btn_importar = self.window.FindName("btnImportar")
        self.btn_cancelar = self.window.FindName("btnCancelar")

        self._bind_eventos()

    def _bind_eventos(self):
        self.btn_xlsx.Click             += self._selecionar_xlsx
        self.cmb_abas.SelectionChanged  += self._validar_botao
        self.btn_importar.Click         += self._importar
        self.btn_cancelar.Click         += lambda s, e: self.window.Close()

    def _selecionar_xlsx(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Title  = u"Selecione o arquivo de calculo xlsx"
        dlg.Filter = u"Planilha Excel (*.xlsx)|*.xlsx"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self.path_xlsx = dlg.FileName
        self.txt_xlsx.Text = self.path_xlsx
        self.cmb_abas.Items.Clear()
        self.cmb_abas.IsEnabled = False
        try:
            abas = get_abas_excel(self.path_xlsx)
            for aba in abas:
                self.cmb_abas.Items.Add(aba)
            self.cmb_abas.IsEnabled = True
            self._set_status(
                u"Arquivo carregado — {} abas encontradas.".format(len(abas)),
                ok=True
            )
        except Exception as ex:
            self._set_status(u"Erro ao ler abas: {}".format(unicode(ex)), ok=False)
        self._validar_botao(None, None)

    def _validar_botao(self, sender, e):
        tem_xlsx = self.path_xlsx is not None
        tem_aba  = self.cmb_abas.SelectedItem is not None
        self.btn_importar.IsEnabled = tem_xlsx and tem_aba

    def _set_status(self, msg, ok=True):
        self.txt_status.Text       = msg
        self.txt_status.Foreground = (
            Media.Brushes.LightGreen if ok else Media.Brushes.Tomato
        )
        self.txt_status.Visibility = System.Windows.Visibility.Visible

    def _importar(self, sender, e):
        nome_aba = unicode(self.cmb_abas.SelectedItem)

        self._set_status(u"Verificando parametros LUM_*...", ok=True)
        ok, msg = criar_parametros_lum()
        if not ok:
            self._set_status(msg, ok=False)
            return

        try:
            self._set_status(u"Lendo planilha...", ok=True)
            linhas = ler_linhas_excel(self.path_xlsx, nome_aba)
            self._set_status(u"Gravando valores nos Rooms...", ok=True)
            atualizados, nao_encontrados, params_faltando = aplicar_valores(linhas)
        except Exception as ex:
            self._set_status(u"Erro: {}".format(unicode(ex)), ok=False)
            return

        resumo = u"{} ambiente(s) atualizados.".format(atualizados)
        if nao_encontrados:
            resumo += u"\n{} linha(s) com GUID nao encontrado no modelo.".format(nao_encontrados)
        if params_faltando:
            resumo += u"\nParametros nao encontrados/somente leitura: {}".format(
                u", ".join(sorted(params_faltando))
            )

        self._set_status(resumo, ok=(atualizados > 0))
        MessageBox.Show(
            resumo,
            u"Importar Luminotecnico — Concluido",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )
        self.window.Close()

    def show(self):
        self.window.ShowDialog()


# ── ENTRY POINT ───────────────────────────────────────────────────────────────

if not doc:
    MessageBox.Show(u"Nenhum documento aberto.", u"Importar Luminotecnico")
else:
    ImportarLuminotecnicoWindow().show()
