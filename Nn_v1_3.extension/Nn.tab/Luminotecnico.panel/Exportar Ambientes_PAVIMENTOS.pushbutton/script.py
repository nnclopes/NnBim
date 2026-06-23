# -*- coding: utf-8 -*-
__title__ = 'Exportar\nAmbientes'
__author__ = 'NnBim Dev'
"""
Titulo: Exportar Ambientes - Exportar Ambientes por Pavimento (TESTE)
Descricao: Versao experimental do Lumin_Calculo que agrupa os Rooms do
           modelo por pavimento (Level). Cada pavimento tem sua propria
           caixa "selecionar todos", seu proprio dropdown de aba de
           destino e seus proprios botoes Nova aba/Renomear, permitindo
           exportar varios pavimentos de uma vez para abas diferentes do
           mesmo arquivo Excel.
Instrucoes de Uso:
    1. Clique em LUM_CAL_PAVIMENTOS na aba NnBim
    2. Para cada pavimento, marque os ambientes a exportar
    3. Selecione o arquivo Excel de calculo
    4. Para cada pavimento com ambientes marcados, escolha a aba de destino
    5. Use "Nova aba" / "Renomear" se precisar criar ou renomear abas
    6. Clique em Exportar e salve o arquivo gerado
"""

import clr
import os
import shutil

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SpatialElement,
    BuiltInParameter,
    BuiltInCategory,
    ExternalDefinitionCreationOptions,
    InstanceBinding,
    Transaction,
    ForgeTypeId,
)

import System
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Windows.Markup import XamlReader
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows import Thickness, FontWeights, GridLength, GridUnitType
from System.Windows.Controls import (
    StackPanel, CheckBox, ComboBox, Button, Grid as WGrid,
    ColumnDefinition,
)
import System.Windows.Media as Media

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# ── CONSTANTES ────────────────────────────────────────────────────────────────

GRUPO_PARAM = u"NnBim - LUMINOTECNICO"
FT2_TO_M2   = 0.09290304
FT_TO_M     = 0.3048

CHARS_INVALIDOS_ABA = u"[]:*?/\\"
TAMANHO_MAX_ABA     = 31

TEMPLATE_XLSX = os.path.join(os.path.dirname(__file__), u"BASE_CALCULO_LUMINOTECNICO_RVT.xlsx")

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

# ── FUNÇÕES AUXILIARES ─────────────────────────────────────────────────────────

def get_rooms():
    rooms = []
    for r in FilteredElementCollector(doc).OfClass(SpatialElement).WhereElementIsNotElementType():
        try:
            if r.GetType().Name == "Room" and r.Area > 0:
                rooms.append(r)
        except Exception:
            pass
    rooms.sort(key=lambda r: unicode(
        r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or u""
    ))
    return rooms


def get_rooms_agrupados():
    """Retorna lista [(nome_nivel, [rooms]), ...] ordenada pela elevacao do nivel."""
    grupos = {}
    for room in get_rooms():
        nivel_nome = u"(Sem nivel)"
        elevacao   = 0.0
        try:
            nivel = doc.GetElement(room.LevelId)
            if nivel is not None:
                nivel_nome = unicode(nivel.Name)
                elevacao   = nivel.Elevation
        except Exception:
            pass
        if nivel_nome not in grupos:
            grupos[nivel_nome] = {"elevacao": elevacao, "rooms": []}
        grupos[nivel_nome]["rooms"].append(room)

    ordenado = sorted(grupos.items(), key=lambda kv: kv[1]["elevacao"])
    return [(nome, dados["rooms"]) for nome, dados in ordenado]


def get_room_data(room):
    guid  = unicode(room.UniqueId)
    eid   = unicode(room.Id.IntegerValue)
    nome  = unicode(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or u"")
    area  = round(room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()      * FT2_TO_M2, 2)
    perim = round(room.get_Parameter(BuiltInParameter.ROOM_PERIMETER).AsDouble() * FT_TO_M,   2)
    alt   = round(room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble()    * FT_TO_M,   2)
    return (guid, eid, nome, area, perim, alt)


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


def aplicar_operacoes_abas(wb, pending_ops):
    for op in pending_ops:
        if op["tipo"] == "duplicar":
            origem_ws = None
            for s in wb.Sheets:
                if unicode(s.Name) == op["origem"]:
                    origem_ws = s
                    break
            if origem_ws is not None:
                origem_ws.Copy(System.Type.Missing, origem_ws)
                wb.ActiveSheet.Name = op["novo_nome"]
        elif op["tipo"] == "renomear":
            for s in wb.Sheets:
                if unicode(s.Name) == op["origem"]:
                    s.Name = op["novo_nome"]
                    break


def exportar_para_excel_multi(path_xlsx, jobs, path_saida):
    """jobs: lista de (nome_aba, rooms_sel, pending_ops). Tudo gravado no
    mesmo arquivo de saida, cada job em sua propria aba.

    O arquivo de calculo (path_xlsx) e somente lido: uma copia dele e
    criada em path_saida e toda a edicao ocorre sobre essa copia, para
    que o arquivo base nunca seja sobrescrito."""
    clr.AddReference("Microsoft.Office.Interop.Excel")
    import Microsoft.Office.Interop.Excel as xl
    import System.Runtime.InteropServices as interop

    LINHA_INICIO = 5
    COL_GUID     = 1
    COL_ID       = 2
    COL_NOME     = 3
    COL_AREA     = 5
    COL_PERIM    = 6
    COL_ALT      = 9
    xlUp         = -4162

    if os.path.normcase(os.path.abspath(path_saida)) == os.path.normcase(os.path.abspath(path_xlsx)):
        raise Exception(u"O arquivo de saida nao pode ser o mesmo arquivo de calculo (base).")

    shutil.copy2(path_xlsx, path_saida)

    excel = None
    wb    = None
    total_rooms = 0
    try:
        excel = xl.ApplicationClass()
        excel.Visible       = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path_saida)

        for nome_aba, rooms_sel, pending_ops in jobs:
            if pending_ops:
                aplicar_operacoes_abas(wb, pending_ops)

            ws = None
            for s in wb.Sheets:
                if unicode(s.Name) == nome_aba:
                    ws = s
                    break
            if ws is None:
                raise Exception(u"Aba '{}' nao encontrada.".format(nome_aba))

            last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            if last_row >= LINHA_INICIO:
                ws.Range(
                    ws.Cells(LINHA_INICIO, 1),
                    ws.Cells(last_row, 22)
                ).ClearContents()

            for i, room in enumerate(rooms_sel):
                row = LINHA_INICIO + i
                guid, eid, nome, area, perim, alt = get_room_data(room)
                ws.Cells(row, COL_GUID).Value2  = guid
                ws.Cells(row, COL_ID).Value2    = eid
                ws.Cells(row, COL_NOME).Value2  = nome
                ws.Cells(row, COL_AREA).Value2  = area
                ws.Cells(row, COL_PERIM).Value2 = perim
                ws.Cells(row, COL_ALT).Value2   = alt

            total_rooms += len(rooms_sel)

        wb.Save()
        return total_rooms

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


# ── XAML ──────────────────────────────────────────────────────────────────────

XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="LUM_CAL_PAVIMENTOS - Exportar Ambientes por Pavimento (TESTE)"
    Width="760" Height="640"
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
        <Style x:Key="ComboAbaItemStyle" TargetType="ComboBoxItem">
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
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,12">
            <TextBlock Text="LUM_CAL_PAVIMENTOS — Exportar Ambientes por Pavimento"
                       Foreground="White" FontSize="15" FontWeight="Bold"/>
            <TextBlock Text="Versao de teste: cada pavimento pode ser exportado para uma aba diferente"
                       Foreground="#5A7A9A" FontSize="11" Margin="0,3,0,0"/>
        </StackPanel>

        <Border Grid.Row="1" Margin="0,0,0,10"
                Background="#243545" BorderBrush="#2E4A6A" BorderThickness="1">
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="10">
                <StackPanel x:Name="spGrupos"/>
            </ScrollViewer>
        </Border>

        <Grid Grid.Row="2" Margin="0,0,0,8">
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

        <TextBlock x:Name="txtStatus" Grid.Row="3"
                   Foreground="#6AF0A0" FontSize="11"
                   Margin="0,0,0,12" TextWrapping="Wrap"
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
            <Button x:Name="btnExportar" Grid.Column="2"
                    Style="{StaticResource FlatButton}"
                    Content="Exportar"
                    Background="#1A6AB0" Foreground="White"
                    FontWeight="Bold"
                    Height="32" Width="120"
                    IsEnabled="False"/>
        </Grid>
    </Grid>
</Window>"""


XAML_INPUT = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="LUM_CAL_PAVIMENTOS"
    Width="420" Height="170"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    Background="#1B2736">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtPrompt" Grid.Row="0"
                   Foreground="#D0E0F0" FontSize="12"
                   TextWrapping="Wrap" Margin="0,0,0,12"/>
        <TextBox x:Name="txtValor" Grid.Row="1"
                 Background="#243545" Foreground="#D0E0F0"
                 BorderBrush="#2E4A6A" BorderThickness="1"
                 Padding="6,4" Height="30"/>
        <Grid Grid.Row="3" Margin="0,16,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnCancelar" Grid.Column="1"
                    Content="Cancelar"
                    Background="#3A4A5A" Foreground="White"
                    BorderThickness="0" Height="30" Width="90"
                    Margin="0,0,10,0" Cursor="Hand"/>
            <Button x:Name="btnOk" Grid.Column="2"
                    Content="OK"
                    Background="#1A6AB0" Foreground="White"
                    FontWeight="Bold" BorderThickness="0"
                    Height="30" Width="90" Cursor="Hand"/>
        </Grid>
    </Grid>
</Window>"""


# ── CONTROLLER ────────────────────────────────────────────────────────────────

class LumCalPavimentosWindow(object):

    def __init__(self):
        self.window    = XamlReader.Parse(XAML)
        self.path_xlsx = None
        self.grupos    = []

        self.sp_grupos    = self.window.FindName("spGrupos")
        self.txt_xlsx     = self.window.FindName("txtXlsx")
        self.btn_xlsx     = self.window.FindName("btnXlsx")
        self.txt_status   = self.window.FindName("txtStatus")
        self.btn_exportar = self.window.FindName("btnExportar")
        self.btn_cancelar = self.window.FindName("btnCancelar")

        self._popular_grupos()
        self._bind_eventos()

    # ── construcao da UI dinamica ──────────────────────────────────────────

    def _popular_grupos(self):
        self.sp_grupos.Children.Clear()
        self.grupos = []

        for nivel_nome, rooms in get_rooms_agrupados():
            grupo = {"nivel": nivel_nome, "rooms": rooms, "pending_ops": []}

            outer = StackPanel()
            outer.Margin = Thickness(0, 0, 0, 16)

            header = WGrid()
            header.Margin = Thickness(0, 0, 0, 4)
            for w in (GridLength(1, GridUnitType.Star), GridLength.Auto, GridLength.Auto, GridLength.Auto):
                cd = ColumnDefinition()
                cd.Width = w
                header.ColumnDefinitions.Add(cd)

            cb_all = CheckBox()
            cb_all.Content    = nivel_nome
            cb_all.Foreground = Media.Brushes.White
            cb_all.FontWeight = FontWeights.Bold
            cb_all.VerticalAlignment = System.Windows.VerticalAlignment.Center
            WGrid.SetColumn(cb_all, 0)
            header.Children.Add(cb_all)

            cmb_aba = ComboBox()
            cmb_aba.Width      = 200
            cmb_aba.Height     = 28
            cmb_aba.Background = Media.SolidColorBrush(Media.ColorConverter.ConvertFromString("#243545"))
            cmb_aba.Foreground = Media.Brushes.Black
            cmb_aba.BorderBrush = Media.SolidColorBrush(Media.ColorConverter.ConvertFromString("#2E4A6A"))
            cmb_aba.IsEnabled  = False
            cmb_aba.ItemContainerStyle = self.window.FindResource("ComboAbaItemStyle")
            cmb_aba.Margin    = Thickness(8, 0, 0, 0)
            WGrid.SetColumn(cmb_aba, 1)
            header.Children.Add(cmb_aba)

            btn_nova = Button()
            btn_nova.Content    = u"Nova aba"
            btn_nova.Style      = self.window.FindResource("FlatButton")
            btn_nova.Background = Media.SolidColorBrush(Media.ColorConverter.ConvertFromString("#2E4A6A"))
            btn_nova.Foreground = Media.Brushes.White
            btn_nova.Width      = 80
            btn_nova.Height     = 28
            btn_nova.Margin     = Thickness(8, 0, 0, 0)
            btn_nova.IsEnabled  = False
            WGrid.SetColumn(btn_nova, 2)
            header.Children.Add(btn_nova)

            btn_renomear = Button()
            btn_renomear.Content    = u"Renomear"
            btn_renomear.Style      = self.window.FindResource("FlatButton")
            btn_renomear.Background = Media.SolidColorBrush(Media.ColorConverter.ConvertFromString("#2E4A6A"))
            btn_renomear.Foreground = Media.Brushes.White
            btn_renomear.Width      = 80
            btn_renomear.Height     = 28
            btn_renomear.Margin     = Thickness(8, 0, 0, 0)
            btn_renomear.IsEnabled  = False
            WGrid.SetColumn(btn_renomear, 3)
            header.Children.Add(btn_renomear)

            outer.Children.Add(header)

            rooms_panel = StackPanel()
            rooms_panel.Margin = Thickness(24, 0, 0, 0)
            items = []
            for room in rooms:
                guid, eid, nome, area, perim, alt = get_room_data(room)
                label = u"[{}]  {}  —  {} m²".format(eid, nome, area)
                cb = CheckBox()
                cb.Content    = label
                cb.Foreground = Media.Brushes.LightGray
                cb.Tag        = room
                cb.Margin     = Thickness(2)
                rooms_panel.Children.Add(cb)
                items.append(cb)

            outer.Children.Add(rooms_panel)
            self.sp_grupos.Children.Add(outer)

            grupo["cb_all"]       = cb_all
            grupo["cmb_aba"]      = cmb_aba
            grupo["btn_nova"]     = btn_nova
            grupo["btn_renomear"] = btn_renomear
            grupo["items"]        = items
            self.grupos.append(grupo)

    # ── eventos ─────────────────────────────────────────────────────────────

    def _bind_eventos(self):
        self.btn_xlsx.Click     += self._selecionar_xlsx
        self.btn_exportar.Click += self._exportar
        self.btn_cancelar.Click += lambda s, e: self.window.Close()

        for grupo in self.grupos:
            grupo["cb_all"].Checked          += self._make_toggle_grupo(grupo)
            grupo["cb_all"].Unchecked        += self._make_toggle_grupo(grupo)
            grupo["cmb_aba"].SelectionChanged += self._validar_botao
            grupo["btn_nova"].Click           += self._make_nova_aba(grupo)
            grupo["btn_renomear"].Click       += self._make_renomear_aba(grupo)
            for cb in grupo["items"]:
                cb.Checked   += self._validar_botao
                cb.Unchecked += self._validar_botao

    def _make_toggle_grupo(self, grupo):
        def handler(sender, e):
            estado = grupo["cb_all"].IsChecked
            for cb in grupo["items"]:
                cb.IsChecked = estado
            self._validar_botao(None, None)
        return handler

    def _selecionar_xlsx(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Title  = u"Selecione o arquivo de calculo xlsx"
        dlg.Filter = u"Planilha Excel (*.xlsx)|*.xlsx"
        if os.path.isfile(TEMPLATE_XLSX):
            dlg.InitialDirectory = os.path.dirname(TEMPLATE_XLSX)
            dlg.FileName         = os.path.basename(TEMPLATE_XLSX)
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self.path_xlsx = dlg.FileName
        self.txt_xlsx.Text = self.path_xlsx
        try:
            abas = get_abas_excel(self.path_xlsx)
            for grupo in self.grupos:
                cmb = grupo["cmb_aba"]
                cmb.Items.Clear()
                for aba in abas:
                    cmb.Items.Add(aba)
                cmb.IsEnabled = True
                grupo["pending_ops"] = []
            self._set_status(
                u"Arquivo carregado — {} abas encontradas.".format(len(abas)),
                ok=True
            )
        except Exception as ex:
            self._set_status(u"Erro ao ler abas: {}".format(unicode(ex)), ok=False)
        self._validar_botao(None, None)

    def _validar_botao(self, sender, e):
        tem_xlsx = self.path_xlsx is not None
        algum_valido = False
        abas_destino = []

        for grupo in self.grupos:
            tem_room = any(cb.IsChecked for cb in grupo["items"])
            tem_aba  = grupo["cmb_aba"].SelectedItem is not None
            grupo["btn_nova"].IsEnabled     = tem_aba
            grupo["btn_renomear"].IsEnabled = tem_aba
            if tem_room and tem_aba:
                algum_valido = True
                abas_destino.append(unicode(grupo["cmb_aba"].SelectedItem))

        sem_duplicidade = len(abas_destino) == len(set(abas_destino))
        self.btn_exportar.IsEnabled = tem_xlsx and algum_valido and sem_duplicidade

        if algum_valido and not sem_duplicidade:
            self._set_status(
                u"Dois pavimentos nao podem exportar para a mesma aba.",
                ok=False
            )

    # ── nomes/validacao de abas (por grupo) ────────────────────────────────

    def _nomes_abas_grupo(self, cmb):
        return [unicode(cmb.Items[i]) for i in range(cmb.Items.Count)]

    def _validar_nome_aba(self, cmb, nome, nome_atual=None):
        nome = nome.strip()
        if not nome:
            return False, u"Nome da aba nao pode ser vazio."
        if len(nome) > TAMANHO_MAX_ABA:
            return False, u"Nome da aba deve ter no maximo {} caracteres.".format(TAMANHO_MAX_ABA)
        for ch in CHARS_INVALIDOS_ABA:
            if ch in nome:
                return False, u"Nome da aba contem caractere invalido: {}".format(ch)
        existentes = self._nomes_abas_grupo(cmb)
        if nome_atual is not None and nome_atual in existentes:
            existentes.remove(nome_atual)
        if nome in existentes:
            return False, u"Ja existe uma aba com esse nome neste pavimento."
        return True, nome

    def _pedir_nome_aba(self, prompt, default):
        win          = XamlReader.Parse(XAML_INPUT)
        txt_prompt   = win.FindName("txtPrompt")
        txt_valor    = win.FindName("txtValor")
        btn_ok       = win.FindName("btnOk")
        btn_cancelar = win.FindName("btnCancelar")

        txt_prompt.Text = prompt
        txt_valor.Text  = default
        txt_valor.SelectAll()

        resultado = {"valor": u""}

        def confirmar(s, e):
            resultado["valor"] = txt_valor.Text
            win.Close()

        btn_ok.Click       += confirmar
        btn_cancelar.Click += lambda s, e: win.Close()
        win.Loaded         += lambda s, e: txt_valor.Focus()
        win.ShowDialog()

        return unicode(resultado["valor"]).strip()

    def _make_nova_aba(self, grupo):
        def handler(sender, e):
            cmb = grupo["cmb_aba"]
            if cmb.SelectedItem is None:
                return
            origem = unicode(cmb.SelectedItem)
            novo_nome = self._pedir_nome_aba(
                u"[{}] Nome da nova aba (copia de '{}'):".format(grupo["nivel"], origem), u""
            )
            if not novo_nome:
                return
            ok, resultado = self._validar_nome_aba(cmb, novo_nome)
            if not ok:
                self._set_status(resultado, ok=False)
                return
            cmb.Items.Add(resultado)
            cmb.SelectedItem = resultado
            grupo["pending_ops"].append({"tipo": "duplicar", "origem": origem, "novo_nome": resultado})
            self._set_status(
                u"[{}] Nova aba '{}' (copia de '{}') sera criada ao exportar.".format(
                    grupo["nivel"], resultado, origem
                ),
                ok=True
            )
        return handler

    def _make_renomear_aba(self, grupo):
        def handler(sender, e):
            cmb = grupo["cmb_aba"]
            if cmb.SelectedItem is None:
                return
            atual = unicode(cmb.SelectedItem)
            novo_nome = self._pedir_nome_aba(
                u"[{}] Novo nome para a aba '{}':".format(grupo["nivel"], atual), atual
            )
            if not novo_nome or novo_nome == atual:
                return
            ok, resultado = self._validar_nome_aba(cmb, novo_nome, nome_atual=atual)
            if not ok:
                self._set_status(resultado, ok=False)
                return

            op_existente = None
            for op in grupo["pending_ops"]:
                if op["novo_nome"] == atual:
                    op_existente = op
                    break

            if op_existente is not None:
                op_existente["novo_nome"] = resultado
            else:
                grupo["pending_ops"].append({"tipo": "renomear", "origem": atual, "novo_nome": resultado})

            idx = cmb.Items.IndexOf(atual)
            cmb.Items[idx] = resultado
            cmb.SelectedItem = resultado
            self._set_status(
                u"[{}] Aba '{}' sera renomeada para '{}' ao exportar.".format(
                    grupo["nivel"], atual, resultado
                ),
                ok=True
            )
        return handler

    def _set_status(self, msg, ok=True):
        self.txt_status.Text       = msg
        self.txt_status.Foreground = (
            Media.Brushes.LightGreen if ok else Media.Brushes.Tomato
        )
        self.txt_status.Visibility = System.Windows.Visibility.Visible

    # ── exportacao ──────────────────────────────────────────────────────────

    def _exportar(self, sender, e):
        jobs = []
        for grupo in self.grupos:
            rooms_sel = [cb.Tag for cb in grupo["items"] if cb.IsChecked]
            if not rooms_sel or grupo["cmb_aba"].SelectedItem is None:
                continue
            nome_aba = unicode(grupo["cmb_aba"].SelectedItem)
            jobs.append((nome_aba, rooms_sel, grupo["pending_ops"]))

        if not jobs:
            self._set_status(u"Selecione ao menos um ambiente e uma aba de destino.", ok=False)
            return

        self._set_status(u"Verificando parametros LUM_*...", ok=True)
        ok, msg = criar_parametros_lum()
        if not ok:
            self._set_status(msg, ok=False)
            return

        dlg = SaveFileDialog()
        dlg.Title    = u"Salvar planilha exportada"
        dlg.Filter   = u"Planilha Excel (*.xlsx)|*.xlsx"
        dlg.FileName = u"LUM_AMBIENTES_PAVIMENTOS.xlsx"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        path_saida = dlg.FileName

        try:
            self._set_status(u"Exportando...", ok=True)
            total = exportar_para_excel_multi(self.path_xlsx, jobs, path_saida)
            nome_arq = os.path.basename(path_saida)
            self._set_status(
                u"{} ambientes exportados ({} pavimento(s)) para {}".format(
                    total, len(jobs), nome_arq
                ),
                ok=True
            )
            MessageBox.Show(
                u"{} ambientes exportados com sucesso em {} aba(s)!\n\nArquivo:\n{}".format(
                    total, len(jobs), path_saida
                ),
                u"LUM_CAL_PAVIMENTOS — Concluido",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
            self.window.Close()
        except Exception as ex:
            self._set_status(u"Erro: {}".format(unicode(ex)), ok=False)

    def show(self):
        self.window.ShowDialog()


# ── ENTRY POINT ───────────────────────────────────────────────────────────────

if not doc:
    MessageBox.Show(u"Nenhum documento aberto.", u"LUM_CAL_PAVIMENTOS")
else:
    _rooms = get_rooms()
    if not _rooms:
        MessageBox.Show(
            u"Nenhum ambiente com area > 0 encontrado no modelo.",
            u"LUM_CAL_PAVIMENTOS",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        )
    else:
        LumCalPavimentosWindow().show()
