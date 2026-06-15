# -*- coding: utf-8 -*-
__title__ = 'Lumin_Calculo\nExportar Ambientes Luminotecnico'
__author__ = 'NnBim Dev'
"""
Titulo: Lumin_Calculo - Exportar Ambientes Luminotecnico
Descricao: Exporta os Rooms do modelo Revit para a planilha de calculo
           luminotecnico, criando automaticamente os parametros LUM_*
           no grupo NnBim - LUMINOTECNICO do arquivo de parametros
           compartilhados. Preenche GUID, ID, AMBIENTE, AREA, PERIMETRO
           e ALTURA LUMINARIA na aba selecionada a partir da linha 5.
Instrucoes de Uso:
    1. Clique em Lumin_Calculo na aba NnBim
    2. Selecione os ambientes a exportar
    3. Indique o arquivo .txt de parametros compartilhados se solicitado
    4. Selecione o arquivo Excel de calculo
    5. Escolha a aba de destino no dropdown
    6. Clique em Exportar e salve o arquivo gerado
"""

import clr
import os

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
import System.Windows.Media as Media

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# ── CONSTANTES ────────────────────────────────────────────────────────────────

GRUPO_PARAM = u"NnBim - LUMINOTECNICO"
FT2_TO_M2   = 0.09290304
FT_TO_M     = 0.3048

CHARS_INVALIDOS_ABA = u"[]:*?/\\"
TAMANHO_MAX_ABA     = 31

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


def exportar_para_excel(path_xlsx, nome_aba, rooms_sel, path_saida, pending_ops=None):
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

    excel = None
    wb    = None
    try:
        excel = xl.ApplicationClass()
        excel.Visible       = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path_xlsx)

        if pending_ops:
            aplicar_operacoes_abas(wb, pending_ops)

        ws = None
        for s in wb.Sheets:
            if unicode(s.Name) == nome_aba:
                ws = s
                break
        if ws is None:
            raise Exception(u"Aba '{}' nao encontrada.".format(nome_aba))

        xlUp     = -4162
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

        wb.SaveAs(path_saida)
        return len(rooms_sel)

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
    Title="Lumin_Calculo - Exportar Ambientes"
    Width="660" Height="600"
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,12">
            <TextBlock Text="Lumin_Calculo — Exportar Ambientes"
                       Foreground="White" FontSize="15" FontWeight="Bold"/>
            <TextBlock Text="Exporta Rooms para a planilha de calculo luminotecnico"
                       Foreground="#5A7A9A" FontSize="11" Margin="0,3,0,0"/>
        </StackPanel>

        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0"
                       Text="Ambientes do modelo (selecione os que deseja exportar):"
                       Foreground="#90AABF" FontSize="11" Margin="0,0,0,5"/>
            <ListBox x:Name="lstRooms" Grid.Row="1"
                     Background="#243545" Foreground="#D0E0F0"
                     BorderBrush="#2E4A6A" BorderThickness="1"
                     ScrollViewer.VerticalScrollBarVisibility="Auto">
                <ListBox.ItemContainerStyle>
                    <Style TargetType="ListBoxItem">
                        <Setter Property="Background" Value="Transparent"/>
                        <Setter Property="Padding" Value="4,2"/>
                        <Style.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#2E4A6A"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#2E4A6A"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </ListBox.ItemContainerStyle>
            </ListBox>
            <Button x:Name="btnToggle" Grid.Row="2"
                    Style="{StaticResource FlatButton}"
                    Content="Selecionar todos"
                    Background="#2E4A6A" Foreground="White"
                    Height="28" Width="130"
                    HorizontalAlignment="Left" Margin="0,6,0,0"/>
        </Grid>

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

        <Grid Grid.Row="3" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Gravar na aba:"
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
            <Button x:Name="btnNovaAba" Grid.Column="2"
                    Style="{StaticResource FlatButton}"
                    Content="Nova aba"
                    Background="#2E4A6A" Foreground="White"
                    Height="30" Width="90"
                    Margin="8,0,0,0" IsEnabled="False"/>
            <Button x:Name="btnRenomearAba" Grid.Column="3"
                    Style="{StaticResource FlatButton}"
                    Content="Renomear"
                    Background="#2E4A6A" Foreground="White"
                    Height="30" Width="90"
                    Margin="8,0,0,0" IsEnabled="False"/>
        </Grid>

        <TextBlock x:Name="txtStatus" Grid.Row="4"
                   Foreground="#6AF0A0" FontSize="11"
                   Margin="0,0,0,12" TextWrapping="Wrap"
                   Text="" Visibility="Collapsed"/>

        <Grid Grid.Row="5">
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
    Title="Lumin_Calculo"
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

class LuminCalculoWindow(object):

    def __init__(self):
        self.window     = XamlReader.Parse(XAML)
        self.rooms      = get_rooms()
        self.path_xlsx  = None
        self.todos_sel  = False
        self.pending_ops = []

        self.lst             = self.window.FindName("lstRooms")
        self.btn_toggle      = self.window.FindName("btnToggle")
        self.txt_xlsx        = self.window.FindName("txtXlsx")
        self.btn_xlsx        = self.window.FindName("btnXlsx")
        self.cmb_abas        = self.window.FindName("cmbAbas")
        self.btn_nova_aba    = self.window.FindName("btnNovaAba")
        self.btn_renomear_aba = self.window.FindName("btnRenomearAba")
        self.txt_status      = self.window.FindName("txtStatus")
        self.btn_exportar    = self.window.FindName("btnExportar")
        self.btn_cancelar    = self.window.FindName("btnCancelar")

        self._popular_lista()
        self._bind_eventos()

    def _popular_lista(self):
        from System.Windows.Controls import CheckBox
        self.lst.Items.Clear()
        for room in self.rooms:
            guid, eid, nome, area, perim, alt = get_room_data(room)
            label = u"[{}]  {}  —  {} m\u00b2".format(eid, nome, area)
            cb = CheckBox()
            cb.Content    = label
            cb.Foreground = Media.Brushes.LightGray
            cb.Tag        = room
            cb.Margin     = System.Windows.Thickness(2)
            self.lst.Items.Add(cb)

    def _bind_eventos(self):
        self.btn_toggle.Click          += self._toggle_todos
        self.btn_xlsx.Click            += self._selecionar_xlsx
        self.cmb_abas.SelectionChanged += self._validar_botao
        self.btn_nova_aba.Click         += self._nova_aba
        self.btn_renomear_aba.Click     += self._renomear_aba
        self.btn_exportar.Click        += self._exportar
        self.btn_cancelar.Click        += lambda s, e: self.window.Close()
        for item in self.lst.Items:
            item.Checked   += self._validar_botao
            item.Unchecked += self._validar_botao

    def _toggle_todos(self, sender, e):
        self.todos_sel = not self.todos_sel
        for item in self.lst.Items:
            item.IsChecked = self.todos_sel
        self.btn_toggle.Content = (
            u"Desmarcar todos" if self.todos_sel else u"Selecionar todos"
        )
        self._validar_botao(None, None)

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
        self.pending_ops = []
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
        tem_room  = any(item.IsChecked for item in self.lst.Items)
        tem_xlsx  = self.path_xlsx is not None
        tem_aba   = self.cmb_abas.SelectedItem is not None
        self.btn_exportar.IsEnabled     = tem_room and tem_xlsx and tem_aba
        self.btn_nova_aba.IsEnabled     = tem_aba
        self.btn_renomear_aba.IsEnabled = tem_aba

    def _nomes_abas_atuais(self):
        return [unicode(self.cmb_abas.Items[i]) for i in range(self.cmb_abas.Items.Count)]

    def _validar_nome_aba(self, nome, nome_atual=None):
        nome = nome.strip()
        if not nome:
            return False, u"Nome da aba nao pode ser vazio."
        if len(nome) > TAMANHO_MAX_ABA:
            return False, u"Nome da aba deve ter no maximo {} caracteres.".format(TAMANHO_MAX_ABA)
        for ch in CHARS_INVALIDOS_ABA:
            if ch in nome:
                return False, u"Nome da aba contem caractere invalido: {}".format(ch)
        existentes = self._nomes_abas_atuais()
        if nome_atual is not None and nome_atual in existentes:
            existentes.remove(nome_atual)
        if nome in existentes:
            return False, u"Ja existe uma aba com esse nome."
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

    def _nova_aba(self, sender, e):
        if self.cmb_abas.SelectedItem is None:
            return
        origem = unicode(self.cmb_abas.SelectedItem)
        novo_nome = self._pedir_nome_aba(
            u"Nome da nova aba (copia de '{}'):".format(origem), u""
        )
        if not novo_nome:
            return
        ok, resultado = self._validar_nome_aba(novo_nome)
        if not ok:
            self._set_status(resultado, ok=False)
            return
        self.cmb_abas.Items.Add(resultado)
        self.cmb_abas.SelectedItem = resultado
        self.pending_ops.append({"tipo": "duplicar", "origem": origem, "novo_nome": resultado})
        self._set_status(
            u"Nova aba '{}' (copia de '{}') sera criada ao exportar.".format(resultado, origem),
            ok=True
        )

    def _renomear_aba(self, sender, e):
        if self.cmb_abas.SelectedItem is None:
            return
        atual = unicode(self.cmb_abas.SelectedItem)
        novo_nome = self._pedir_nome_aba(
            u"Novo nome para a aba '{}':".format(atual), atual
        )
        if not novo_nome or novo_nome == atual:
            return
        ok, resultado = self._validar_nome_aba(novo_nome, nome_atual=atual)
        if not ok:
            self._set_status(resultado, ok=False)
            return

        op_existente = None
        for op in self.pending_ops:
            if op["novo_nome"] == atual:
                op_existente = op
                break

        if op_existente is not None:
            op_existente["novo_nome"] = resultado
        else:
            self.pending_ops.append({"tipo": "renomear", "origem": atual, "novo_nome": resultado})

        idx = self.cmb_abas.Items.IndexOf(atual)
        self.cmb_abas.Items[idx] = resultado
        self.cmb_abas.SelectedItem = resultado
        self._set_status(
            u"Aba '{}' sera renomeada para '{}' ao exportar.".format(atual, resultado),
            ok=True
        )

    def _set_status(self, msg, ok=True):
        self.txt_status.Text       = msg
        self.txt_status.Foreground = (
            Media.Brushes.LightGreen if ok else Media.Brushes.Tomato
        )
        self.txt_status.Visibility = System.Windows.Visibility.Visible

    def _exportar(self, sender, e):
        rooms_sel = [item.Tag for item in self.lst.Items if item.IsChecked]
        if not rooms_sel:
            self._set_status(u"Selecione ao menos um ambiente.", ok=False)
            return

        nome_aba = unicode(self.cmb_abas.SelectedItem)

        self._set_status(u"Verificando parametros LUM_*...", ok=True)
        ok, msg = criar_parametros_lum()
        if not ok:
            self._set_status(msg, ok=False)
            return

        dlg = SaveFileDialog()
        dlg.Title    = u"Salvar planilha exportada"
        dlg.Filter   = u"Planilha Excel (*.xlsx)|*.xlsx"
        dlg.FileName = u"LUM_AMBIENTES.xlsx"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        path_saida = dlg.FileName

        try:
            self._set_status(u"Exportando...", ok=True)
            total    = exportar_para_excel(
                self.path_xlsx, nome_aba, rooms_sel, path_saida, self.pending_ops
            )
            nome_arq = os.path.basename(path_saida)
            self._set_status(
                u"{} ambientes exportados para {}".format(total, nome_arq),
                ok=True
            )
            MessageBox.Show(
                u"{} ambientes exportados com sucesso!\n\nArquivo:\n{}".format(
                    total, path_saida
                ),
                u"Lumin_Calculo — Concluido",
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
    MessageBox.Show(u"Nenhum documento aberto.", u"Lumin_Calculo")
else:
    _rooms = get_rooms()
    if not _rooms:
        MessageBox.Show(
            u"Nenhum ambiente com area > 0 encontrado no modelo.",
            u"Lumin_Calculo",
            MessageBoxButton.OK,
            MessageBoxImage.Warning
        )
    else:
        LuminCalculoWindow().show()