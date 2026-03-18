# -*- coding: utf-8 -*-
"""
NnBim: EXCEL LINK PRO
DESCRICAO:
Sincronização bidirecional entre Revit e Excel padrão DiRoots. 
Permite exportar tabelas com UniqueId e mapeamento de cores.

COMO USAR:
1. Escolha a aba desejada (Tabelas ou Categorias).
2. Exporte os dados. Parâmetros editáveis ficarão em VERDE.
3. Altere o Excel e use a função IMPORTAR para atualizar o Revit.
"""
__title__ = 'Excel\nLink'
__author__ = 'Nn_Dev'

from pyrevit import revit, DB, forms, script
import os
import sys
import clr

# Referências Externas (Excel e Cores do Windows)
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

clr.AddReference("System.Drawing")
import System.Drawing

# ---------------------------------------------------------
# IMPORTAÇÃO SEGURA DA SUA BIBLIOTECA (WPF_Base)
# ---------------------------------------------------------
lib_path = r'C:\Users\Nn_1tb\OneDrive\Documentos\GitHub\NnBim\Nn_v1_3.extension\lib'
if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    from GUI.WPF_Base import my_WPF as WPFWindow
except ImportError as e:
    forms.alert("Erro ao carregar a biblioteca de interface:\n{}".format(e), title="Erro NnBim")
    sys.exit()

doc = revit.doc
uidoc = revit.uidoc

class ExcelLinkWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        import wpf
        wpf.LoadComponent(self, xaml_file_name)
        self._setup_initial_data()

    def _setup_initial_data(self):
        """Preenche as tabelas e categorias ao abrir a janela."""
        schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
        self.list_schedules.ItemsSource = sorted([s.Name for s in schedules if not s.IsTitleblockRevisionSchedule])
        
        categories = doc.Settings.Categories
        valid_cats = sorted([c.Name for c in categories if c.AllowsBoundParameters])
        self.cb_categories.ItemsSource = valid_cats

    # ---------------------------------------------------------
    # EVENTOS DA INTERFACE (ABA 2 - CATEGORIAS)
    # ---------------------------------------------------------
    def on_category_changed(self, sender, e):
        """Busca TODOS os parâmetros (Instância, Tipo, Compartilhados) da categoria."""
        self.list_available_params.Items.Clear()
        self.list_selected_params.Items.Clear()
        
        cat_name = self.cb_categories.SelectedItem
        if not cat_name: return

        # Encontrar a Categoria (BuiltInCategory)
        target_cat = next((c for c in doc.Settings.Categories if c.Name == cat_name), None)
        if not target_cat: return

        # Garimpo: Pegar um elemento dessa categoria para extrair seus parâmetros
        collector = DB.FilteredElementCollector(doc).OfCategoryId(target_cat.Id).WhereElementIsNotElementType()
        el = collector.FirstElement()
        
        param_names = set()
        
        if el:
            # 1. Parâmetros de Instância (Inclui os Compartilhados atrelados à instância)
            for p in el.Parameters:
                param_names.add(p.Definition.Name)
            
            # 2. Parâmetros de Tipo (Inclui os Compartilhados atrelados ao tipo)
            el_type = doc.GetElement(el.GetTypeId())
            if el_type:
                for p in el_type.Parameters:
                    param_names.add(p.Definition.Name)
        else:
            # Se não houver nenhum elemento modelado, varre o mapa de ligações (Bindings)
            iterator = doc.ParameterBindings.ForwardIterator()
            while iterator.MoveNext():
                if target_cat in iterator.Current.Categories:
                    param_names.add(iterator.Key.Name)

        # Adiciona na interface em ordem alfabética
        for p_name in sorted(param_names):
            self.list_available_params.Items.Add(p_name)

    def on_add_param(self, sender, e):
        selected = self.list_available_params.SelectedItem
        if selected and not self.list_selected_params.Items.Contains(selected):
            self.list_selected_params.Items.Add(selected)

    def on_remove_param(self, sender, e):
        selected = self.list_selected_params.SelectedItem
        if selected:
            self.list_selected_params.Items.Remove(selected)

    # ---------------------------------------------------------
    # AÇÕES PRINCIPAIS (EXPORTAR / IMPORTAR)
    # ---------------------------------------------------------
    def on_export_click(self, sender, e):
        selected_tab = self.tab_main.SelectedIndex
        
        if selected_tab == 0:
            selected_schedules = list(self.list_schedules.SelectedItems)
            if not selected_schedules:
                forms.alert("Selecione pelo menos uma tabela.")
                return
            
            # Transação temporária para forçar itemização
            with DB.Transaction(doc, "NnBim: Prep Export") as t:
                t.Start()
                for sched_name in selected_schedules:
                    sched = self._get_schedule_by_name(sched_name)
                    
                    # Força a tabela a mostrar elemento por elemento
                    sched.Definition.IsItemized = True
                    doc.Regenerate() # Atualiza o modelo para o Revit ler as linhas novas
                    
                    self._export_schedule_to_excel(sched)
                t.RollBack() # Desfaz a alteração para não estragar a tabela da Nívea!
                
        else:
            forms.alert("Exportação por categoria na próxima atualização!", title="NnBim Dev")
        self.Close()

    def on_import_click(self, sender, e):
        file_path = forms.pick_file(file_ext='xlsx')
        if file_path:
            forms.alert("Arquivo {} selecionado para importação!", title="NnBim Dev")
        self.Close()

    # ---------------------------------------------------------
    # MOTOR DE EXPORTAÇÃO (PADRÃO DIROOTS)
    # ---------------------------------------------------------
    def _get_schedule_by_name(self, name):
        schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
        return next((s for s in schedules if s.Name == name), None)

    def _export_schedule_to_excel(self, schedule):
        """Lê os dados e a editabilidade da tabela e envia ao Excel com Cores e UniqueId"""
        try:
            app = Excel.ApplicationClass()
            app.Visible = True
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Name = (schedule.Name[:30] + '..') if len(schedule.Name) > 31 else schedule.Name
            
            # 1. Obter a definição das colunas do Revit
            definition = schedule.Definition
            n_cols = definition.GetFieldCount()
            fields = [definition.GetField(i) for i in range(n_cols)]
            
            # Definir Cores (Convertendo para o padrão que o Excel entende)
            color_readonly = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
            color_editable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
            color_id = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon)
            
            # 2. Escrever Cabeçalhos e Pintar Colunas
            ws.Cells[1, 1] = "UniqueId"
            ws.Cells[1, 1].Interior.Color = color_id
            
            for c in range(n_cols):
                col_idx = c + 2
                field = fields[c]
                ws.Cells[1, col_idx] = field.ColumnHeading
                
                # Pinta o cabeçalho de acordo com a editabilidade (Read-Only)
                if field.IsReadOnly:
                    ws.Cells[1, col_idx].Interior.Color = color_readonly
                else:
                    ws.Cells[1, col_idx].Interior.Color = color_editable
            
            # 3. Extrair os Elementos na mesma ordem da Tabela
            elements = list(DB.FilteredElementCollector(doc, schedule.Id).ToElements())
            
            # 4. Extrair os Textos das Células
            table_data = schedule.GetTableData()
            section_data = table_data.GetSectionData(DB.SectionType.Body)
            n_rows = section_data.NumberOfRows
            
            for r in range(n_rows):
                row_idx = r + 2
                
                # Inserir UniqueId na primeira coluna
                if r < len(elements):
                    ws.Cells[row_idx, 1] = elements[r].UniqueId
                else:
                    ws.Cells[row_idx, 1] = "N/A"
                
                # Inserir Valores dos Parâmetros
                for c in range(n_cols):
                    val = schedule.GetCellText(DB.SectionType.Body, r, c)
                    ws.Cells[row_idx, c + 2] = str(val) if val else ""
            
            # Formatando o Excel
            ws.Range[ws.Cells[1, 1], ws.Cells[1, n_cols + 1]].Font.Bold = True
            ws.Columns.AutoFit()
            
        except Exception as ex:
            forms.alert("Erro ao exportar Excel: {}".format(ex))

if __name__ == '__main__':
    xaml_file = script.get_bundle_file("ExcelLink.xaml")
    if os.path.exists(xaml_file):
        window = ExcelLinkWindow(xaml_file)
        window.ShowDialog()
    else:
        forms.alert("Arquivo 'ExcelLink.xaml' não encontrado na pasta do script.", title="Erro de UI")