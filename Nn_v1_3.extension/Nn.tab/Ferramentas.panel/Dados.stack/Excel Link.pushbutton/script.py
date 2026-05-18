# -*- coding: utf-8 -*-
"""
NnBim: EXCEL LINK PRO
DESCRICAO:
Sincronizacao bidirecional entre Revit e Excel padrao DiRoots.
Permite exportar tabelas com UniqueId e mapeamento de cores.
Na importacao, atualiza apenas os parametros editaveis (verdes).

COMO USAR:
1. Escolha a aba desejada (Tabelas ou Categorias).
2. Exporte os dados. Parametros editaveis ficarao em VERDE.
3. Altere o Excel e use a funcao IMPORTAR para atualizar o Revit.
"""
__title__ = 'Excel\nLink'
__author__ = 'Nn_Dev'

from pyrevit import revit, DB, forms, script
import os
import sys
import clr

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

clr.AddReference("System.Drawing")
import System.Drawing

cur_dir   = os.path.dirname(__file__)
panel_dir = os.path.dirname(cur_dir)
tab_dir   = os.path.dirname(panel_dir)
ext_dir   = os.path.dirname(tab_dir)
root_dir  = os.path.dirname(ext_dir)
lib_path  = os.path.join(root_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    from GUI.WPF_Base import my_WPF as WPFWindow
except ImportError as e:
    forms.alert(
        "Erro ao carregar a biblioteca:\n{}\n\nLib esperada em:\n{}".format(e, lib_path),
        title="Erro NnBim"
    )
    sys.exit()

doc   = revit.doc
uidoc = revit.uidoc


class ExcelLinkWindow(WPFWindow):

    def __init__(self, xaml_file_name):
        import wpf
        wpf.LoadComponent(self, xaml_file_name)
        self._setup_initial_data()

    def _setup_initial_data(self):
        all_schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
        self._schedule_dict = {
            s.Name: s
            for s in all_schedules
            if not s.IsTitleblockRevisionSchedule
        }
        self.list_schedules.ItemsSource = sorted(self._schedule_dict.keys())

        categories = doc.Settings.Categories
        valid_cats = sorted([c.Name for c in categories if c.AllowsBoundParameters])
        self.cb_categories.ItemsSource = valid_cats

    def on_category_changed(self, sender, e):
        self.list_available_params.Items.Clear()
        self.list_selected_params.Items.Clear()

        cat_name = self.cb_categories.SelectedItem
        if not cat_name:
            return

        target_cat = next((c for c in doc.Settings.Categories if c.Name == cat_name), None)
        if not target_cat:
            return

        collector = DB.FilteredElementCollector(doc).OfCategoryId(target_cat.Id).WhereElementIsNotElementType()
        el = collector.FirstElement()
        param_names = set()

        if el:
            for p in el.Parameters:
                param_names.add(p.Definition.Name)
            el_type = doc.GetElement(el.GetTypeId())
            if el_type:
                for p in el_type.Parameters:
                    param_names.add(p.Definition.Name)
        else:
            iterator = doc.ParameterBindings.ForwardIterator()
            while iterator.MoveNext():
                if target_cat in iterator.Current.Categories:
                    param_names.add(iterator.Key.Name)

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

    def on_export_click(self, sender, e):
        selected_tab = self.tab_main.SelectedIndex

        if selected_tab == 0:
            selected_names = list(self.list_schedules.SelectedItems)
            if not selected_names:
                forms.alert("Selecione pelo menos uma tabela.")
                return

            with DB.Transaction(doc, "NnBim: Prep Export") as t:
                t.Start()
                schedules_para_exportar = []
                for name in selected_names:
                    sched = self._schedule_dict.get(name)
                    if sched is None:
                        continue
                    sched.Definition.IsItemized = True
                    schedules_para_exportar.append(sched)

                doc.Regenerate()

                for sched in schedules_para_exportar:
                    self._export_schedule_to_excel(sched)

                t.RollBack()
        else:
            forms.alert("Exportacao por categoria na proxima atualizacao!", title="NnBim Dev")

        self.Close()

    def on_import_click(self, sender, e):
        file_path = forms.pick_file(file_ext='xlsx')
        if not file_path:
            return

        self.Close()

        try:
            app_excel = Excel.ApplicationClass()
            app_excel.Visible = False
            wb = app_excel.Workbooks.Open(file_path)
            ws = wb.ActiveSheet

            n_cols = ws.UsedRange.Columns.Count
            headers = []
            for c in range(2, n_cols + 1):
                cell_val = ws.Cells[1, c].Value2
                headers.append(str(cell_val) if cell_val else "")

            color_editable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
            editaveis = []
            for i, col_excel in enumerate(range(2, n_cols + 1)):
                bg = ws.Cells[1, col_excel].Interior.Color
                if int(bg) == int(color_editable):
                    editaveis.append((i, col_excel))

            if not editaveis:
                forms.alert(
                    "Nenhuma coluna editavel (verde) encontrada.\n"
                    "Use um arquivo exportado pelo NnBim Excel Link.",
                    title="NnBim"
                )
                wb.Close(False)
                app_excel.Quit()
                return

            n_rows = ws.UsedRange.Rows.Count
            linhas = []
            for r in range(2, n_rows + 1):
                uid = ws.Cells[r, 1].Value2
                if not uid or str(uid).strip() in ("", "N/A"):
                    continue
                valores = {}
                for i, col_excel in editaveis:
                    param_name = headers[i]
                    val = ws.Cells[r, col_excel].Value2
                    valores[param_name] = str(val) if val is not None else ""
                linhas.append((str(uid).strip(), valores))

            wb.Close(False)
            app_excel.Quit()

            if not linhas:
                forms.alert("Nenhuma linha de dados encontrada.", title="NnBim")
                return

        except Exception as ex:
            forms.alert("Erro ao ler o Excel:\n{}".format(ex), title="NnBim")
            return

        sucessos        = 0
        nao_encontrados = 0
        erros_param     = 0

        with revit.Transaction("NnBim: Importar Excel Link"):
            with forms.ProgressBar(title="Importando... ({value} de {max_value})", total=len(linhas)) as pb:
                for idx, (uid, valores) in enumerate(linhas):
                    pb.update_progress(idx + 1, len(linhas))
                    el = doc.GetElement(uid)
                    if el is None:
                        nao_encontrados += 1
                        continue
                    for param_name, valor in valores.items():
                        param = el.LookupParameter(param_name)
                        if param is None or param.IsReadOnly:
                            continue
                        try:
                            if param.StorageType == DB.StorageType.String:
                                param.Set(valor)
                            elif param.StorageType == DB.StorageType.Double:
                                if valor: param.Set(float(valor))
                            elif param.StorageType == DB.StorageType.Integer:
                                if valor: param.Set(int(float(valor)))
                        except Exception:
                            erros_param += 1
                    sucessos += 1

        msg = "Importacao concluida!\n\n{} elementos atualizados.".format(sucessos)
        if nao_encontrados:
            msg += "\n{} UniqueIds nao encontrados.".format(nao_encontrados)
        if erros_param:
            msg += "\n{} valores com tipo incompativel.".format(erros_param)
        forms.alert(msg, title="NnBim Excel Link")

    def _export_schedule_to_excel(self, schedule):
        try:
            app = Excel.ApplicationClass()
            app.Visible = True
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Name = (schedule.Name[:30] + '..') if len(schedule.Name) > 31 else schedule.Name

            definition = schedule.Definition
            n_cols     = definition.GetFieldCount()
            fields     = [definition.GetField(i) for i in range(n_cols)]

            color_readonly = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
            color_editable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
            color_id       = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon)

            ws.Cells[1, 1] = "UniqueId"
            ws.Cells[1, 1].Interior.Color = color_id

            for c in range(n_cols):
                col_idx = c + 2
                field   = fields[c]
                ws.Cells[1, col_idx] = field.ColumnHeading
                ws.Cells[1, col_idx].Interior.Color = color_readonly if field.IsReadOnly else color_editable

            elements     = list(DB.FilteredElementCollector(doc, schedule.Id).ToElements())
            table_data   = schedule.GetTableData()
            section_data = table_data.GetSectionData(DB.SectionType.Body)
            n_rows       = section_data.NumberOfRows

            for r in range(n_rows):
                row_idx = r + 2
                ws.Cells[row_idx, 1] = elements[r].UniqueId if r < len(elements) else "N/A"
                for c in range(n_cols):
                    val = schedule.GetCellText(DB.SectionType.Body, r, c)
                    ws.Cells[row_idx, c + 2] = str(val) if val else ""

            ws.Range[ws.Cells[1, 1], ws.Cells[1, n_cols + 1]].Font.Bold = True
            ws.Columns.AutoFit()

        except Exception as ex:
            forms.alert("Erro ao exportar '{}': {}".format(schedule.Name, ex))


if __name__ == '__main__':
    xaml_file = script.get_bundle_file("ExcelLink.xaml")
    if os.path.exists(xaml_file):
        window = ExcelLinkWindow(xaml_file)
        window.ShowDialog()
    else:
        forms.alert(
            "Arquivo 'ExcelLink.xaml' nao encontrado na pasta do script.",
            title="Erro de UI"
        )