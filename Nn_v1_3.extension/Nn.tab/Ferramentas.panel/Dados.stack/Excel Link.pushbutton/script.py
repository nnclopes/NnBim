# -*- coding: utf-8 -*-
"""
NnBim: EXCEL LINK PRO

DESCRICAO:
Substitui o SheetLink (DiRoots). Exporta qualquer tabela (schedule) do Revit
para uma planilha Excel - uma aba por tabela, SEMPRE com 1 linha por elemento
(independente de agrupamento/ordenacao/totais configurados na tabela).

Cores das colunas:
- SALMAO : "UniqueId" (identificador interno, nao apague nem edite).
- VERDE  : parametro editavel - pode ser alterado no Excel.
- CINZA  : parametro somente leitura (informativo).

Campos calculados/combinados do schedule (sem parametro real, ex: "Contagem")
nao sao exportados, pois nao podem ser escritos de volta no Revit.

COMO USAR:
1. Na aba "Tabelas (Schedules)", selecione uma ou mais tabelas e clique em
   EXPORTAR DADOS. Todas as tabelas selecionadas vao para UMA UNICA planilha
   (uma aba por tabela).
2. Edite somente as celulas com fundo VERDE.
3. NAO exclua a coluna "UniqueId" e NAO desoculte/edite a linha 2 (uso
   interno do NnBim - guarda o nome real dos parametros para a importacao).
4. Salve o arquivo e use IMPORTAR EXCEL para aplicar as alteracoes no modelo.
"""
__title__ = 'Excel\nLink'
__author__ = 'NnBim Dev'

from pyrevit import revit, DB, forms, script
import os
import sys
import gc
import clr

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

clr.AddReference("System.Drawing")
import System.Drawing

clr.AddReference("System")
from System.Runtime.InteropServices import Marshal

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

# Marcador gravado na linha 2 (oculta) de cada aba exportada, usado para
# reconhecer um arquivo gerado pelo NnBim e localizar o mapeamento de
# parametros na importacao.
META_MARKER = "#NnBim_META#"


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

    # ------------------------------------------------------------------
    # EXPORTACAO
    # ------------------------------------------------------------------
    def on_export_click(self, sender, e):
        selected_tab = self.tab_main.SelectedIndex

        if selected_tab == 0:
            selected_names = list(self.list_schedules.SelectedItems)
            if not selected_names:
                forms.alert("Selecione pelo menos uma tabela.")
                return

            schedules_para_exportar = [
                self._schedule_dict[name] for name in selected_names
                if name in self._schedule_dict
            ]

            self.Close()

            try:
                self._exportar_schedules(schedules_para_exportar)
            except Exception as ex:
                forms.alert("Erro ao exportar:\n{}".format(ex), title="NnBim Excel Link")
        else:
            forms.alert("Exportacao por categoria na proxima atualizacao!", title="NnBim Dev")
            self.Close()

    def _exportar_schedules(self, schedules):
        """Abre UMA instancia do Excel e cria uma aba por tabela selecionada."""
        app = Excel.ApplicationClass()
        app.Visible = True
        wb = app.Workbooks.Add()

        # Remove as abas extras criadas por padrao, mantendo apenas a primeira
        while wb.Worksheets.Count > 1:
            wb.Worksheets[wb.Worksheets.Count].Delete()

        nomes_usados = set()
        for idx, sched in enumerate(schedules):
            if idx == 0:
                ws = wb.Worksheets[1]
            else:
                ws = wb.Worksheets.Add(After=wb.Worksheets[wb.Worksheets.Count])

            ws.Name = self._nome_aba_valido(sched.Name, nomes_usados)
            nomes_usados.add(ws.Name)
            self._preencher_aba(ws, sched)

        wb.Worksheets[1].Activate()

    def _preencher_aba(self, ws, schedule):
        """Preenche uma aba: 1 coluna por parametro real do schedule, 1 linha por elemento."""
        definition = schedule.Definition
        fields = [definition.GetField(i) for i in range(definition.GetFieldCount())]
        elements = list(DB.FilteredElementCollector(doc, schedule.Id).ToElements())

        color_readonly = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)
        color_editable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
        color_id       = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon)

        # Analisa cada campo do schedule e descarta os que nao correspondem
        # a um parametro real (campos calculados, combinados, contagem, etc.)
        campos_incluidos = []  # [(field, nome_real_parametro, editavel)]
        for f in fields:
            if f.IsCalculatedField or getattr(f, 'IsCombinedParameterField', False):
                continue

            param_id = f.ParameterId
            editavel = None
            for el in elements:
                el_type = doc.GetElement(el.GetTypeId())
                param = self._find_param(el, param_id) or self._find_param(el_type, param_id)
                if param is not None:
                    editavel = not param.IsReadOnly
                    break

            if editavel is None:
                continue  # nenhum elemento possui esse parametro -> ignora coluna

            campos_incluidos.append((f, f.GetName(), editavel))

        # Linha 1: cabecalhos visiveis
        ws.Cells[1, 1] = "UniqueId"
        ws.Cells[1, 1].Interior.Color = color_id

        # Linha 2 (oculta): marcador + nome real do parametro de cada coluna
        ws.Cells[2, 1] = META_MARKER

        for c, (field, nome_real, editavel) in enumerate(campos_incluidos):
            col = c + 2
            ws.Cells[1, col] = field.ColumnHeading
            ws.Cells[1, col].Interior.Color = color_editable if editavel else color_readonly
            ws.Cells[2, col] = nome_real

        ws.Rows[2].Hidden = True

        # Linhas de dados: 1 elemento = 1 linha
        for r, el in enumerate(elements):
            row = r + 3
            ws.Cells[row, 1] = el.UniqueId
            el_type = doc.GetElement(el.GetTypeId())

            for c, (field, nome_real, editavel) in enumerate(campos_incluidos):
                col = c + 2
                param = self._find_param(el, field.ParameterId) or self._find_param(el_type, field.ParameterId)
                valor = ""
                if param is not None:
                    try:
                        valor = param.AsValueString() or ""
                    except Exception:
                        valor = ""
                ws.Cells[row, col] = valor

        n_cols_total = len(campos_incluidos) + 1
        ws.Range[ws.Cells[1, 1], ws.Cells[1, n_cols_total]].Font.Bold = True
        ws.Columns.AutoFit()

    # ------------------------------------------------------------------
    # IMPORTACAO
    # ------------------------------------------------------------------
    def on_import_click(self, sender, e):
        file_path = forms.pick_file(file_ext='xlsx')
        if not file_path:
            return

        self.Close()

        app_excel = None
        wb = None
        dados_abas = []

        try:
            app_excel = Excel.ApplicationClass()
            app_excel.Visible = False
            wb = app_excel.Workbooks.Open(file_path)

            for ws in wb.Worksheets:
                if ws.Cells[2, 1].Value2 == META_MARKER:
                    dados_abas.append(self._ler_aba(ws))

            if not dados_abas:
                forms.alert(
                    "Nenhuma aba reconhecida como exportacao do NnBim Excel Link.\n"
                    "Use um arquivo gerado pela funcao EXPORTAR.",
                    title="NnBim"
                )
                return

        except Exception as ex:
            forms.alert("Erro ao ler o Excel:\n{}".format(ex), title="NnBim")
            return
        finally:
            if wb is not None:
                wb.Close(False)
                self._liberar_com(wb)
            if app_excel is not None:
                app_excel.Quit()
                self._liberar_com(app_excel)
            gc.collect()
            gc.WaitForPendingFinalizers()

        total_sucessos        = 0
        total_nao_encontrados = 0
        total_erros           = 0

        with revit.Transaction("NnBim: Importar Excel Link"):
            for nome_aba, linhas in dados_abas:
                if not linhas:
                    continue
                with forms.ProgressBar(title=nome_aba + " ({value} de {max_value})", total=len(linhas)) as pb:
                    for idx, (uid, valores) in enumerate(linhas):
                        pb.update_progress(idx + 1, len(linhas))

                        el = doc.GetElement(uid)
                        if el is None:
                            total_nao_encontrados += 1
                            continue

                        el_type = doc.GetElement(el.GetTypeId())
                        atualizado = False

                        for nome_param, valor in valores.items():
                            param = el.LookupParameter(nome_param)
                            if param is None and el_type:
                                param = el_type.LookupParameter(nome_param)
                            if param is None or param.IsReadOnly:
                                continue

                            try:
                                if param.SetValueString(valor) or self._set_valor_bruto(param, valor):
                                    atualizado = True
                            except Exception:
                                total_erros += 1

                        if atualizado:
                            total_sucessos += 1

        msg = "Importacao concluida!\n\n{} elementos atualizados.".format(total_sucessos)
        if total_nao_encontrados:
            msg += "\n{} UniqueIds nao encontrados.".format(total_nao_encontrados)
        if total_erros:
            msg += "\n{} valores nao puderam ser aplicados.".format(total_erros)
        forms.alert(msg, title="NnBim Excel Link")

    def _ler_aba(self, ws):
        """Le uma aba exportada pelo NnBim.

        Retorna (nome_aba, [(uid, {nome_real_param: valor_texto})]).
        """
        n_cols = ws.UsedRange.Columns.Count
        n_rows = ws.UsedRange.Rows.Count

        # Linha 2 (oculta) guarda o nome real do parametro de cada coluna
        parametros = {}
        for c in range(2, n_cols + 1):
            nome_real = ws.Cells[2, c].Value2
            if nome_real:
                parametros[c] = str(nome_real)

        linhas = []
        for r in range(3, n_rows + 1):
            uid = ws.Cells[r, 1].Value2
            if not uid or str(uid).strip() in ("", "N/A"):
                continue

            valores = {}
            for col, nome_param in parametros.items():
                val = ws.Cells[r, col].Value2
                valores[nome_param] = str(val) if val is not None else ""

            linhas.append((str(uid).strip(), valores))

        return (ws.Name, linhas)

    # ------------------------------------------------------------------
    # AUXILIARES
    # ------------------------------------------------------------------
    @staticmethod
    def _find_param(elemento, parameter_id):
        """Procura, nos parametros de um elemento, aquele cuja Definition.Id
        corresponde ao ParameterId de um campo do schedule."""
        if elemento is None:
            return None
        for p in elemento.Parameters:
            if p.Id == parameter_id:
                return p
        return None

    @staticmethod
    def _set_valor_bruto(param, valor):
        """Fallback quando SetValueString nao aceita o texto formatado."""
        if param.StorageType == DB.StorageType.String:
            param.Set(valor)
            return True
        elif param.StorageType == DB.StorageType.Double:
            if valor:
                param.Set(float(valor))
                return True
        elif param.StorageType == DB.StorageType.Integer:
            if valor:
                param.Set(int(float(valor)))
                return True
        return False

    @staticmethod
    def _nome_aba_valido(nome, nomes_usados):
        """Garante nome de aba valido no Excel (max 31 caracteres, sem
        caracteres invalidos e sem repeticao)."""
        invalidos = ['\\', '/', '?', '*', '[', ']', ':']
        limpo = nome
        for ch in invalidos:
            limpo = limpo.replace(ch, '-')

        base = limpo if len(limpo) <= 31 else limpo[:28] + "..."

        final = base
        contador = 1
        while final in nomes_usados:
            contador += 1
            sufixo = " ({})".format(contador)
            final = base[:31 - len(sufixo)] + sufixo

        return final

    @staticmethod
    def _liberar_com(com_obj):
        try:
            Marshal.ReleaseComObject(com_obj)
        except Exception:
            pass


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
