# -*- coding: utf-8 -*-
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import os, clr

#>>>>>>>>>> pyRevit
from pyrevit import forms  # Needed for wpf import to work.

# Custom Imports
from GUI.forms import my_WPF

#>>>>>>>>>> .NET IMPORTS
clr.AddReference("System")
from System.Collections.Generic import List
import wpf

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
PATH_SCRIPT = os.path.dirname(__file__)


class ListItem:
    """Item exibido na lista de revisao de Ambientes.
    :param Name:        Identificacao do Ambiente (Numero - Nome).
    :param Detalhe:     Texto auxiliar com a origem e os valores Atual/Novo.
    :param element:     Elemento Room associado a este item.
    :param novo_offset: Novo valor (em pes) a ser aplicado em ROOM_UPPER_OFFSET.
    :param checked:     Estado inicial do checkbox."""
    def __init__(self, Name='Unnamed', Detalhe='', element=None, novo_offset=0.0, checked=True):
        self.Name = Name
        self.Detalhe = Detalhe
        self.IsChecked = checked
        self.element = element
        self.novo_offset = novo_offset


# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================
class RoomCeilingSync(my_WPF):
    def __init__(self, candidatos,
                 title       = '__title__',
                 label       = 'Ajustes encontrados:',
                 button_name = 'Aplicar',
                 version     = 'Version: 1.0'):
        """Janela de revisao em lote dos ajustes de Limite Superior dos Ambientes.
        :param candidatos: lista de dicts {'room', 'nome', 'detalhe', 'novo_offset'}"""
        self.candidatos     = candidatos
        self.items          = self.generate_list_items()
        self.selected_items = []

        #>>>>>>>>>> SET RESOURCES AND LOAD WPF
        self.add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'RoomCeilingSync.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # UPDATE GUI ELEMENTS
        self.main_title.Text     = title
        self.text_label.Content  = label
        self.button_main.Content = button_name
        self.footer_version.Text = version

        self.main_ListBox.ItemsSource = self.items
        self.ShowDialog()

    def __iter__(self):
        """Retorna os itens selecionados pelo usuario."""
        return iter(self.selected_items)

    def generate_list_items(self):
        """Converte a lista de candidatos em ListItems para o ListBox."""
        list_of_items = List[type(ListItem())]()
        for c in self.candidatos:
            list_of_items.Add(ListItem(c['nome'], c['detalhe'], c['room'], c['novo_offset'], True))
        return list_of_items

    #>>>>>>>>>> INHERIT WPF RESOURCES
    def add_wpf_resource(self):
        super(RoomCeilingSync, self).add_wpf_resource()

    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def text_filter_updated(self, sender, e):
        """Filtra os itens do ListBox pelo nome do Ambiente."""
        filtered_list_of_items = List[type(ListItem())]()
        filter_keyword = self.textbox_filter.Text

        if not filter_keyword:
            self.main_ListBox.ItemsSource = self.items
            return

        for item in self.items:
            if filter_keyword.lower() in item.Name.lower():
                filtered_list_of_items.Add(item)

        self.main_ListBox.ItemsSource = filtered_list_of_items

    # ╔╗ ╦ ╦╔╦╗╔╦╗╔═╗╔╗╔╔═╗
    # ╠╩╗║ ║ ║  ║ ║ ║║║║╚═╗
    # ╚═╝╚═╝ ╩  ╩ ╚═╝╝╚╝╚═╝ BUTTONS
    #==================================================
    def select_mode(self, mode):
        """Marca/Desmarca todos os itens visiveis no ListBox."""
        list_of_items = List[type(ListItem())]()
        checked = True if mode == 'all' else False
        for item in self.main_ListBox.ItemsSource:
            item.IsChecked = checked
            list_of_items.Add(item)
        self.main_ListBox.ItemsSource = list_of_items

    def button_select_all(self, sender, e):
        self.select_mode(mode='all')

    def button_select_none(self, sender, e):
        self.select_mode(mode='none')

    def button_apply(self, sender, e):
        """Fecha a janela e retorna apenas os itens marcados."""
        self.textbox_filter.Text = ''
        self.Close()

        self.selected_items = [item for item in self.items if item.IsChecked]


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝MAIN
#====================================================================================================
def review_room_adjustments(candidatos,
                             title       = 'NnBim',
                             label       = 'Ajustes encontrados:',
                             button_name = 'Aplicar',
                             version     = 'Version: 1.0'):
    #type:(list, str, str, str, str) -> list
    """Exibe a janela de revisao em lote dos Ambientes a serem ajustados.
    :param candidatos: lista de dicts {'room', 'nome', 'detalhe', 'novo_offset'}
    :return:           lista de ListItem marcados pelo usuario (ListItem.element = Room)."""

    gui = RoomCeilingSync(candidatos,
                          title       = title,
                          label       = label,
                          button_name = button_name,
                          version     = version)
    return list(gui)
