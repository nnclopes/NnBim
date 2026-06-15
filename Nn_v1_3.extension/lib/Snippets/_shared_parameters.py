# -*- coding: utf-8 -*-
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
from Autodesk.Revit.DB import BuiltInParameterGroup, ExternalDefinitionCreationOptions
from pyrevit import forms

from Snippets._context_manager import ef_Transaction

# Em versoes mais novas (Revit 2024+) o enum ParameterType foi removido em
# favor do ForgeTypeId (SpecTypeId). Garimpo: detecta qual esta disponivel.
try:
    from Autodesk.Revit.DB import ParameterType
    _PARAMETER_TYPE_DISPONIVEL = True
except ImportError:
    from Autodesk.Revit.DB import SpecTypeId
    _PARAMETER_TYPE_DISPONIVEL = False

app = __revit__.Application


# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝ FUNCTIONS
#====================================================================================================
def parametros_ja_vinculados(doc, nomes_parametros):
    """Retorna a lista de nomes (dentre 'nomes_parametros') que AINDA NAO
    estao vinculados (Bound) ao projeto atual."""
    vinculados = set()
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        vinculados.add(it.Key.Name)

    return [nome for nome in nomes_parametros if nome not in vinculados]


def _obter_ou_criar_definicao(spf, grupo_nome, nome_parametro, tipo='Length'):
    """Procura a Definition de 'nome_parametro' dentro do grupo 'grupo_nome'
    do arquivo de Parametros Compartilhados informado. Cria o grupo e/ou a
    Definition caso nao existam - a criacao e gravada direto no .txt
    selecionado pelo usuario."""
    grupo = None
    for g in spf.Groups:
        if g.Name == grupo_nome:
            grupo = g
            break
    if grupo is None:
        grupo = spf.Groups.Create(grupo_nome)

    for definicao in grupo.Definitions:
        if definicao.Name == nome_parametro:
            return definicao

    if _PARAMETER_TYPE_DISPONIVEL:
        tipo_param = ParameterType.Length if tipo == 'Length' else ParameterType.Text
    else:
        tipo_param = SpecTypeId.Length if tipo == 'Length' else SpecTypeId.String.Text

    opcoes = ExternalDefinitionCreationOptions(nome_parametro, tipo_param)
    opcoes.Visible = True
    return grupo.Definitions.Create(opcoes)


def _vincular_parametro(doc, definicao, categorias, grupo_param, instancia=True):
    """Vincula a Definition informada as categorias do projeto, caso ainda
    nao esteja vinculada. Retorna True se vinculou agora, False se ja existia."""
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        if it.Key.Name == definicao.Name:
            return False

    cat_set = app.Create.NewCategorySet()
    for categoria in categorias:
        cat_set.Insert(doc.Settings.Categories.get_Item(categoria))

    if instancia:
        binding = app.Create.NewInstanceBinding(cat_set)
    else:
        binding = app.Create.NewTypeBinding(cat_set)

    doc.ParameterBindings.Insert(definicao, binding, grupo_param)
    return True


def garantir_parametros_compartilhados(doc, nomes_parametros, categorias,
                                        grupo_definicao='NnBim',
                                        grupo_param=BuiltInParameterGroup.PG_GEOMETRY,
                                        tipo='Length', instancia=True):
    """Garimpo: garante que os parametros compartilhados 'nomes_parametros'
    estejam criados e vinculados as 'categorias' do projeto atual.

    Caso algum ainda nao esteja vinculado, solicita ao usuario o arquivo de
    Parametros Compartilhados (.txt) e cria a Definition (se necessario) e o
    Binding com as categorias informadas - direto nesse .txt. Isso permite
    levar o botao para outros projetos/computadores: na primeira execucao em
    cada projeto, basta apontar para o .txt de Parametros Compartilhados da
    empresa (ou criar um novo .txt vazio para esse fim).

    :param nomes_parametros: lista de nomes dos parametros (ex: ['Nn_Largura'])
    :param categorias:       lista de BuiltInCategory para o Binding
    :param grupo_definicao:  nome do grupo dentro do arquivo .txt
    :param grupo_param:      BuiltInParameterGroup exibido no Revit (ex: PG_GEOMETRY)
    :param tipo:             'Length' ou 'Text'
    :param instancia:        True para Instance Binding, False para Type Binding
    :return: lista de nomes que NAO foram vinculados (vazia = sucesso total)"""

    faltantes = parametros_ja_vinculados(doc, nomes_parametros)
    if not faltantes:
        return []

    txt_path = forms.pick_file(
        file_ext='txt',
        title='Selecione o arquivo de Parametros Compartilhados (.txt) para criar: {}'.format(", ".join(faltantes))
    )
    if not txt_path:
        return faltantes

    nao_vinculados = []
    original_spf = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = txt_path
        spf = app.OpenSharedParameterFile()
        if not spf:
            return faltantes

        with ef_Transaction(doc, 'NnBim - Criar Parametros Compartilhados', debug=True):
            for nome in faltantes:
                try:
                    definicao = _obter_ou_criar_definicao(spf, grupo_definicao, nome, tipo=tipo)
                    _vincular_parametro(doc, definicao, categorias, grupo_param, instancia=instancia)
                except Exception:
                    nao_vinculados.append(nome)
    finally:
        app.SharedParametersFilename = original_spf

    return nao_vinculados
