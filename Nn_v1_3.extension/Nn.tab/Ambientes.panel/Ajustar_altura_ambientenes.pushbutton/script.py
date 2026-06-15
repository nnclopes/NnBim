# -*- coding: utf-8 -*-
"""
Titulo: Sincronizar Altura de Ambientes com Forro/Piso/Telhado
Descricao: Calcula largura/comprimento do Ambiente e localiza o elemento que
realmente fecha o seu teto - primeiro um Forro (Ceiling), e na ausencia deste
um Piso (Floor) ou um Telhado (Roof). Os parametros 'Nn_Largura',
'Nn_Comprimento' e 'Nn_Altura do Forro' sao atualizados com os valores reais
encontrados no modelo. Em seguida, uma janela de revisao em lote lista todos
os Ambientes cujo Limite Superior (Upper Offset) diverge do elemento
encontrado, permitindo escolher quais ajustes aplicar de uma so vez.
Instrucoes: Nao e necessario selecionar nada - o script percorre todos os
Ambientes do projeto. Na primeira execucao em um projeto novo, caso os
parametros 'Nn_Largura', 'Nn_Comprimento' e 'Nn_Altura do Forro' (Tipo
Comprimento) ainda nao existam, sera solicitado o arquivo de Parametros
Compartilhados (.txt) para cria-los e vincula-los automaticamente as
Ambientes. Ao final, marque os Ambientes desejados na janela e clique em
'Aplicar'.
"""
__title__ = 'Ajustar altura\ndos ambientes'
__author__ = 'NnBim Dev'

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, XYZ
from pyrevit import revit, forms

from Snippets._context_manager import ef_Transaction
from Snippets._boundingbox import is_point_in_BB_2D
from Snippets._convert import convert_internal_to_m, convert_m_to_feet
from Snippets._shared_parameters import garantir_parametros_compartilhados
from GUI.Tools.RoomCeilingSync import review_room_adjustments

doc = revit.doc

# Parametros compartilhados (tipo Comprimento) atualizados pelo script.
# Caso ainda nao existam no projeto, o usuario sera solicitado a indicar o
# arquivo de Parametros Compartilhados (.txt) para cria-los/vincula-los.
NOMES_PARAMETROS_NN = ['Nn_Largura', 'Nn_Comprimento', 'Nn_Altura do Forro']

# Tolerancia para comparacoes de offset (~1cm em pes)
TOL_AJUSTE = 0.0328

# Pe-direito minimo (em pes) para um elemento ser aceito como "teto" do Ambiente.
# Garimpo: evita falsos positivos com Pisos/Telhados de OUTROS ambientes que se
# sobrepoem em planta mas estao praticamente no mesmo nivel do piso do Ambiente
# analisado (ex: piso de acabamento do pavimento superior apoiado sobre a laje).
ALTURA_MINIMA = convert_m_to_feet(1.0)

# Categorias pesquisadas, em ordem de prioridade, como possivel "teto" do Ambiente
ORIGENS = (
    ('Forro',   BuiltInCategory.OST_Ceilings),
    ('Piso',    BuiltInCategory.OST_Floors),
    ('Telhado', BuiltInCategory.OST_Roofs),
)


def get_dados_elementos(categoria):
    """Garimpo: pre-calcula a Bounding Box e a elevacao da face inferior de cada
    elemento da categoria informada (Forro, Piso ou Telhado)."""
    elementos = FilteredElementCollector(doc).OfCategory(categoria) \
                                              .WhereElementIsNotElementType().ToElements()

    dados = []
    for el in elementos:
        bbox = el.get_BoundingBox(None)
        if not bbox:
            continue

        if categoria == BuiltInCategory.OST_Ceilings:
            # Forro: usa o Nivel + 'Height Offset From Level' (face inferior real)
            nivel = doc.GetElement(el.LevelId)
            if not nivel:
                continue
            p_offset = el.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
            offset = p_offset.AsDouble() if p_offset else 0.0
            elevacao = nivel.Elevation + offset
        else:
            # Piso/Telhado: usa a face inferior da Bounding Box como referencia
            elevacao = bbox.Min.Z

        dados.append({'bbox': bbox, 'elevacao': elevacao})

    return dados


def encontrar_elemento_acima(bbox_room, base_elevacao, elementos):
    """Procura, entre os elementos informados, o que esta diretamente acima do
    Ambiente. Retorna o elemento mais baixo cuja projecao em planta contenha o
    centro do Ambiente e que esteja no nivel do piso ou acima dele."""

    centro = XYZ(
        (bbox_room.Max.X + bbox_room.Min.X) / 2.0,
        (bbox_room.Max.Y + bbox_room.Min.Y) / 2.0,
        bbox_room.Min.Z
    )

    melhor = None
    for el in elementos:
        # Descarta elementos cuja face inferior nao garanta o pe-direito minimo
        # (evita "tetos" falsos de pisos/forros muito proximos do piso do Ambiente)
        if el['elevacao'] - base_elevacao < ALTURA_MINIMA - TOL_AJUSTE:
            continue

        if not is_point_in_BB_2D(el['bbox'], centro):
            continue

        if melhor is None or el['elevacao'] < melhor['elevacao']:
            melhor = el

    return melhor


def main():
    # Garante que os parametros 'Nn_Largura', 'Nn_Comprimento' e
    # 'Nn_Altura do Forro' existam no projeto (cria/vincula se necessario)
    parametros_pendentes = garantir_parametros_compartilhados(
        doc, NOMES_PARAMETROS_NN, [BuiltInCategory.OST_Rooms]
    )

    aviso_parametros = ""
    if parametros_pendentes:
        aviso_parametros = "\n\nParametro(s) nao vinculado(s) ao projeto - os campos correspondentes nao serao preenchidos:\n- {}".format(
            "\n- ".join(parametros_pendentes)
        )

    rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms) \
                                          .WhereElementIsNotElementType().ToElements()

    # Pre-calcula os dados de Forros, Pisos e Telhados (sob demanda)
    dados_por_origem = {}
    for nome_origem, categoria in ORIGENS:
        dados_por_origem[nome_origem] = get_dados_elementos(categoria)

    sem_elemento = []
    candidatos = []

    # ETAPA 1: Atualiza dimensoes e 'Nn_Altura do Forro' (silencioso)
    with ef_Transaction(doc, 'NnBim - Sincronizar Dimensoes', debug=True):
        for room in rooms:
            try:
                bbox = room.get_BoundingBox(None)
                if not bbox:
                    continue

                nome_ambiente = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
                numero_ambiente = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()

                # Dimensoes do Ambiente (Bounding Box)
                largura = bbox.Max.X - bbox.Min.X
                comprimento = bbox.Max.Y - bbox.Min.Y

                p_largura = room.LookupParameter("Nn_Largura")
                p_comprimento = room.LookupParameter("Nn_Comprimento")
                p_altura = room.LookupParameter("Nn_Altura do Forro")

                if p_largura: p_largura.Set(largura)
                if p_comprimento: p_comprimento.Set(comprimento)

                # Elevacao do piso do Ambiente (Nivel + Offset Inferior)
                base_elevacao = room.Level.Elevation + room.get_Parameter(BuiltInParameter.ROOM_LOWER_OFFSET).AsDouble()

                # Garimpo: localiza o Forro, Piso ou Telhado diretamente acima do Ambiente
                elemento = None
                origem = None
                for nome_origem, _ in ORIGENS:
                    elemento = encontrar_elemento_acima(bbox, base_elevacao, dados_por_origem[nome_origem])
                    if elemento:
                        origem = nome_origem
                        break

                if not elemento:
                    sem_elemento.append("{} ({})".format(nome_ambiente, numero_ambiente))
                    continue

                # Altura real do Ambiente (piso -> face inferior do elemento encontrado)
                altura_real = elemento['elevacao'] - base_elevacao
                if p_altura: p_altura.Set(altura_real)

                # Novo Upper Offset para que o topo do Ambiente coincida com o elemento encontrado
                upper_level = doc.GetElement(room.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).AsElementId())
                if not upper_level:
                    continue

                novo_offset = elemento['elevacao'] - upper_level.Elevation
                offset_atual = room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble()

                # Verificacao de Sincronia
                if abs(offset_atual - novo_offset) > TOL_AJUSTE:
                    detalhe = "{} | Atual: {:.2f} m -> Novo: {:.2f} m  (Dif: {:.2f} m)".format(
                        origem,
                        convert_internal_to_m(offset_atual + upper_level.Elevation),
                        convert_internal_to_m(elemento['elevacao']),
                        convert_internal_to_m(abs(offset_atual - novo_offset))
                    )

                    candidatos.append({
                        'room': room,
                        'nome': "{} - {}".format(numero_ambiente, nome_ambiente),
                        'detalhe': detalhe,
                        'novo_offset': novo_offset
                    })

            except Exception as e:
                print("Erro no Ambiente {}: {}".format(room.Id, e))

    # ETAPA 2: Revisao em lote dos ajustes de Limite Superior
    if not candidatos:
        resumo = "Nenhum ajuste de Limite Superior necessario."
        if sem_elemento:
            resumo += "\n\nNenhum Forro/Piso/Telhado encontrado sobre {} ambiente(s):\n- {}".format(
                len(sem_elemento), "\n- ".join(sem_elemento)
            )
        resumo += aviso_parametros
        forms.alert(resumo, title=__title__)
        return

    selecionados = review_room_adjustments(
        candidatos,
        title=__title__,
        label="Ambientes com divergencia de Limite Superior:"
    )

    if not selecionados:
        return

    # ETAPA 3: Aplica os ajustes selecionados
    ajustados = 0
    with ef_Transaction(doc, 'NnBim - Ajustar Limite Superior', debug=True):
        for item in selecionados:
            room = item.element
            p_offset_param = room.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET)
            p_offset_param.Set(item.novo_offset)
            ajustados += 1

    # Resumo final
    resumo = "Ambientes ajustados: {}".format(ajustados)
    if sem_elemento:
        resumo += "\n\nNenhum Forro/Piso/Telhado encontrado sobre {} ambiente(s):\n- {}".format(
            len(sem_elemento), "\n- ".join(sem_elemento)
        )
    resumo += aviso_parametros

    forms.alert(resumo, title=__title__)


if __name__ == '__main__':
    main()
