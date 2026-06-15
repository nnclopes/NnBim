# -*- coding: utf-8 -*-
"""
Título: NnBim — Funções de Consulta Luminotécnica
Descrição: Funções centralizadas de cálculo e busca nas tabelas
           das normas NBR 8995-1:2013 e NBR 5413:1992.
           Todos os botões do painel Luminotécnico usam este módulo.
Instruções de Uso: Importe as funções necessárias nos scripts:
                   from luminotecnico.consulta import calcular_n
                   from luminotecnico.consulta import sugerir_luminaria
"""

import math
from luminotecnico.nbr_8995  import AMBIENTES_8995,  buscar_por_palavras_chave as buscar_8995
from luminotecnico.nbr_5413  import AMBIENTES_5413,  buscar_por_palavras_chave as buscar_5413
from luminotecnico.luminarias import TODAS_LUMINARIAS

# =============================================================================
# 1. CLASSIFICAÇÃO DE ROOMS
# =============================================================================

def classificar_por_nome(nome_room, norma="8995"):
    """
    Busca o ambiente mais adequado para o Room pelo nome.

    Parâmetros:
        nome_room (str): Nome do Room no Revit
        norma     (str): "8995" ou "5413"

    Retorna:
        dict com os dados do ambiente, ou None se não encontrado
    """
    if norma == "8995":
        return buscar_8995(nome_room)
    elif norma == "5413":
        return buscar_5413(nome_room)
    return None


def listar_grupos(norma="8995"):
    """
    Retorna lista de grupos da norma selecionada.

    Parâmetros:
        norma (str): "8995" ou "5413"

    Retorna:
        list de strings com os nomes dos grupos
    """
    if norma == "8995":
        from luminotecnico.nbr_8995 import listar_grupos as lg
    else:
        from luminotecnico.nbr_5413 import listar_grupos as lg
    return lg()


def listar_atividades_por_grupo(grupo, norma="8995"):
    """
    Retorna todas as atividades de um grupo da norma.

    Parâmetros:
        grupo (str): Nome do grupo
        norma (str): "8995" ou "5413"

    Retorna:
        list de dicts com os dados de cada atividade
    """
    if norma == "8995":
        from luminotecnico.nbr_8995 import buscar_por_grupo as bg
    else:
        from luminotecnico.nbr_5413 import buscar_por_grupo as bg
    return bg(grupo)


# =============================================================================
# 2. CÁLCULOS LUMINOTÉCNICOS
# =============================================================================

def calcular_k(comprimento, largura, altura_util):
    """
    Calcula o Índice do Local (K).

    Fórmula: K = (C × L) / (HU × (C + L))

    Parâmetros:
        comprimento  (float): Comprimento do ambiente (m)
        largura      (float): Largura do ambiente (m)
        altura_util  (float): HU = Altura_Luminaria - HT

    Retorna:
        float: Índice K arredondado em 3 casas
    """
    if altura_util <= 0 or (comprimento + largura) <= 0:
        return 0.0
    k = (comprimento * largura) / (altura_util * (comprimento + largura))
    return round(k, 3)


def calcular_n(Em, area, fluxo_lm, FU, FM):
    """
    Calcula a quantidade de luminárias pelo método dos lúmens.

    Fórmula: N = (Em × A) / (φ × FU × FM)

    Parâmetros:
        Em       (float): Iluminância mantida (lux)
        area     (float): Área do ambiente (m²)
        fluxo_lm (float): Fluxo luminoso da luminária (lm)
        FU       (float): Fator de utilização
        FM       (float): Fator de manutenção

    Retorna:
        dict:
            N_calculado (float): resultado exato
            N_minimo    (int)  : arredondado para cima
            lux_real    (float): iluminância com N_minimo
    """
    if fluxo_lm <= 0 or FU <= 0 or FM <= 0 or area <= 0:
        return {"N_calculado": 0, "N_minimo": 0, "lux_real": 0}

    n_calc   = (Em * area) / (fluxo_lm * FU * FM)
    n_minimo = math.ceil(n_calc)
    lux_real = (n_minimo * fluxo_lm * FU * FM) / area

    return {
        "N_calculado": round(n_calc, 4),
        "N_minimo":    n_minimo,
        "lux_real":    round(lux_real, 1)
    }


def calcular_lux_real(N_adotado, fluxo_lm, FU, FM, area):
    """
    Calcula a iluminância real com o N adotado.

    Fórmula: E_real = (N × φ × FU × FM) / A

    Parâmetros:
        N_adotado (int)  : Luminárias instaladas
        fluxo_lm  (float): Fluxo luminoso (lm)
        FU        (float): Fator de utilização
        FM        (float): Fator de manutenção
        area      (float): Área do ambiente (m²)

    Retorna:
        float: Iluminância real em lux
    """
    if area <= 0:
        return 0.0
    return round((N_adotado * fluxo_lm * FU * FM) / area, 1)


# =============================================================================
# 3. VERIFICAÇÃO NBR
# =============================================================================

def verificar_status_nbr(lux_real, Em, tol_min=0.9, tol_max=1.2):
    """
    Verifica se a iluminância real atende à NBR.

    Tolerância padrão: ±10% conforme NBR 8995-1 item 6.7

    Parâmetros:
        lux_real (float): Iluminância real calculada
        Em       (float): Iluminância mantida exigida
        tol_min  (float): Fator mínimo — padrão 0.9
        tol_max  (float): Fator máximo — padrão 1.2

    Retorna:
        str: "OK" / "INSUFICIENTE" / "ACIMA"
    """
    if lux_real < Em * tol_min:
        return "INSUFICIENTE"
    elif lux_real > Em * tol_max:
        return "ACIMA"
    return "OK"


# =============================================================================
# 4. SUGESTÃO DE LUMINÁRIA
# =============================================================================

def sugerir_luminaria(Em, area, FU, FM):
    """
    Sugere a menor luminária LED que atende o Em.

    Percorre TODAS_LUMINARIAS da menor para a maior
    e retorna a primeira que atende com N razoável.

    Parâmetros:
        Em   (float): Iluminância mantida exigida (lux)
        area (float): Área do ambiente (m²)
        FU   (float): Fator de utilização
        FM   (float): Fator de manutenção

    Retorna:
        dict:
            luminaria   (dict) : dados da luminária sugerida
            N_calculado (float): quantidade calculada
            N_minimo    (int)  : arredondado para cima
            lux_real    (float): iluminância com N_minimo
        None se nenhuma atender
    """
    for lum in TODAS_LUMINARIAS:
        resultado = calcular_n(Em, area, lum["fluxo_lm"], FU, FM)
        n_min = resultado["N_minimo"]
        if n_min > 0:
            return {
                "luminaria":   lum,
                "N_calculado": resultado["N_calculado"],
                "N_minimo":    n_min,
                "lux_real":    resultado["lux_real"]
            }
    return None


def sugerir_luminaria_para_n_adotado(Em, area, FU, FM, N_adotado):
    """
    Sugere a menor luminária que atende o Em com N já fixo.

    Útil quando o layout do forro não permite alterar
    a quantidade de luminárias.

    Parâmetros:
        Em        (float): Iluminância mantida exigida (lux)
        area      (float): Área do ambiente (m²)
        FU        (float): Fator de utilização
        FM        (float): Fator de manutenção
        N_adotado (int)  : Quantidade já definida

    Retorna:
        dict com luminária sugerida e lux resultante,
        ou None se nenhuma atender
    """
    for lum in TODAS_LUMINARIAS:
        lux = calcular_lux_real(N_adotado, lum["fluxo_lm"], FU, FM, area)
        if verificar_status_nbr(lux, Em) in ("OK", "ACIMA"):
            return {
                "luminaria": lum,
                "lux_real":  lux,
                "status":    verificar_status_nbr(lux, Em)
            }
    return None