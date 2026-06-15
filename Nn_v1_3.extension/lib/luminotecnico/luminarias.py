# -*- coding: utf-8 -*-
"""
Título: NnBim — Luminárias Disponíveis
Descrição: Tabela de luminárias disponíveis para cálculo luminotécnico.
           Organizada por tipo em blocos separados.
Instruções de Uso: Para adicionar nova luminária, copie uma entrada
                   existente dentro do bloco correspondente e altere
                   os dados. Para adicionar novo tipo, crie um novo
                   bloco e adicione-o em TODAS_LUMINARIAS.
"""

# =============================================================================
# PLACA LED EMBUTIR
# Luminária de embutir no forro — painel LED integrado com difusor
# Para adicionar nova placa: copie uma entrada e altere os dados
# =============================================================================

PLACAS_LED_EMBUTIR = [
    {
        "descricao": "Placa LED 10W",
        "tipo": "Placa LED embutir",
        "potencia_W": 10,
        "fluxo_lm": 700,
        "eficacia": 70.0,
        "obs": ""
    },
    {
        "descricao": "Placa LED 14W",
        "tipo": "Placa LED embutir",
        "potencia_W": 14,
        "fluxo_lm": 1000,
        "eficacia": 71.4,
        "obs": ""
    },
    {
        "descricao": "Placa LED 20W",
        "tipo": "Placa LED embutir",
        "potencia_W": 20,
        "fluxo_lm": 1700,
        "eficacia": 85.0,
        "obs": ""
    },
    {
        "descricao": "Placa LED 32W",
        "tipo": "Placa LED embutir",
        "potencia_W": 32,
        "fluxo_lm": 2800,
        "eficacia": 87.5,
        "obs": ""
    },
    {
        "descricao": "Placa LED 40W",
        "tipo": "Placa LED embutir",
        "potencia_W": 40,
        "fluxo_lm": 3900,
        "eficacia": 97.5,
        "obs": "Painel 62,5x62,5cm com difusor"
    },
    {
        "descricao": "Placa LED 50W",
        "tipo": "Placa LED embutir",
        "potencia_W": 50,
        "fluxo_lm": 4500,
        "eficacia": 90.0,
        "obs": ""
    },
]

# =============================================================================
# LUMINÁRIAS DO LEVANTAMENTO PRF
# =============================================================================

LUMINARIAS_PRF = [
    {
        "descricao": "Plafon Embutir Retangular 2x20W (Tubular LED)",
        "tipo": "Luminária de embutir",
        "potencia_W": 40,
        "fluxo_lm": 3700,
        "eficacia": 92.5,
        "obs": "Equipada com 2 lâmpadas tubulares LED de 20W"
    },
    {
        "descricao": "Plafon de Sobrepor LED 12W",
        "tipo": "Plafon LED sobrepor",
        "potencia_W": 12,
        "fluxo_lm": 900,
        "eficacia": 75.0,
        "obs": "Painel de sobrepor"
    },
    {
        "descricao": "Refletor LED 50W Alumínio",
        "tipo": "Refletor Externo",
        "potencia_W": 50,
        "fluxo_lm": 4000,
        "eficacia": 80.0,
        "obs": "Refletor com alça/suporte em alumínio"
    },
    {
        "descricao": "Pública Philips/G-Light 150W",
        "tipo": "Iluminação Pública",
        "potencia_W": 150,
        "fluxo_lm": 15495,
        "eficacia": 103.3,
        "obs": "IP66, IK09, Temp. cor 5000k"
    },
    {
        "descricao": "Pública Philips BGP323 250W",
        "tipo": "Iluminação Pública",
        "potencia_W": 250,
        "fluxo_lm": 27775,
        "eficacia": 111.1,
        "obs": "IP66, IK09, Temp. cor 5000k"
    },
    {
        "descricao": "Luminária de Emergência 30 LEDs (2W)",
        "tipo": "Emergência",
        "potencia_W": 2,
        "fluxo_lm": 100,
        "eficacia": 50.0,
        "obs": "Uso apenas para rotas de fuga"
    }
]

# =============================================================================
# OUTROS TIPOS — adicione novos blocos aqui quando necessário
# =============================================================================

# =============================================================================
# LISTA UNIFICADA — consultada pelos scripts
# =============================================================================

TODAS_LUMINARIAS = PLACAS_LED_EMBUTIR + LUMINARIAS_PRF