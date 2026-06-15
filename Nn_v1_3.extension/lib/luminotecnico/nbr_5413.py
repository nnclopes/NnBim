# -*- coding: utf-8 -*-
"""
Título: NnBim — Tabela NBR 5413:1992
Descrição: Tabela de iluminâncias por tipo de atividade da
           ABNT NBR 5413:1992 — Iluminância de interiores.
           Valor armazenado: valor médio da faixa (min - MED - max).
Instruções de Uso: Não altere os dados da norma.
                   Palavras-chave podem ser complementadas conforme
                   necessidade do projeto.
"""

# =============================================================================
# ESTRUTURA DE CADA ENTRADA:
#   grupo          : número e nome do grupo conforme item 5.3
#   atividade      : descrição da tarefa ou atividade
#   Em             : valor médio da faixa (lux)
#   Em_min         : valor mínimo da faixa (lux)
#   Em_max         : valor máximo da faixa (lux)
#   HT             : altura do plano de trabalho (m)
#   palavras_chave : lista de termos para identificar o Room pelo nome
#   obs            : observações
# =============================================================================

AMBIENTES_5413 = [

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.2 AUDITÓRIOS E ANFITEATROS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.2 Auditórios e anfiteatros",
        "atividade": "Tribuna",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.2 Auditórios e anfiteatros",
        "atividade": "Plateia",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.2 Auditórios e anfiteatros",
        "atividade": "Sala de espera",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.2 Auditórios e anfiteatros",
        "atividade": "Bilheterias",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.3 BANCOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Atendimento ao público",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Estatística e contabilidade",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Salas de gerentes",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Salas de recepção",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Salas de conferências",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Guichês",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Arquivos",
        "Em": 300, "Em_min": 200, "Em_max": 500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Saguão",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.3 Bancos",
        "atividade": "Cantinas",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.5 BIBLIOTECAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.5 Bibliotecas",
        "atividade": "Sala de leitura",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.5 Bibliotecas",
        "atividade": "Recinto das estantes",
        "Em": 300, "Em_min": 200, "Em_max": 500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.5 Bibliotecas",
        "atividade": "Fichário",
        "Em": 300, "Em_min": 200, "Em_max": 500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.8 CINEMAS E TEATROS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.8 Cinemas e teatros",
        "atividade": "Sala de espetáculos — durante o intervalo",
        "Em": 50, "Em_min": 30, "Em_max": 75, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.8 Cinemas e teatros",
        "atividade": "Sala de espera, foyer",
        "Em": 100, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.8 Cinemas e teatros",
        "atividade": "Bilheterias",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.10 CORREDORES E ESCADAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.10 Corredores e escadas",
        "atividade": "Geral",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.13 ESCOLAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Salas de aulas",
        "Em": 300, "Em_min": 200, "Em_max": 500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Quadros negros",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 1.50,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Laboratórios — geral",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Sala de desenho",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Sala de reuniões",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.13 Escolas",
        "atividade": "Salas de educação física",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.14 ESCRITÓRIOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.14 Escritórios",
        "atividade": "Registros, cartografia etc.",
        "Em": 1000, "Em_min": 750, "Em_max": 1500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.14 Escritórios",
        "atividade": "Desenho, engenharia mecânica e arquitetura",
        "Em": 1000, "Em_min": 750, "Em_max": 1500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.14 Escritórios",
        "atividade": "Desenho decorativo e esboço",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.28 HOSPITAIS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Sala dos médicos ou enfermeiras — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Sala dos médicos ou enfermeiras — mesa de trabalho",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Farmácia — geral",
        "Em": 150, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Banheiros — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Pronto-socorro — geral",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Corredores e escadas",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Sala de operação — iluminação geral",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.28 Hospitais",
        "atividade": "Quartos particulares para pacientes — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.29 HOTÉIS E RESTAURANTES
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.29 Hotéis e restaurantes",
        "atividade": "Banheiros",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.29 Hotéis e restaurantes",
        "atividade": "Corredores e escadas",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.29 Hotéis e restaurantes",
        "atividade": "Cozinha — geral",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.29 Hotéis e restaurantes",
        "atividade": "Restaurantes",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.29 Hotéis e restaurantes",
        "atividade": "Portaria e recepção",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.30 IGREJAS E TEMPLOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.30 Igrejas e templos",
        "atividade": "Nave, entrada, auditórios — com ofício",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.30 Igrejas e templos",
        "atividade": "Púlpito — com ofício",
        "Em": 300, "Em_min": 200, "Em_max": 500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.56 LAVATÓRIOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.56 Lavatórios",
        "atividade": "Geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.57 LOCAIS DE ARMAZENAMENTO
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.57 Locais de armazenamento",
        "atividade": "Armazéns gerais (não usados frequentemente)",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.57 Locais de armazenamento",
        "atividade": "Armazéns de fábricas — volumes grandes",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.57 Locais de armazenamento",
        "atividade": "Armazéns de fábricas — volumes pequenos",
        "Em": 200, "Em_min": 150, "Em_max": 300, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.58 LOJAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.58 Lojas",
        "atividade": "Vitrines e balcões — centros comerciais",
        "Em": 1000, "Em_min": 750, "Em_max": 1500, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.58 Lojas",
        "atividade": "Interior de loja de artigos diversos",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.61 MUSEUS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.61 Museus",
        "atividade": "Geral",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.61 Museus",
        "atividade": "Esculturas e outros objetos",
        "Em": 500, "Em_min": 300, "Em_max": 750, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 5.3.65 RESIDÊNCIAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "5.3.65 Residências",
        "atividade": "Salas de estar — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.65 Residências",
        "atividade": "Cozinhas — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.65 Residências",
        "atividade": "Quartos de dormir — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.65 Residências",
        "atividade": "Hall, escadas, despensas, garagens — geral",
        "Em": 100, "Em_min": 75, "Em_max": 150, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "5.3.65 Residências",
        "atividade": "Banheiros — geral",
        "Em": 150, "Em_min": 100, "Em_max": 200, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
]

# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def listar_grupos():
    """Retorna lista de grupos únicos da NBR 5413."""
    grupos = []
    for a in AMBIENTES_5413:
        if a["grupo"] not in grupos:
            grupos.append(a["grupo"])
    return grupos


def buscar_por_grupo(grupo):
    """Retorna todas as atividades de um grupo."""
    return [a for a in AMBIENTES_5413 if a["grupo"] == grupo]


def buscar_por_palavras_chave(nome_room):
    """
    Busca o ambiente pelo nome do Room usando palavras-chave.
    Retorna o primeiro que tiver correspondência.
    """
    nome_lower = nome_room.lower()
    for ambiente in AMBIENTES_5413:
        for palavra in ambiente["palavras_chave"]:
            if palavra.lower() in nome_lower:
                return ambiente
    return None