# -*- coding: utf-8 -*-
"""
Título: NnBim — Tabela NBR ISO/CIE 8995-1:2013
Descrição: Ambientes dos grupos 1, 22, 23, 24, 25, 26, 27, 28, 29 e 30
           da Seção 5 da ABNT NBR ISO/CIE 8995-1:2013.
           Dados: Em (lux), UGR limite, IRC mínimo (Ra),
           altura do plano de trabalho (HT) e palavras-chave.
Instruções de Uso: Não altere os dados da norma.
                   Palavras-chave podem ser complementadas conforme
                   necessidade do projeto.
"""

# =============================================================================
# ESTRUTURA DE CADA ENTRADA:
#   grupo          : número e nome do grupo conforme Seção 5
#   atividade      : descrição da tarefa ou atividade
#   Em             : iluminância mantida (lux)
#   UGR            : índice limite de ofuscamento unificado
#   IRC            : índice de reprodução de cor mínimo (Ra)
#   HT             : altura do plano de trabalho (m)
#   palavras_chave : lista de termos para identificar o Room pelo nome
#   obs            : observações da norma
# =============================================================================

AMBIENTES_8995 = [

    # ─────────────────────────────────────────────────────────────────────────
    # 1. ÁREAS GERAIS DA EDIFICAÇÃO
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Saguão de entrada",
        "Em": 100, "UGR": 22, "IRC": 60, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Sala de espera",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Áreas de circulação e corredores",
        "Em": 100, "UGR": 28, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Nas entradas e saídas, estabelecer zona de transição."
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Escadas, escadas rolantes e esteiras rolantes",
        "Em": 150, "UGR": 25, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Rampas de carregamento",
        "Em": 150, "UGR": 25, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Refeitório/Cantinas",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Salas de descanso",
        "Em": 100, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Salas para exercícios físicos",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Vestiários, banheiros, toaletes",
        "Em": 200, "UGR": 25, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Enfermaria",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Salas para atendimento médico",
        "Em": 500, "UGR": 16, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Tcp no mínimo 4 000 K."
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Estufas, sala dos disjuntores",
        "Em": 200, "UGR": 25, "IRC": 60, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Correios, quadros de distribuição",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Depósito, estoques, câmara fria",
        "Em": 100, "UGR": 25, "IRC": 60, "HT": 0.75,
        "palavras_chave": [],
        "obs": "200 lux, se forem continuamente ocupados."
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Expedição",
        "Em": 300, "UGR": 25, "IRC": 60, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "1. Áreas gerais da edificação",
        "atividade": "Estação de controle",
        "Em": 150, "UGR": 22, "IRC": 60, "HT": 0.75,
        "palavras_chave": [],
        "obs": "200 lux se forem continuamente ocupadas."
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 22. ESCRITÓRIOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "22. Escritórios",
        "atividade": "Arquivamento, cópia, circulação etc.",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Escrever, teclar, ler, processar dados",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Para trabalho com VDT, ver 4.10."
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Desenho técnico",
        "Em": 750, "UGR": 16, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Estações de projeto assistido por computador (CAD)",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Para trabalho com VDT, ver 4.10."
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Salas de reunião e conferência",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Recomenda-se que a iluminação seja controlável."
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Recepção",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "22. Escritórios",
        "atividade": "Arquivos",
        "Em": 200, "UGR": 25, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 23. VAREJO
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "23. Varejo",
        "atividade": "Área de vendas pequena",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "23. Varejo",
        "atividade": "Área de vendas grande",
        "Em": 500, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "23. Varejo",
        "atividade": "Área da caixa registradora",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "23. Varejo",
        "atividade": "Mesa do empacotador",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 24. RESTAURANTES E HOTÉIS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Recepção/caixa/portaria",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Cozinha",
        "Em": 500, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Restaurante, sala de jantar, sala de eventos",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Recomenda-se iluminação para criar ambiente íntimo."
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Restaurante self-service",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Bufê",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Salas de conferência",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Recomenda-se que a iluminação seja controlável."
    },
    {
        "grupo": "24. Restaurantes e hotéis",
        "atividade": "Corredores",
        "Em": 100, "UGR": 25, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Durante o período da noite são aceitáveis baixos níveis."
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 25. LOCAIS DE ENTRETENIMENTO
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "25. Locais de entretenimento",
        "atividade": "Teatros e salas de concerto",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "25. Locais de entretenimento",
        "atividade": "Salas com multiuso",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "25. Locais de entretenimento",
        "atividade": "Salas de ensaio, camarins",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Iluminação do espelho isenta de ofuscamento para maquiagem."
    },
    {
        "grupo": "25. Locais de entretenimento",
        "atividade": "Museus (em geral)",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Proteção contra efeitos de radiação."
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 26. BIBLIOTECAS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "26. Bibliotecas",
        "atividade": "Estantes",
        "Em": 200, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "26. Bibliotecas",
        "atividade": "Área de leitura",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "26. Bibliotecas",
        "atividade": "Bibliotecárias",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 27. ESTACIONAMENTOS PÚBLICOS (INTERNOS)
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "27. Estacionamentos públicos (internos)",
        "atividade": "Rampas de entrada e saída (durante o dia)",
        "Em": 300, "UGR": 25, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": "As cores para segurança devem ser reconhecíveis."
    },
    {
        "grupo": "27. Estacionamentos públicos (internos)",
        "atividade": "Rampas de entrada e saída (durante a noite)",
        "Em": 75, "UGR": 25, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": "As cores para segurança devem ser reconhecíveis."
    },
    {
        "grupo": "27. Estacionamentos públicos (internos)",
        "atividade": "Pistas de tráfego",
        "Em": 75, "UGR": 25, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": "As cores para segurança devem ser reconhecíveis."
    },
    {
        "grupo": "27. Estacionamentos públicos (internos)",
        "atividade": "Estacionamento",
        "Em": 75, "UGR": 28, "IRC": 40, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Iluminância vertical elevada aumenta reconhecimento de faces."
    },
    {
        "grupo": "27. Estacionamentos públicos (internos)",
        "atividade": "Guichê",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Evitar reflexões nas janelas."
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 28. CONSTRUÇÕES EDUCACIONAIS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Brinquedoteca",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Berçário",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de aula, salas de aulas particulares",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Recomenda-se que a iluminação seja controlável."
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de aulas noturnas, classes e educação de adultos",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Sala de leitura",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de arte e artesanato",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de desenho técnico",
        "Em": 750, "UGR": 16, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de aplicação e laboratórios",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Oficina de ensino",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de ensino de música",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de ensino de computador",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Para trabalho com VDT, ver 4.10."
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas dos professores",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "28. Construções educacionais",
        "atividade": "Salas de esportes, ginásios e piscinas",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 29. LOCAIS DE ASSISTÊNCIA MÉDICA
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Salas de espera",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Iluminância ao nível do piso."
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Corredores: durante o dia",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Iluminância ao nível do piso."
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Quartos com claridade",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Iluminância ao nível do piso."
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Escritório dos funcionários",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Sala dos funcionários",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Enfermarias — iluminação em geral",
        "Em": 100, "UGR": 19, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": "Iluminância ao nível do piso."
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Exames e tratamento",
        "Em": 1000, "UGR": 19, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Sala de exames em geral",
        "Em": 500, "UGR": 19, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Sala de cirurgia",
        "Em": 1000, "UGR": 19, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "UTI — iluminação em geral",
        "Em": 100, "UGR": 19, "IRC": 90, "HT": 0.00,
        "palavras_chave": [],
        "obs": "No nível do piso."
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Salas pré-operatórias e salas de recuperação",
        "Em": 500, "UGR": 19, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Salas de esterilização",
        "Em": 300, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "29. Locais de assistência médica",
        "atividade": "Salas de autópsia e necrotérios",
        "Em": 500, "UGR": 19, "IRC": 90, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },

    # ─────────────────────────────────────────────────────────────────────────
    # 30. AEROPORTOS
    # ─────────────────────────────────────────────────────────────────────────
    {
        "grupo": "30. Aeroportos",
        "atividade": "Saguões de embarque e desembarque",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.00,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "30. Aeroportos",
        "atividade": "Balcão de informações, check-in",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Para VDT, ver 4.10."
    },
    {
        "grupo": "30. Aeroportos",
        "atividade": "Alfândega e balcão de controle do passaporte",
        "Em": 500, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "É importante a iluminância vertical."
    },
    {
        "grupo": "30. Aeroportos",
        "atividade": "Salas de espera",
        "Em": 200, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
    {
        "grupo": "30. Aeroportos",
        "atividade": "Áreas da verificação de segurança",
        "Em": 300, "UGR": 19, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": "Para VDT, ver 4.10."
    },
    {
        "grupo": "30. Aeroportos",
        "atividade": "Hangares de reparos e testes",
        "Em": 500, "UGR": 22, "IRC": 80, "HT": 0.75,
        "palavras_chave": [],
        "obs": ""
    },
]

# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def listar_grupos():
    """Retorna lista de grupos únicos da NBR 8995-1."""
    grupos = []
    for a in AMBIENTES_8995:
        if a["grupo"] not in grupos:
            grupos.append(a["grupo"])
    return grupos


def buscar_por_grupo(grupo):
    """Retorna todas as atividades de um grupo."""
    return [a for a in AMBIENTES_8995 if a["grupo"] == grupo]


def buscar_por_palavras_chave(nome_room):
    """
    Busca o ambiente pelo nome do Room usando palavras-chave.
    Retorna o primeiro que tiver correspondência.
    """
    nome_lower = nome_room.lower()
    for ambiente in AMBIENTES_8995:
        for palavra in ambiente["palavras_chave"]:
            if palavra.lower() in nome_lower:
                return ambiente
    return None