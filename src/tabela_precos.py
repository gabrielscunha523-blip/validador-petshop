# tabela_precos.py - Tabela de preços de referência do pet shop
# Esta tabela serve para o validador comparar o valor cobrado
# com o valor "tabelado". Se a diferença for muito grande, ele avisa.

# Estrutura: TABELA[servico][porte] = preco em reais
TABELA_PRECOS = {
    'Banho': {
        'P': 40.00,
        'M': 55.00,
        'G': 70.00,
        'GG': 90.00,
    },
    'Tosa': {
        'P': 50.00,
        'M': 65.00,
        'G': 85.00,
        'GG': 110.00,
    },
    'Banho+Tosa': {
        'P': 70.00,
        'M': 95.00,
        'G': 120.00,
        'GG': 150.00,
    },
    'Tosa Higienica': {
        'P': 30.00,
        'M': 40.00,
        'G': 50.00,
        'GG': 65.00,
    },
    'Hidratacao': {
        'P': 25.00,
        'M': 35.00,
        'G': 45.00,
        'GG': 60.00,
    },
}

# Tolerancia: o quanto o valor cobrado pode variar do tabelado
# sem o sistema reclamar. 0.30 = 30% para mais ou para menos.
TOLERANCIA_PRECO = 0.30

# Valores aceitos em cada coluna de dominio fechado
SERVICOS_VALIDOS = list(TABELA_PRECOS.keys())
PORTES_VALIDOS = ['P', 'M', 'G', 'GG']
STATUS_VALIDOS = ['Pago', 'Pendente']
FORMAS_PAGAMENTO_VALIDAS = ['Dinheiro', 'Pix', 'Cartao', 'Fiado', '']
