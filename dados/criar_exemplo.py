# criar_exemplo.py - Cria uma planilha de exemplo para testar o validador
# Esta planilha simula um dia de atendimentos no pet shop,
# COM alguns erros propositais para o validador pegar.

import pandas as pd

# Dados de exemplo (alguns ok, outros com erros)
dados_exemplo = {
    'Data_Atendimento': [
        '2026-05-01', '2026-05-01', '2026-05-02', '2026-05-02',
        '2026-05-03', '2026-05-03', '2026-05-04', '2026-05-04',
        '2026-05-05', '2026-05-05', '2026-05-06', '2026-05-06',
    ],
    'Nome_Tutor': [
        'Maria Silva', 'Joao Souza', 'Ana Costa', 'Maria Silva',
        'Pedro Lima', '', 'Carla Mendes', 'Joao Souza',
        'Ana Costa', 'Roberto Alves', 'Maria Silva', 'Pedro Lima',
    ],
    'Telefone_Tutor': [
        '(11) 98765-1111', '(11) 99876-2222', '(21) 97654-3333', '(11) 98765-1111',
        '(31) 96543-4444', '(11) 95432-5555', 'abc123', '(11) 99876-2222',
        '(21) 97654-3333', '(41) 94321-6666', '(11) 98765-1111', '(31) 96543-4444',
    ],
    'Nome_Pet': [
        'Rex', 'Mel', 'Thor', 'Luna',
        'Bidu', 'Pingo', 'Nina', 'Mel',
        'Thor', 'Bob', 'Rex', 'Bidu',
    ],
    'Porte': [
        'M', 'P', 'GG', 'M',
        'G', 'P', 'M', 'P',
        'GG', 'XL', 'M', 'G',
    ],
    'Servico': [
        'Banho', 'Banho+Tosa', 'Tosa', 'Hidratacao',
        'Banho+Tosa', 'Banho', 'Tosa Higienica', 'Banho+Tosa',
        'Tosa', 'Spa', 'Banho', 'Banho+Tosa',
    ],
    'Valor_Servico': [
        55.00, 70.00, 110.00, 35.00,
        120.00, 40.00, 40.00, 0.00,
        110.00, 200.00, 55.00, 250.00,
    ],
    'Status_Pagamento': [
        'Pago', 'Pago', 'Pendente', 'Pago',
        'Pendente', 'Pago', 'Pago', 'Pendente',
        'Pendente', 'Pago', 'Pago', 'Pendente',
    ],
    'Forma_Pagamento': [
        'Pix', 'Dinheiro', '', 'Cartao',
        '', 'Pix', 'Dinheiro', '',
        '', 'Pix', 'Cartao', '',
    ],
    'Data_Pagamento': [
        '2026-05-01', '2026-05-01', '', '2026-05-02',
        '', '2026-05-03', '2026-05-04', '',
        '', '2026-05-05', '2026-05-06', '',
    ],
}

# Cria o DataFrame
df = pd.DataFrame(dados_exemplo)

# Garante que a pasta existe
import os
os.makedirs('dados', exist_ok=True)

# Salva como Excel
df.to_excel('dados/comandas_input.xlsx', index=False)

print("Arquivo 'comandas_input.xlsx' criado com sucesso!")
print(f"Total de linhas: {len(df)}")
print()
print("Pre-visualizacao dos dados:")
print(df.to_string())
