# Validador de Comandas - TechPet

Sistema em Python para validação de comandas de pet shop e gestão de contas a receber. Lê uma planilha Excel com atendimentos do dia (banho, tosa, hidratação etc.), aponta erros e gera um relatório financeiro completo com lista de inadimplentes.

Desenvolvido como **Atividade de Extensão Universitária (2026.1)** em parceria com o **Momento Pet** (Suzano-SP), microempreendimento do Sr. Fernando, com o objetivo de qualificar o controle financeiro do negócio e reduzir perdas por inadimplência.

---

## Objetivo

Microempreendedores de pet shop costumam anotar comandas em caderno ou em planilhas soltas, sem qualquer validação. Erros comuns:

- Cobrar valor errado para o porte do pet
- Esquecer de marcar quem pagou e quem ficou pendente
- Não saber, no fim do mês, **quanto** tem a receber e **de quem**
- Telefone do tutor mal cadastrado, impossibilitando a cobrança

Este programa lê a planilha de comandas, **valida 11 regras de negócio** e gera um relatório multi-aba com:

- Resumo geral da operação
- Lista de problemas detectados (linha por linha)
- Registros válidos
- **Contas a Receber consolidadas por tutor** (com aging)
- Resumo financeiro (faturamento, recebido, pendente, ticket médio)

---

## Funcionalidades

### Validações implementadas

1. Data de atendimento preenchida e em formato válido
2. Nome do tutor preenchido
3. Telefone do tutor no formato `(11) 99999-9999`
4. Nome do pet preenchido
5. Porte dentro do domínio `P / M / G / GG`
6. Serviço dentro da lista de serviços oferecidos
7. Valor do serviço maior que zero
8. Valor cobrado coerente com a tabela de preços (tolerância de 30%)
9. Status de pagamento no domínio `Pago / Pendente`
10. Forma de pagamento válida (`Dinheiro`, `Pix`, `Cartao`, `Fiado`) quando status = Pago
11. Coerência entre status e data de pagamento

### Saídas geradas

O programa produz um arquivo Excel com 5 abas:

| Aba | Conteúdo |
|---|---|
| **Resumo** | Métricas gerais: total processado, problemas encontrados, registros válidos |
| **Problemas** | Linhas com erro, indicando qual regra foi violada |
| **Dados_OK** | Registros que passaram em todas as validações |
| **Contas_a_Receber** | Agregação por tutor com valor em aberto e dias em atraso |
| **Resumo_Financeiro** | Faturamento bruto, recebido, pendente, ticket médio, distribuição por serviço |

---

## Pré-requisitos

- Python 3.10 ou superior
- pip
- Sistema operacional: Windows, macOS ou Linux

---

## Instalação

Clone o repositório:

```bash
git clone https://github.com/gabrielscunha523-blip/validador-petshop.git
cd validador-petshop
```

Crie e ative um ambiente virtual (recomendado):

```bash
# Windows (PowerShell)
python -m venv venv
.\venv\Scripts\Activate.ps1

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

Instale as dependências:

```bash
pip install pandas openpyxl xlsxwriter
```

---

## Como usar

### 1. Gerar uma planilha de exemplo (opcional)

Se você ainda não tem uma planilha de comandas, gere uma de teste:

```bash
python dados/criar_exemplo.py
```

Isso cria `dados/comandas_input.xlsx` com 12 registros (alguns com erros propositais).

### 2. Rodar o validador

```bash
python src/main.py
```

Por padrão o programa lê `dados/comandas_input.xlsx` e gera `relatorio_validacao.xlsx` na pasta atual.

### Parâmetros opcionais

```bash
python src/main.py --entrada caminho/da/planilha.xlsx --saida nome_do_relatorio.xlsx
```

---

## Estrutura da planilha de entrada

A planilha deve ter as seguintes colunas (nessa ordem, sem acento):

| Coluna | Tipo | Exemplo |
|---|---|---|
| Data_Atendimento | data | 2026-05-15 |
| Nome_Tutor | texto | Maria Silva |
| Telefone_Tutor | texto | (11) 98765-1111 |
| Nome_Pet | texto | Thor |
| Porte | P/M/G/GG | M |
| Servico | texto | Banho+Tosa |
| Valor_Servico | número | 95.00 |
| Status_Pagamento | Pago/Pendente | Pago |
| Forma_Pagamento | texto | Pix |
| Data_Pagamento | data ou vazio | 2026-05-15 |

---

## Estrutura do projeto

```
validador-petshop/
├── src/
│   ├── main.py              # Ponto de entrada (argparse, orquestração)
│   ├── tabela_precos.py     # Tabela de preços e listas de domínio
│   ├── validador.py         # 11 regras de validação
│   └── gerar_relatorio.py   # Geração do Excel multi-aba
├── dados/
│   ├── criar_exemplo.py     # Gerador de planilha de teste
│   └── comandas_input.xlsx  # Planilha de entrada (gerada)
├── .gitignore
└── README.md
```

---

## Tabela de preços de referência

| Serviço | P | M | G | GG |
|---|---:|---:|---:|---:|
| Banho | 40,00 | 55,00 | 70,00 | 90,00 |
| Tosa | 50,00 | 65,00 | 85,00 | 110,00 |
| Banho + Tosa | 70,00 | 95,00 | 120,00 | 150,00 |
| Tosa Higiênica | 30,00 | 40,00 | 50,00 | 65,00 |
| Hidratação | 25,00 | 35,00 | 45,00 | 60,00 |

Valores configuráveis em `src/tabela_precos.py`.

---

## Contexto acadêmico

Projeto desenvolvido para a disciplina de **Atividade de Extensão Universitária** no semestre **2026.1**, alinhado ao **Objetivo de Desenvolvimento Sustentável (ODS) 8 — Trabalho Decente e Crescimento Econômico** e ao **Programa de Inovação e Empreendedorismo**.

**Período da atividade:** 08/02/2026 a 22/05/2026
**Beneficiário:** Momento Pet — Suzano/SP
**Proprietário:** Sr. Fernando

A metodologia seguiu as quatro fases do **Design Thinking**: Imersão, Ideação, Prototipação e Implantação/Revisão. O ferramental foi entregue ao microempreendedor com tutorial passo a passo e acompanhamento durante o primeiro mês de uso.

---

## Autor

**Gabriel Cunha**
Estudante de Graduação - 2026.1

---

## Licença

Uso acadêmico e livre para o microempreendedor parceiro.
