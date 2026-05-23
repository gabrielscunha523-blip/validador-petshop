# validador.py - Le a planilha e aplica as regras de validacao

import pandas as pd
import re
from datetime import datetime, timedelta

# Importa a tabela de precos e os dominios validos
from tabela_precos import (
    TABELA_PRECOS,
    TOLERANCIA_PRECO,
    SERVICOS_VALIDOS,
    PORTES_VALIDOS,
    STATUS_VALIDOS,
    FORMAS_PAGAMENTO_VALIDAS,
)


def ler_planilha(caminho):
    """Le a planilha Excel e retorna um DataFrame.
    Se der erro, retorna None.
    """
    print(f"Lendo a planilha: {caminho}")
    try:
        df = pd.read_excel(caminho)
        print(f"Planilha lida com sucesso! {len(df)} linhas encontradas.")
        return df
    except FileNotFoundError:
        print(f"ERRO: Arquivo '{caminho}' nao encontrado.")
        return None
    except Exception as e:
        print(f"ERRO ao ler planilha: {e}")
        return None


def _valor_vazio(valor):
    """Devolve True se o valor for vazio, NaN ou so espacos."""
    if pd.isna(valor):
        return True
    if isinstance(valor, str) and valor.strip() == '':
        return True
    return False


def _telefone_valido(telefone):
    """Verifica se o telefone esta no formato (xx) xxxxx-xxxx."""
    if _valor_vazio(telefone):
        return False
    padrao = r'^\(\d{2}\)\s?\d{4,5}-\d{4}$'
    return bool(re.match(padrao, str(telefone).strip()))


def _converter_data(valor):
    """Tenta converter um valor para data. Retorna None se nao der."""
    if _valor_vazio(valor):
        return None
    try:
        return pd.to_datetime(valor).date()
    except Exception:
        return None


def validar_linha(linha, numero_linha):
    """Aplica todas as regras a uma linha. Retorna lista de problemas."""
    problemas = []

    # REGRA 1: Data de atendimento deve existir e ser valida
    data_atendimento = _converter_data(linha.get('Data_Atendimento'))
    if data_atendimento is None:
        problemas.append('Data_Atendimento invalida ou vazia')
    else:
        # Atendimento nao pode estar mais de 1 ano no futuro
        limite_futuro = datetime.now().date() + timedelta(days=365)
        if data_atendimento > limite_futuro:
            problemas.append('Data_Atendimento muito no futuro')

    # REGRA 2: Nome do tutor preenchido
    if _valor_vazio(linha.get('Nome_Tutor')):
        problemas.append('Nome_Tutor em branco')

    # REGRA 3: Telefone do tutor em formato valido
    if not _telefone_valido(linha.get('Telefone_Tutor')):
        problemas.append('Telefone_Tutor em formato invalido')

    # REGRA 4: Nome do pet preenchido
    if _valor_vazio(linha.get('Nome_Pet')):
        problemas.append('Nome_Pet em branco')

    # REGRA 5: Porte dentro do dominio (P, M, G, GG)
    porte = str(linha.get('Porte', '')).strip()
    if porte not in PORTES_VALIDOS:
        problemas.append(f"Porte '{porte}' invalido (use P/M/G/GG)")

    # REGRA 6: Servico dentro da lista permitida
    servico = str(linha.get('Servico', '')).strip()
    if servico not in SERVICOS_VALIDOS:
        problemas.append(f"Servico '{servico}' nao esta na lista permitida")

    # REGRA 7: Valor > 0
    valor = linha.get('Valor_Servico')
    if pd.isna(valor) or valor <= 0:
        problemas.append(f"Valor_Servico invalido ({valor})")

    # REGRA 8: Valor compativel com a tabela de precos
    if servico in SERVICOS_VALIDOS and porte in PORTES_VALIDOS:
        if not pd.isna(valor) and valor > 0:
            preco_tabela = TABELA_PRECOS[servico][porte]
            diferenca = abs(valor - preco_tabela) / preco_tabela
            if diferenca > TOLERANCIA_PRECO:
                problemas.append(
                    f"Valor R${valor:.2f} diverge da tabela "
                    f"(esperado ~R${preco_tabela:.2f})"
                )

    # REGRA 9: Status_Pagamento dentro do dominio
    status = str(linha.get('Status_Pagamento', '')).strip()
    if status not in STATUS_VALIDOS:
        problemas.append(f"Status_Pagamento '{status}' invalido (use Pago/Pendente)")

    # REGRA 10: Forma de pagamento (opcional, mas se preenchida deve ser valida)
    forma = str(linha.get('Forma_Pagamento', '')).strip() if not _valor_vazio(linha.get('Forma_Pagamento')) else ''
    if forma and forma not in FORMAS_PAGAMENTO_VALIDAS:
        problemas.append(f"Forma_Pagamento '{forma}' invalida")

    # REGRA 11: Coerencia entre Status e Data_Pagamento
    data_pagamento = _converter_data(linha.get('Data_Pagamento'))
    if status == 'Pago' and data_pagamento is None:
        problemas.append('Status=Pago mas Data_Pagamento em branco')
    if status == 'Pendente' and data_pagamento is not None:
        problemas.append('Status=Pendente mas Data_Pagamento preenchida (inconsistente)')

    return problemas


def validar_dados(df):
    """Valida todas as linhas do DataFrame.
    Tambem detecta duplicidades. Retorna lista de problemas.
    """
    print("Iniciando validacao dos dados...")
    problemas_totais = []

    # Validacao linha a linha
    for index, linha in df.iterrows():
        numero_linha = index + 2  # +2 porque Excel comeca em 1 e tem cabecalho
        problemas_linha = validar_linha(linha, numero_linha)
        if problemas_linha:
            problemas_totais.append({
                'Linha_Excel': numero_linha,
                'Tutor': linha.get('Nome_Tutor', ''),
                'Pet': linha.get('Nome_Pet', ''),
                'Servico': linha.get('Servico', ''),
                'Data': linha.get('Data_Atendimento', ''),
                'Valor': linha.get('Valor_Servico', ''),
                'Problemas': ' | '.join(problemas_linha),
            })

    # Validacao de duplicidades:
    # mesmo tutor + mesmo pet + mesmo servico + mesma data = suspeito
    chaves_duplicadas = df.duplicated(
        subset=['Nome_Tutor', 'Nome_Pet', 'Servico', 'Data_Atendimento'],
        keep=False
    )
    if chaves_duplicadas.any():
        duplicadas = df[chaves_duplicadas]
        for index, linha in duplicadas.iterrows():
            numero_linha = index + 2
            ja_listado = any(p['Linha_Excel'] == numero_linha for p in problemas_totais)
            if ja_listado:
                # Apenda o aviso ao registro existente
                for p in problemas_totais:
                    if p['Linha_Excel'] == numero_linha:
                        p['Problemas'] += ' | Possivel duplicidade'
            else:
                problemas_totais.append({
                    'Linha_Excel': numero_linha,
                    'Tutor': linha.get('Nome_Tutor', ''),
                    'Pet': linha.get('Nome_Pet', ''),
                    'Servico': linha.get('Servico', ''),
                    'Data': linha.get('Data_Atendimento', ''),
                    'Valor': linha.get('Valor_Servico', ''),
                    'Problemas': 'Possivel duplicidade (mesmo tutor/pet/servico/data)',
                })

    print(f"Validacao concluida. {len(problemas_totais)} linhas com problemas.")
    return problemas_totais