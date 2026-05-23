# gerar_relatorio.py - Gera o relatorio em Excel com varias abas

import pandas as pd
from datetime import datetime


def _ids_problemas(problemas):
    """Retorna o conjunto de linhas que tem problemas."""
    return {p['Linha_Excel'] for p in problemas}


def gerar_relatorio(df, problemas, caminho_saida):
    """Gera um arquivo Excel com 5 abas:
    1. Resumo            - visao geral
    2. Problemas         - linhas com erros encontrados
    3. Dados_OK          - linhas validadas sem erro
    4. Contas_a_Receber  - inadimplencia agrupada por tutor
    5. Resumo_Financeiro - totais financeiros
    """
    # ABA 1: Resumo geral
    total_linhas = len(df)
    com_problema = len(problemas)
    sem_problema = total_linhas - com_problema
    taxa_erro = (com_problema / total_linhas * 100) if total_linhas > 0 else 0

    resumo = pd.DataFrame({
        'Metrica': [
            'Total de atendimentos',
            'Linhas com problemas',
            'Linhas OK',
            'Taxa de erro (%)',
            'Gerado em',
        ],
        'Valor': [
            total_linhas,
            com_problema,
            sem_problema,
            f'{taxa_erro:.1f}%',
            datetime.now().strftime('%d/%m/%Y %H:%M'),
        ],
    })

    # ABA 2: Problemas (ja vem pronto)
    df_problemas = pd.DataFrame(problemas) if problemas else pd.DataFrame(
        columns=['Linha_Excel', 'Tutor', 'Pet', 'Servico', 'Data', 'Valor', 'Problemas']
    )

    # ABA 3: Dados OK (linhas que nao apareceram nos problemas)
    ids_com_problema = _ids_problemas(problemas)
    linhas_ok = []
    for index, linha in df.iterrows():
        numero_linha = index + 2
        if numero_linha not in ids_com_problema:
            linhas_ok.append({
                'Linha_Excel': numero_linha,
                'Data': linha.get('Data_Atendimento', ''),
                'Tutor': linha.get('Nome_Tutor', ''),
                'Telefone': linha.get('Telefone_Tutor', ''),
                'Pet': linha.get('Nome_Pet', ''),
                'Porte': linha.get('Porte', ''),
                'Servico': linha.get('Servico', ''),
                'Valor': linha.get('Valor_Servico', ''),
                'Status': linha.get('Status_Pagamento', ''),
            })
    df_ok = pd.DataFrame(linhas_ok) if linhas_ok else pd.DataFrame(
        columns=['Linha_Excel', 'Data', 'Tutor', 'Telefone', 'Pet', 'Porte', 'Servico', 'Valor', 'Status']
    )

    # ABA 4: Contas a Receber (pendentes agrupados por tutor)
    pendentes = df[df['Status_Pagamento'].astype(str).str.strip() == 'Pendente'].copy()
    if len(pendentes) > 0:
        pendentes['Data_Atendimento'] = pd.to_datetime(
            pendentes['Data_Atendimento'], errors='coerce'
        )
        agrupado = pendentes.groupby('Nome_Tutor').agg(
            Telefone=('Telefone_Tutor', 'first'),
            Total_Devido=('Valor_Servico', 'sum'),
            Qtd_Atendimentos=('Valor_Servico', 'count'),
            Atendimento_Mais_Antigo=('Data_Atendimento', 'min'),
            Lista_Pets=('Nome_Pet', lambda s: ', '.join(sorted(set(s.astype(str))))),
        ).reset_index()
        # Calcula quantos dias a divida mais antiga tem
        hoje = pd.Timestamp.today().normalize()
        agrupado['Dias_em_Aberto'] = (hoje - agrupado['Atendimento_Mais_Antigo']).dt.days
        agrupado = agrupado.sort_values('Total_Devido', ascending=False)
        df_pendentes = agrupado[[
            'Nome_Tutor', 'Telefone', 'Total_Devido',
            'Qtd_Atendimentos', 'Atendimento_Mais_Antigo',
            'Dias_em_Aberto', 'Lista_Pets',
        ]]
    else:
        df_pendentes = pd.DataFrame(columns=[
            'Nome_Tutor', 'Telefone', 'Total_Devido',
            'Qtd_Atendimentos', 'Atendimento_Mais_Antigo',
            'Dias_em_Aberto', 'Lista_Pets',
        ])

    # ABA 5: Resumo financeiro
    pagos = df[df['Status_Pagamento'].astype(str).str.strip() == 'Pago']
    total_recebido = pagos['Valor_Servico'].sum() if len(pagos) > 0 else 0
    total_pendente = pendentes['Valor_Servico'].sum() if len(pendentes) > 0 else 0
    total_geral = total_recebido + total_pendente
    inadimplencia = (total_pendente / total_geral * 100) if total_geral > 0 else 0
    ticket_medio = df['Valor_Servico'].mean() if len(df) > 0 else 0

    resumo_financeiro = pd.DataFrame({
        'Metrica': [
            'Total recebido (R$)',
            'Total pendente (R$)',
            'Faturamento bruto (R$)',
            'Taxa de inadimplencia (%)',
            'Ticket medio (R$)',
            'Atendimentos pagos',
            'Atendimentos pendentes',
        ],
        'Valor': [
            f'R$ {total_recebido:.2f}',
            f'R$ {total_pendente:.2f}',
            f'R$ {total_geral:.2f}',
            f'{inadimplencia:.1f}%',
            f'R$ {ticket_medio:.2f}',
            len(pagos),
            len(pendentes),
        ],
    })

    # Escreve tudo no Excel
    with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
        df_problemas.to_excel(writer, sheet_name='Problemas', index=False)
        df_ok.to_excel(writer, sheet_name='Dados_OK', index=False)
        df_pendentes.to_excel(writer, sheet_name='Contas_a_Receber', index=False)
        resumo_financeiro.to_excel(writer, sheet_name='Resumo_Financeiro', index=False)

        # Formatacao simples: ajusta largura das colunas em cada aba
        workbook = writer.book
        for sheet_name in ['Resumo', 'Problemas', 'Dados_OK', 'Contas_a_Receber', 'Resumo_Financeiro']:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:Z', 22)

    print(f"Relatorio salvo em: {caminho_saida}")
    return caminho_saida