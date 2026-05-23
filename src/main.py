# main.py - Ponto de entrada do programa
# Orquestra a leitura, validacao e geracao de relatorio.

import argparse
import os
import sys
from datetime import datetime

# Garante que o Python encontre os modulos dentro de src/
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from validador import ler_planilha, validar_dados
from gerar_relatorio import gerar_relatorio


def main():
    # Configuracao dos argumentos de linha de comando
    parser = argparse.ArgumentParser(
        description='Validador de Comandas de Banho e Tosa'
    )
    parser.add_argument(
        '--entrada',
        default='dados/comandas_input.xlsx',
        help='Caminho do arquivo Excel de entrada',
    )
    parser.add_argument(
        '--saida',
        default=None,
        help='Caminho do arquivo Excel de saida (opcional)',
    )
    args = parser.parse_args()

    # Define caminho de saida padrao (com data e hora)
    if args.saida is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        args.saida = f'dados/relatorio_validacao_{timestamp}.xlsx'

    print('=' * 50)
    print('INICIANDO VALIDACAO DE COMANDAS DO PET SHOP')
    print('=' * 50)

    # Passo 1: Ler a planilha
    df = ler_planilha(args.entrada)
    if df is None:
        print('Encerrando: nao foi possivel ler a planilha.')
        return

    # Passo 2: Validar os dados
    problemas = validar_dados(df)

    # Passo 3: Gerar o relatorio Excel
    gerar_relatorio(df, problemas, args.saida)

    print('=' * 50)
    print('VALIDACAO CONCLUIDA!')
    print(f'Total de atendimentos lidos: {len(df)}')
    print(f'Problemas encontrados: {len(problemas)}')
    print(f'Relatorio salvo em: {args.saida}')
    print('=' * 50)


if __name__ == '__main__':
    main()