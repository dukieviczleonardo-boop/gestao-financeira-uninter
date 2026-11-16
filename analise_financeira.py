"""Projeto extensionista UNINTER — análise de planilha Excel com Pandas e Matplotlib.
Instalação: pip install pandas matplotlib openpyxl
Execução: python analise_financeira.py
Os dados iniciais da planilha são fictícios e servem somente para testes.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent
ARQUIVO = BASE / 'Gestao_Financeira_UNINTER.xlsx'
SAIDA = BASE / 'resultados'
SAIDA.mkdir(exist_ok=True)


def ler_aba(nome):
    return pd.read_excel(ARQUIVO, sheet_name=nome)


def executar():
    if not ARQUIVO.exists():
        raise FileNotFoundError(f'Coloque a planilha na mesma pasta do script: {ARQUIVO}')
    produtos = ler_aba('PRODUTOS').dropna(subset=['Código'])
    vendas = ler_aba('VENDAS')[['Data', 'Código produto', 'Quantidade']].dropna(subset=['Código produto', 'Quantidade'])
    despesas = ler_aba('DESPESAS').dropna(subset=['Valor (R$)'])
    for col in ['Custo unitário (R$)', 'Preço venda (R$)', 'Estoque inicial']:
        produtos[col] = pd.to_numeric(produtos[col], errors='coerce').fillna(0)
    vendas['Quantidade'] = pd.to_numeric(vendas['Quantidade'], errors='coerce').fillna(0)
    despesas['Valor (R$)'] = pd.to_numeric(despesas['Valor (R$)'], errors='coerce').fillna(0)
    if (vendas['Quantidade'] < 0).any() or (despesas['Valor (R$)'] < 0).any():
        raise ValueError('Não são aceitas quantidades ou despesas negativas neste modelo.')
    if produtos['Código'].duplicated().any():
        raise ValueError('Existem códigos duplicados na aba PRODUTOS.')
    dados = vendas.merge(produtos, left_on='Código produto', right_on='Código', how='left', validate='many_to_one')
    if dados['Produto'].isna().any():
        raise ValueError('Há vendas com códigos de produtos não cadastrados.')
    dados['Receita'] = dados['Quantidade'] * dados['Preço venda (R$)']
    dados['Custo'] = dados['Quantidade'] * dados['Custo unitário (R$)']
    receita = dados['Receita'].sum()
    custo = dados['Custo'].sum()
    despesa = despesas['Valor (R$)'].sum()
    lucro = receita - custo - despesa
    margem = lucro / receita * 100 if receita else 0
    vendidos = dados.groupby('Código produto')['Quantidade'].sum()
    estoque = produtos[['Código', 'Produto', 'Estoque inicial']].copy()
    estoque['Quantidade vendida'] = estoque['Código'].map(vendidos).fillna(0)
    estoque['Saldo estimado'] = estoque['Estoque inicial'] - estoque['Quantidade vendida']
    estoque.to_csv(SAIDA / 'estoque_estimado.csv', index=False, encoding='utf-8-sig')
    resumo = pd.DataFrame({'Indicador':['Receita total','Custo de produtos vendidos','Despesas operacionais','Lucro gerencial estimado','Margem (%)'], 'Valor':[receita,custo,despesa,lucro,margem]})
    resumo.to_csv(SAIDA / 'resumo_financeiro.csv', index=False, encoding='utf-8-sig')
    print('\nRESUMO FINANCEIRO — DADOS DE EXEMPLO OU DADOS INFORMADOS')
    print(resumo.to_string(index=False, float_format=lambda x: f'{x:.2f}'))
    print('\nESTOQUE ESTIMADO')
    print(estoque.to_string(index=False))
    if not dados.empty:
        por_produto = dados.groupby('Produto')['Receita'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(9, 5))
        por_produto.plot(kind='bar', ax=ax)
        ax.set_title('Receita por produto')
        ax.set_ylabel('Receita (R$)')
        ax.set_xlabel('Produto')
        fig.tight_layout()
        fig.savefig(SAIDA / 'receita_por_produto.png', dpi=150)
        plt.close(fig)
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(['Receita','Custo produtos','Despesas','Lucro estimado'],[receita,custo,despesa,lucro])
    ax.set_title('Resumo financeiro')
    ax.set_ylabel('R$')
    fig.tight_layout()
    fig.savefig(SAIDA / 'resumo_financeiro.png', dpi=150)
    plt.close(fig)
    print(f'\nArquivos gerados em: {SAIDA}')

if __name__ == '__main__':
    executar()
