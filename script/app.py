import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

arquivo = os.getenv('ARQUIVO_ENTRADA')
arquivo_saida = os.getenv('ARQUIVO_SAIDA')

df = pd.read_excel(arquivo)

col_concorrentes = ['DOG CHOW', 'PEDIGREE', 'CHAMP', 'KITEKAT', 'PURINA FRISKIES', 'WHISKAS']

def analisar_concorrencia(row):
    precos = {}
    for col in col_concorrentes:
        val = str(row[col])
        if 'R$' in val:
            try:
                preco = float(val.replace('R$', '').replace(',', '.').split()[0])
                precos[col] = preco
            except:
                continue
    if precos:
        marca_mais_barata = min(precos, key=precos.get)
        preco_mais_barato = precos[marca_mais_barata]
        diff_valor = preco_mais_barato - row['Preço de Venda']
        diff_percentual = (diff_valor / row['Preço de Venda']) * 100
        return pd.Series([marca_mais_barata, preco_mais_barato, diff_valor, diff_percentual])
    else:
        return pd.Series([None, None, None, None])

df[['Concorrente Mais Barato', 'Preço Concorrente', 'Diferença R$', 'Diferença %']] = df.apply(analisar_concorrencia, axis=1)

colunas_finais = ['Nome do Produto', 'Preço de Venda', 'Concorrente Mais Barato', 'Preço Concorrente', 'Diferença R$', 'Diferença %']
resultado = df[colunas_finais]

resultado.to_excel(arquivo_saida, index=False)
print(f"Análise gerada com sucesso em '{arquivo_saida}'")
