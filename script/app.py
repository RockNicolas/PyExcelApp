import pandas as pd
import os
import time
import threading
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk, messagebox

load_dotenv()

arquivo = os.getenv('ARQUIVO_ENTRADA')
arquivo_saida = os.getenv('ARQUIVO_SAIDA')

col_concorrentes = ['DOG CHOW', 'PEDIGREE', 'CHAMP', 'KITEKAT', 'PURINA FRISKIES', 'WHISKAS']

def executar_analise(progress_bar, janela):

    for i in range(101):
        progress_bar["value"] = i
        janela.update_idletasks()
        time.sleep(0.05)

    df = pd.read_excel(arquivo)

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

    messagebox.showinfo("Finalizado", f"Análise gerada com sucesso em '{arquivo_saida}'")
    janela.destroy()

def iniciar_interface():
    janela = tk.Tk()
    janela.title("Analisando concorrência...")
    janela.geometry("400x100")

    label = tk.Label(janela, text="Gerando arquivo de comparação, aguarde...")
    label.pack(pady=10)

    progress_bar = ttk.Progressbar(janela, orient="horizontal", length=300, mode="determinate", maximum=100)
    progress_bar.pack(pady=5)

    thread = threading.Thread(target=executar_analise, args=(progress_bar, janela))
    thread.start()

    janela.mainloop()

iniciar_interface()
