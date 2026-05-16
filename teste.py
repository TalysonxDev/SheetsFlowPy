import os
import pandas as pd

arquivo = r"C:\Users\Talyson Negrao\OneDrive\Documentos\Sitran\Projeto Automação Fluxo\FluxoDiario.xlsx"


if os.path.exists(arquivo):
    print("Arquivo encontrado! Abrindo...")
    df = pd.read_excel(arquivo)
    print(df.head())  # mostra as primeiras linhas
else:
    print("Arquivo NÃO encontrado. Verifique o caminho:")
    print(arquivo)