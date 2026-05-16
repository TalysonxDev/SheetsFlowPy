import pandas as pd

arquivo = r"C:\Users\Talyson Negrao\OneDrive\Documentos\Sitran\Projeto Automação Fluxo\Fluxo.xlsx"

df = pd.read_excel(arquivo)
print("Arquivo lido com sucesso!")