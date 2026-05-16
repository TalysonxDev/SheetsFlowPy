import os

arquivo = r"C:\Users\Talyson Negrao\OneDrive\Documentos\Sitran\Projeto Automação Fluxo\Fluxo.xlsx"

print("Existe?", os.path.exists(arquivo))
print("É arquivo?", os.path.isfile(arquivo))