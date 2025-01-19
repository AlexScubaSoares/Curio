#import os

#diretorio = "/OLGA"
#arquivos = os.listdir(diretorio)

#print(arquivos)

import os
from os import chdir, getcwd, listdir
import pathlib
import time


caminho = input('Digite o caminho')
validos = 0
lidos = 0

chdir(caminho)
print(getcwd())

for c in listdir():
    if os.path.isfile(c) and c.lower().endswith(".docx") and c.lower().startswith('material_'):
        arq = pathlib.Path(c)
        docri = os.path.getmtime(arq)
        criacao = time.ctime(docri)
        dtformated = time.strptime(criacao)
        dt_final = time.strftime("%Y-%m-%d", dtformated)
        lidos += 1
        if len(c) > 32:
            print(c,';', dt_final,';', len(c),';', getcwd())
            validos += 1

print(f"Total de documentos Word => Lidos: {lidos} - Maiores que 32: {validos}")
