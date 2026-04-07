# listar_certificados.py
import os
import json

pasta = './certificados'
certificados = []

for arquivo in os.listdir(pasta):
    if arquivo.endswith('.pfx'):
        certificados.append({
            "path": os.path.join(pasta, arquivo),
            "password": "",  # preencha a senha depois
            "cnpj": ""
        })

print(json.dumps({"certificados": certificados}, indent=2, ensure_ascii=False))