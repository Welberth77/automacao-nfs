# verificar_cnpjs.py
import os
from certidoes import extrair_cnpj_do_pfx

pasta = "./dados/sao_paulo/certidoes"
senhas = {
    "PAULO HENRIQUE WODEWOTZKY LTDA_60720594000137.pfx": "Paulo2025"
}

for arquivo, senha in senhas.items():
    caminho = os.path.join(pasta, arquivo)
    try:
        cnpj = extrair_cnpj_do_pfx(caminho, senha)
        print(f"✅ {arquivo} → CNPJ extraído: {cnpj}")
    except Exception as e:
        print(f"❌ Erro: {e}")