# main.py
import os
import importlib
import pandas as pd
from certificados import extrair_cnpj_do_pfx
import json

PASTA_NOTAS = "./notas"

# Mapeamento: cidade → qual sistema usar
SISTEMAS = {
    "sao_paulo":      "cidades.sao_paulo",
    "campinas":       "cidades.betha",
    "rio_de_janeiro": "cidades.rio_de_janeiro",
    "belo_horizonte": "cidades.betha",
    # adicione novas cidades aqui...
}

def carregar_certificados():
    with open('./config.json', 'r') as f:
        config = json.load(f)

    certificados = {}
    for cert in config['certificados']:
        try:
            cnpj = extrair_cnpj_do_pfx(cert['path'], cert['password'])
            certificados[cnpj] = cert
            print(f"✅ Carregado: {cert['path']} → CNPJ: {cnpj}")
        except Exception as e:
            print(f"⚠️ Erro ao carregar {cert['path']}: {e}")

    return certificados

def main():
    print("🔍 Carregando certificados...\n")
    certificados = carregar_certificados()

    # Varre todas as planilhas da pasta /notas
    for arquivo in os.listdir(PASTA_NOTAS):
        if not arquivo.endswith(".xlsx"):
            continue

        cidade = arquivo.replace(".xlsx", "")
        caminho = os.path.join(PASTA_NOTAS, arquivo)

        if cidade not in SISTEMAS:
            print(f"⚠️ Cidade '{cidade}' não tem sistema mapeado, pulando...")
            continue

        print(f"\n🏙️ Processando cidade: {cidade.upper()}")

        # Importa o módulo da cidade dinamicamente
        modulo = importlib.import_module(SISTEMAS[cidade])

        # Lê as notas pendentes da planilha
        df = pd.read_excel(caminho, sheet_name="Notas", dtype=str)
        df.columns = df.columns.str.strip()
        pendentes = df[df["STATUS"].str.upper() == "PENDENTE"]

        print(f"📋 {len(pendentes)} nota(s) pendente(s) em {cidade}")

        for idx, (_, nota) in enumerate(pendentes.iterrows()):
            cnpj_prestador = ''.join(filter(str.isdigit, nota['CNPJ PRESTADOR']))
            cert = certificados.get(cnpj_prestador)

            if not cert:
                print(f"⚠️ Certificado não encontrado para CNPJ: {cnpj_prestador}, pulando...")
                continue

            # Chama a função de emissão do sistema correto
            modulo.emitir_nfse(cert, nota.to_dict(), idx, caminho)

    print("\n🎉 Processo finalizado para todas as cidades!")

if __name__ == "__main__":
    main()