# certificados.py
import json
from cryptography.hazmat.primitives.serialization.pkcs12 import load_key_and_certificates
from cryptography.hazmat.backends import default_backend

def extrair_cnpj_do_pfx(pfx_path: str, password: str) -> str:
    with open(pfx_path, 'rb') as f:
        pfx_data = f.read()

    _, cert, _ = load_key_and_certificates(
        pfx_data,
        password.encode('utf-8'),
        default_backend()
    )

    for attr in cert.subject:
        numeros = ''.join(filter(str.isdigit, attr.value))
        if len(numeros) >= 14:
            return numeros[-14:]

    raise Exception(f"CNPJ não encontrado no certificado: {pfx_path}")


def encontrar_certificado(cnpj_alvo: str) -> dict:
    cnpj_limpo = ''.join(filter(str.isdigit, cnpj_alvo))

    with open('./config.json', 'r') as f:
        config = json.load(f)

    for cert in config['certificados']:
        print(f"\n🔍 Verificando: {cert['path']}")
        try:
            cnpj_cert = extrair_cnpj_do_pfx(cert['path'], cert['password'])
            print(f"  CNPJ encontrado: {cnpj_cert}")

            if cnpj_cert == cnpj_limpo:
                print(f"  ✅ Certificado correto!")
                return cert

        except Exception as e:
            print(f"  ⚠️ Erro: {e}")
            continue

    raise Exception(f"❌ Nenhum certificado encontrado para o CNPJ: {cnpj_alvo}")