# certidoes.py
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