# cidades/sao_paulo.py
import time
import os
from datetime import datetime
from playwright.sync_api import sync_playwright
import sys

if getattr(sys, 'frozen', False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(sys._MEIPASS, "ms-playwright")

def aguardar(page, segundos=2):
    page.wait_for_load_state("networkidle")
    time.sleep(segundos)

def emitir_nfse(cert: dict, nota: dict, linha_idx: int, planilha: str, headless: bool = False):
    cnpj_tomador = str(nota.get('CNPJ Tomador', '')).strip()
    if not cnpj_tomador or cnpj_tomador == 'nan':
        print(f"❌ CNPJ Tomador vazio, pulando...")
        return

    print(f"\n🚀 Emitindo nota → CNPJ Tomador: {cnpj_tomador}")

    with open(cert['path'], 'rb') as f:
        pfx_bytes = f.read()

    url = str(nota.get('URL do Sistema', '')).strip()
    if not url or url == 'nan':
        url = "https://nfe.prefeitura.sp.gov.br/login.aspx"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            ignore_https_errors=True,
            client_certificates=[
                {
                    "origin": "https://nfe.prefeitura.sp.gov.br",
                    "pfx": pfx_bytes,
                    "passphrase": cert['password']
                }
            ]
        )

        page = context.new_page()

        try:
            # Passo 1 — Acessa o site
            print("🌐 Acessando NFS-e SP...")
            page.goto(url)
            aguardar(page, 2)

            # Passo 2 — Clica em Login único
            print("🖱️ Clicando em Login único...")
            page.click("button.oauth-button")
            aguardar(page, 2)
            print("✅ Clicou em Login único!")

            # Passo 3 — Via Certificado ICP-Brasil
            print("🖱️ Clicando em Via Certificado ICP-Brasil...")
            page.click("#btnCertificado")
            aguardar(page, 3)
            print("✅ Certificado enviado automaticamente!")

            # Passo 4 — Acessar o Sistema
            print("🖱️ Clicando em Acessar o Sistema...")
            page.wait_for_selector("#ctl00_body_btAcesso", timeout=10000)
            page.click("#ctl00_body_btAcesso")
            aguardar(page, 3)
            print("✅ Acessou o sistema!")

            # Passo 5 — Emissão de NFS-e
            print("🖱️ Clicando em Emissão de NFS-e...")
            page.wait_for_selector("a[href='nota.aspx']", timeout=10000)
            page.click("a[href='nota.aspx']")
            aguardar(page, 2)
            print("✅ Na página de emissão de NFS-e!")

            # Passo 6 — Preenche CNPJ do tomador da planilha
            print(f"✏️ Preenchendo CNPJ do tomador: {cnpj_tomador}...")
            page.wait_for_selector("#ctl00_body_tbCPFCNPJTomador", timeout=10000)

            cnpj_limpo = ''.join(filter(str.isdigit, cnpj_tomador))
            if len(cnpj_limpo) == 14:
                cnpj_formatado = f"{cnpj_limpo[:2]}.{cnpj_limpo[2:5]}.{cnpj_limpo[5:8]}/{cnpj_limpo[8:12]}-{cnpj_limpo[12:]}"
            elif len(cnpj_limpo) == 11:
                cnpj_formatado = f"{cnpj_limpo[:3]}.{cnpj_limpo[3:6]}.{cnpj_limpo[6:9]}-{cnpj_limpo[9:]}"
            else:
                cnpj_formatado = cnpj_tomador

            page.fill("#ctl00_body_tbCPFCNPJTomador", cnpj_formatado)
            aguardar(page, 1)
            page.click("#ctl00_body_btAvancar")
            aguardar(page, 3)
            print("✅ CNPJ do tomador preenchido!")

            # Passo 7 — Lê o código pré-selecionado no dropdown
            print("🔍 Lendo código pré-selecionado no dropdown...")
            frame = page
            for f in page.frames:
                if f.query_selector("#ctl00_body_ddlAtividade"):
                    frame = f
                    print(f"  → Frame encontrado: {f.url}")
                    break

            frame.wait_for_selector("#ctl00_body_ddlAtividade", timeout=10000)
            codigo_selecionado = frame.eval_on_selector(
                "#ctl00_body_ddlAtividade",
                "select => select.options[select.selectedIndex].value"
            )
            print(f"✅ Código pré-selecionado: {codigo_selecionado}")

            # Passo 8 — Digita o código e clica em >>
            print(f"✏️ Digitando código do serviço: {codigo_selecionado}...")
            frame.wait_for_selector("#ctl00_body_tbServEncerradoCodigo", timeout=10000)
            frame.fill("#ctl00_body_tbServEncerradoCodigo", codigo_selecionado)
            aguardar(frame, 1)
            frame.click("#ctl00_body_btnServEncerradoSelecionar")
            aguardar(frame, 3)
            print("✅ Código do serviço selecionado!")

            # Passo 9 — Preenche descrição do serviço
            print("✏️ Preenchendo descrição do serviço...")
            frame.wait_for_selector("#ctl00_body_tbDiscriminacao", timeout=10000)
            frame.fill("#ctl00_body_tbDiscriminacao", "1 Serviço Prestado: ")
            aguardar(frame, 1)
            print("✅ Descrição preenchida!")

            # Passo 10 — Preenche o valor do serviço
            valor = str(nota.get('Valor', '')).strip()
            print(f"✏️ Preenchendo valor: {valor}...")
            frame.wait_for_selector("#ctl00_body_tbValor", timeout=10000)
            frame.fill("#ctl00_body_tbValor", valor)
            aguardar(frame, 1)
            print("✅ Valor preenchido!")

            # Passo 11 — Clica em EMITIR
            print("🖱️ Clicando em EMITIR...")
            frame.wait_for_selector("#ctl00_body_btEmitir", timeout=10000)
            page.on("dialog", lambda dialog: dialog.accept())
            frame.click("#ctl00_body_btEmitir")
            aguardar(frame, 5)
            print("✅ Nota fiscal emitida!")

            # Passo 12 — Baixa o PDF da NFS-e emitida
            print("🖱️ Baixando PDF da NFS-e...")
            try:
                frame.wait_for_selector("#btDownload", timeout=10000)

                pasta_notas = os.path.join(os.path.dirname(planilha), "notas_emitidas")
                os.makedirs(pasta_notas, exist_ok=True)

                nome_empresa = str(nota.get('Nome da Empresa', 'EMPRESA')).strip()
                nome_empresa = "".join(c for c in nome_empresa if c.isalnum() or c in " _-").strip()
                nome_empresa = nome_empresa.replace(" ", "_").upper()

                data_emissao = datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
                nome_arquivo = f"NFS-e_{nome_empresa}_{data_emissao}.pdf"
                caminho_pdf = os.path.join(pasta_notas, nome_arquivo)

                with page.expect_download() as download_info:
                    frame.click("#btDownload")

                download = download_info.value
                download.save_as(caminho_pdf)
                print(f"✅ PDF salvo: {caminho_pdf}")

            except Exception as e:
                print(f"⚠️ Não foi possível baixar o PDF: {e}")

            from main import atualizar_status
            atualizar_status(planilha, linha_idx, "OK")
            print("✅ Status atualizado na planilha!")

        except Exception as e:
            print(f"❌ Erro: {e}")
            from main import atualizar_status
            atualizar_status(planilha, linha_idx, "ERRO", observacao=str(e))

        finally:
            context.close()
            browser.close()