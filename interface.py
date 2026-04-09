# interface.py
import customtkinter as ctk
import threading
import os
import sys
import importlib
import pandas as pd
from datetime import datetime

class RedirectOutput:
    def __init__(self, callback):
        self.callback = callback

    def write(self, msg):
        if msg.strip():
            self.callback(msg.strip())

    def flush(self):
        pass


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()
PASTA_DADOS = os.path.join(BASE_PATH, "dados")

SISTEMAS = {
    "sao_paulo":      "cidades.sao_paulo",
    "campinas":       "cidades.betha",
    "rio_de_janeiro": "cidades.rio_de_janeiro",
    "belo_horizonte": "cidades.betha",
}

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação NFS-e")
        self.geometry("860x580")
        self.resizable(False, False)
        self.parar = False
        self.rodando = False
        self.headless = ctk.BooleanVar(value=False)

        self._build_ui()
        self._carregar_cidades()

        sys.stdout = RedirectOutput(self._log)
        sys.stderr = RedirectOutput(self._log)

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Painel esquerdo ──
        self.painel_esq = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.painel_esq.grid(row=0, column=0, sticky="nsew")
        self.painel_esq.grid_propagate(False)
        self.painel_esq.grid_rowconfigure(3, weight=1)
        self.painel_esq.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.painel_esq,
            text="NFS-e",
            font=ctk.CTkFont(size=22, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(25, 2))

        ctk.CTkLabel(
            self.painel_esq,
            text="Automação de emissão",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        ).grid(row=1, column=0, sticky="w", padx=20)

        ctk.CTkLabel(
            self.painel_esq,
            text="CIDADES",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="gray"
        ).grid(row=2, column=0, sticky="w", padx=20, pady=(25, 8))

        self.frame_cidades = ctk.CTkScrollableFrame(
            self.painel_esq,
            height=120,
            fg_color="transparent"
        )
        self.frame_cidades.grid(row=3, column=0, sticky="new", padx=12)

        self.var_todas = ctk.BooleanVar()
        ctk.CTkCheckBox(
            self.painel_esq,
            text="Selecionar todas",
            variable=self.var_todas,
            command=self._toggle_todas,
            font=ctk.CTkFont(size=12)
        ).grid(row=4, column=0, sticky="w", padx=20, pady=(12, 4))

        ctk.CTkCheckBox(
            self.painel_esq,
            text="Ocultar browser",
            variable=self.headless,
            font=ctk.CTkFont(size=12)
        ).grid(row=5, column=0, sticky="w", padx=20, pady=(4, 12))

        ctk.CTkFrame(
            self.painel_esq, height=1, fg_color="gray"
        ).grid(row=6, column=0, sticky="ew", padx=16, pady=(0, 12))

        self.btn_iniciar = ctk.CTkButton(
            self.painel_esq,
            text="Iniciar emissão",
            command=self._iniciar,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8
        )
        self.btn_iniciar.grid(row=7, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.btn_parar = ctk.CTkButton(
            self.painel_esq,
            text="Parar",
            command=self._parar,
            height=40,
            font=ctk.CTkFont(size=13),
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray80"),
            corner_radius=8,
            state="disabled"
        )
        self.btn_parar.grid(row=8, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.btn_atualizar = ctk.CTkButton(
            self.painel_esq,
            text="↻  Atualizar",
            command=self._atualizar,
            height=60,
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8
        )
        self.btn_atualizar.grid(row=9, column=0, sticky="ew", padx=16, pady=(0, 20))

        # ── Painel direito ──
        self.painel_dir = ctk.CTkFrame(self, fg_color="transparent")
        self.painel_dir.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.painel_dir.grid_columnconfigure((0, 1, 2), weight=1)
        self.painel_dir.grid_rowconfigure(1, weight=1)

        self.card_pendente = self._card_metrica(self.painel_dir, "Pendentes", "0", 0)
        self.card_ok       = self._card_metrica(self.painel_dir, "Emitidas",  "0", 1, cor="#2CC985")
        self.card_erro     = self._card_metrica(self.painel_dir, "Erros",     "0", 2, cor="#E84040")

        frame_log = ctk.CTkFrame(self.painel_dir)
        frame_log.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        frame_log.grid_rowconfigure(1, weight=1)
        frame_log.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame_log,
            text="Progresso",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(12, 6))

        self.log = ctk.CTkTextbox(
            frame_log,
            font=ctk.CTkFont(family="Courier New", size=12),
            wrap="word",
            state="disabled"
        )
        self.log.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        self.status_var = ctk.StringVar(value="Aguardando...")
        ctk.CTkLabel(
            self,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="gray",
            anchor="w"
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 8))

    def _card_metrica(self, parent, label, valor, col, cor=None):
        frame = ctk.CTkFrame(parent)
        frame.grid(row=0, column=col, sticky="ew", padx=(0, 8) if col < 2 else 0)

        ctk.CTkLabel(
            frame,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(anchor="w", padx=14, pady=(12, 2))

        lbl = ctk.CTkLabel(
            frame,
            text=valor,
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=cor if cor else ("gray10", "gray90")
        )
        lbl.pack(anchor="w", padx=14, pady=(0, 12))

        return lbl

    def _carregar_cidades(self):
        self.checks = {}
        self.total_pendentes = 0

        if not os.path.exists(PASTA_DADOS):
            self._log("⚠ Pasta /dados não encontrada!")
            return

        for cidade in sorted(os.listdir(PASTA_DADOS)):
            caminho = os.path.join(PASTA_DADOS, cidade)
            if not os.path.isdir(caminho) or cidade not in SISTEMAS:
                continue

            planilha = os.path.join(caminho, f"{cidade}.xlsx")
            pendentes = 0
            if os.path.exists(planilha):
                try:
                    df = pd.read_excel(planilha, dtype=str)
                    df.columns = df.columns.str.strip()
                    df = df[df["CNPJ Prestador"].notna()]
                    df = df[df["CNPJ Prestador"] != "12.345.678/0001-99"]
                    if "Status" in df.columns:
                        pendentes = len(df[df["Status"].str.upper() == "PENDENTE"])
                    else:
                        pendentes = len(df)
                except:
                    pass

            self.total_pendentes += pendentes

            var = ctk.BooleanVar()
            label = f"{cidade.replace('_', ' ').title()}  ({pendentes})"
            cb = ctk.CTkCheckBox(
                self.frame_cidades,
                text=label,
                variable=var,
                font=ctk.CTkFont(size=12)
            )
            cb.pack(anchor="w", pady=4)
            self.checks[cidade] = var

        self.card_pendente.configure(text=str(self.total_pendentes))

    def _atualizar(self):
        if self.rodando:
            self._log("⚠ Não é possível atualizar durante a emissão!")
            return

        for widget in self.frame_cidades.winfo_children():
            widget.destroy()

        self.card_pendente.configure(text="0")
        self.card_ok.configure(text="0")
        self.card_erro.configure(text="0")
        self.var_todas.set(False)

        self._carregar_cidades()
        self._log("↻ Interface atualizada!")
        self._set_status("Atualizado!")

    def _toggle_todas(self):
        for var in self.checks.values():
            var.set(self.var_todas.get())

    def _log(self, msg: str):
        def _write():
            self.log.configure(state="normal")
            hora = datetime.now().strftime("%H:%M:%S")
            self.log.insert("end", f"[{hora}] {msg}\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _write)

    def _set_status(self, texto):
        self.after(0, lambda: self.status_var.set(texto))

    def _iniciar(self):
        cidades = [c for c, v in self.checks.items() if v.get()]
        if not cidades:
            self._log("⚠ Selecione pelo menos uma cidade!")
            return

        self.parar = False
        self.rodando = True
        self.total_ok = 0
        self.total_erro = 0

        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.btn_atualizar.configure(state="disabled")
        self.card_ok.configure(text="0")
        self.card_erro.configure(text="0")

        thread = threading.Thread(
            target=self._rodar,
            args=(cidades, self.headless.get()),
            daemon=True
        )
        thread.start()

    def _parar(self):
        self.parar = True
        self._log("⏹ Parando após a nota atual...")
        self.btn_parar.configure(state="disabled")

    def _rodar(self, cidades, headless: bool):
        from main import (
            validar_planilha,
            atualizar_status,
            carregar_certificados_da_cidade,
        )

        for cidade in cidades:
            if self.parar:
                break

            self._log(f"\n── {cidade.upper()} ──")
            self._set_status(f"Processando {cidade}...")

            planilha = os.path.join(PASTA_DADOS, cidade, f"{cidade}.xlsx")
            if not os.path.exists(planilha):
                self._log(f"⚠ Planilha não encontrada: {planilha}")
                continue

            try:
                df = pd.read_excel(planilha, dtype=str)
                df.columns = df.columns.str.strip()
                df = df[df["CNPJ Prestador"].notna()]
                df = df[df["CNPJ Prestador"] != "12.345.678/0001-99"]

                if df.empty:
                    self._log(f"⚠ Nenhuma nota em {cidade}")
                    continue

                if not validar_planilha(df):
                    self._log(f"✗ Planilha inválida em {cidade}, pulando...")
                    continue

                certificados = carregar_certificados_da_cidade(cidade, df)
                if not certificados:
                    self._log(f"✗ Nenhuma certidão carregada para {cidade}")
                    continue

                if "Status" in df.columns:
                    df_pendentes = df[df["Status"].str.upper() == "PENDENTE"].copy()
                else:
                    df_pendentes = df.copy()

                self._log(f"✓ {len(df_pendentes)} nota(s) pendente(s)")

                modulo = importlib.import_module(SISTEMAS[cidade])

                for original_idx, nota in df_pendentes.iterrows():
                    if self.parar:
                        break

                    cnpj_prestador = ''.join(filter(str.isdigit, nota['CNPJ Prestador']))
                    cert = certificados.get(cnpj_prestador)

                    if not cert:
                        self._log(f"⚠ Certidão não encontrada: {cnpj_prestador}")
                        atualizar_status(planilha, original_idx, "ERRO", observacao="Certidão não encontrada")
                        self.total_erro += 1
                        self.after(0, lambda: self.card_erro.configure(text=str(self.total_erro)))
                        continue

                    cnpj_tomador = str(nota.get('CNPJ Tomador', '')).strip()
                    self._log(f"→ Emitindo para CNPJ: {cnpj_tomador}")
                    self._set_status(f"Emitindo nota {original_idx + 1}...")

                    try:
                        modulo.emitir_nfse(cert, nota.to_dict(), original_idx, planilha, headless)
                        self._log(f"✓ Nota emitida com sucesso!")
                        self.total_ok += 1
                        self.after(0, lambda: self.card_ok.configure(text=str(self.total_ok)))
                    except Exception as e:
                        self._log(f"✗ Erro: {e}")
                        atualizar_status(planilha, original_idx, "ERRO", observacao=str(e))
                        self.total_erro += 1
                        self.after(0, lambda: self.card_erro.configure(text=str(self.total_erro)))

            except Exception as e:
                self._log(f"✗ Erro inesperado: {e}")

        self._log(f"\n✓ Finalizado — {self.total_ok} emitida(s), {self.total_erro} erro(s)")
        self._set_status(f"Finalizado — {self.total_ok} ok, {self.total_erro} erro(s)")
        self.after(0, lambda: self.btn_iniciar.configure(state="normal"))
        self.after(0, lambda: self.btn_parar.configure(state="disabled"))
        self.after(0, lambda: self.btn_atualizar.configure(state="normal"))
        self.rodando = False

    def on_closing(self):
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()