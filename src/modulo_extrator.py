import os
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from modulo_Sap import verificar_sap_aberto
from utils_metrics import incrementar_metrica
from utils_os import obter_caminho_script

def obter_script_iq09():
    return obter_caminho_script("SCRIPT DA IQ09.vbs")

SCRIPT_IQ09 = obter_script_iq09()

class AbaExtrator:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando_iq09 = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            widget.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)

        self.btn_extrair.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_iq09.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")

        for btn in [self.btn_copiar_sem, self.btn_copiar_com, self.btn_limpar]:
            btn.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
            
        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_sem.configure(text_color=c["texto"])
        self.lbl_com.configure(text_color=c["texto"])
        self.container.configure(fg_color="transparent")
        self.frame_resultados.configure(fg_color="transparent")
        self.frame_botoes_resultado.configure(fg_color="transparent")
        self.frame_acoes.configure(fg_color="transparent")

    def extrair_seriais(self):
        texto = self.txt_entrada.get("1.0", tk.END)
        seriais = self._extrair_seriais(texto)

        if not seriais:
            messagebox.showwarning("Aviso", "Nenhum serial encontrado para extrair.")
            return

        sem_ponto_virgula = "\n".join(seriais)
        com_ponto_virgula = "\n".join(f"{serial};" for serial in seriais)

        self._set_texto(self.txt_sem_ponto_virgula, sem_ponto_virgula)
        self._set_texto(self.txt_com_ponto_virgula, com_ponto_virgula)
        self._status(f"{len(seriais)} seriais extraídos.")
        incrementar_metrica("seriais_extraidos", len(seriais))

    def copiar_sem_ponto_virgula(self):
        self._copiar_texto(self.txt_sem_ponto_virgula, "Lista sem ponto e vírgula copiada.")

    def copiar_com_ponto_virgula(self):
        self._copiar_texto(self.txt_com_ponto_virgula, "Lista com ponto e vírgula copiada.")

    def abrir_na_iq09(self):
        if self.executando_iq09:
            messagebox.showwarning("Aviso", "A IQ09 já está em execução.")
            return

        seriais = self._obter_lista_sem_ponto_virgula()
        if not seriais:
            self.extrair_seriais()
            seriais = self._obter_lista_sem_ponto_virgula()

        if not seriais:
            return

        script_iq09 = obter_script_iq09()
        if not os.path.exists(script_iq09):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_iq09}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para abrir a IQ09.\nAbra o SAP e tente novamente.")
            return

        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append("\r\n".join(seriais))
        self.parent_frame.update()

        self.executando_iq09 = True
        self.btn_iq09.configure(state="disabled")
        self._status(f"Abrindo IQ09 com {len(seriais)} seriais sem ponto e vírgula...")

        thread = threading.Thread(target=self._executar_iq09, daemon=True)
        thread.start()

    def limpar(self):
        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            self._set_texto(widget, "")
        self._status("Extrator limpo.")

    def _executar_iq09(self):
        script_temp = None
        try:
            script_temp = self._criar_script_iq09_temporario()
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

            resultado = subprocess.run(
                ["cscript.exe", "//nologo", script_temp],
                capture_output=True, text=True, timeout=120, creationflags=creationflags,
            )

            if resultado.returncode == 0:
                self._status_threadsafe("IQ09 aberta com a lista sem ponto e vírgula.")
            else:
                erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", erro))
                self._status_threadsafe("Falha ao abrir IQ09.")
        except subprocess.TimeoutExpired:
            self.parent_frame.after(0, lambda: messagebox.showwarning("Aviso", "Tempo limite atingido ao abrir a IQ09."))
            self._status_threadsafe("IQ09 excedeu o tempo limite.")
        except Exception as e:
            self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", str(e)))
            self._status_threadsafe("Erro ao executar script da IQ09.")
        finally:
            if script_temp:
                try: os.remove(script_temp)
                except OSError: pass
            self.parent_frame.after(0, self._finalizar_iq09)

    def _criar_script_iq09_temporario(self):
        script_iq09 = obter_script_iq09()
        with open(script_iq09, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
        conteudo = self._inserir_reset_sap(conteudo)
        fd, caminho = tempfile.mkstemp(prefix="iq09_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="mbcs") as f:
            f.write(conteudo)
        return caminho

    def _inserir_reset_sap(self, conteudo):
        if "Reset automatico inserido pelo Conecta Hub" in conteudo: return conteudo
        ponto_ancora = 'session.findById("wnd[0]").maximize'
        bloco_reset = (
            "' Reset automatico inserido pelo Conecta Hub\n"
            'On Error Resume Next\n'
            'session.findById("wnd[0]/tbar[0]/btn[12]").press\n'
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nS000"\n'
            'session.findById("wnd[0]").sendVKey 0\n'
            'WScript.Sleep 500\n'
            'If session.Children.Count > 1 Then\n'
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press\n'
            'End If\n'
            'On Error GoTo 0\n'
        )
        if ponto_ancora in conteudo:
            return conteudo.replace(ponto_ancora, ponto_ancora + "\n" + bloco_reset, 1)
        return bloco_reset + conteudo

    def _extrair_seriais(self, texto):
        seriais = []
        for linha in texto.splitlines():
            linha = linha.strip()
            if not linha: continue
            match_intervalo = re.search(r"(\d+)\s*(?:a|A|-|ate|at[eé])\s*(\d+)", linha, flags=re.IGNORECASE)
            if match_intervalo:
                inicio_txt, fim_txt = match_intervalo.groups()
                inicio, fim = int(inicio_txt), int(fim_txt)
                passo = 1 if fim >= inicio else -1
                largura = max(len(inicio_txt), len(fim_txt))
                preservar_zeros = inicio_txt.startswith("0") or fim_txt.startswith("0")
                for numero in range(inicio, fim + passo, passo):
                    seriais.append(str(numero).zfill(largura) if preservar_zeros else str(numero))
                continue
            numeros = re.findall(r"[A-Za-z0-9_\-]+", linha.replace(";", " "))
            seriais.extend(numeros)
        return self._remover_duplicados_mantendo_ordem(seriais)

    def _remover_duplicados_mantendo_ordem(self, seriais):
        vistos = set()
        return [s for s in seriais if not (s in vistos or vistos.add(s))]

    def _obter_lista_sem_ponto_virgula(self):
        texto = self.txt_sem_ponto_virgula.get("1.0", tk.END)
        return [linha.strip().replace(";", "") for linha in texto.splitlines() if linha.strip()]

    def _copiar_texto(self, widget, mensagem):
        texto = widget.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Não há conteúdo para copiar.")
            return
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        self.parent_frame.update()
        self._status(mensagem)

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)

        self.lbl_titulo = ctk.CTkLabel(self.container, text="🔢 Cole os intervalos ou seriais SAP abaixo:", font=("Segoe UI", 14, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 10))

        self.txt_entrada = ctk.CTkTextbox(self.container, height=120, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_entrada.pack(fill="both", expand=True, pady=(0, 15))

        self.btn_extrair = ctk.CTkButton(self.container, text="🔍 EXTRAIR SERIAL", command=self.extrair_seriais, font=("Segoe UI", 13, "bold"), height=40, corner_radius=6)
        self.btn_extrair.pack(fill="x", pady=(0, 15))

        self.frame_resultados = ctk.CTkFrame(self.container)
        self.frame_resultados.pack(fill="both", expand=True)
        self.frame_resultados.columnconfigure(0, weight=1)
        self.frame_resultados.columnconfigure(1, weight=1)
        self.frame_resultados.rowconfigure(1, weight=1)

        self.lbl_sem = ctk.CTkLabel(self.frame_resultados, text='Tabela IQ09 sem ";"')
        self.lbl_sem.grid(row=0, column=0, sticky="w")
        
        self.lbl_com = ctk.CTkLabel(self.frame_resultados, text='Formatado com ";"')
        self.lbl_com.grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.txt_sem_ponto_virgula = ctk.CTkTextbox(self.frame_resultados, height=150, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_sem_ponto_virgula.grid(row=1, column=0, sticky="nsew", pady=(6, 8))

        self.txt_com_ponto_virgula = ctk.CTkTextbox(self.frame_resultados, height=150, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_com_ponto_virgula.grid(row=1, column=1, sticky="nsew", padx=(12, 0), pady=(6, 8))

        self.frame_botoes_resultado = ctk.CTkFrame(self.frame_resultados)
        self.frame_botoes_resultado.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.frame_botoes_resultado.columnconfigure(0, weight=1)
        self.frame_botoes_resultado.columnconfigure(1, weight=1)

        self.btn_copiar_sem = ctk.CTkButton(self.frame_botoes_resultado, text='📋 Copiar tabela sem ";"', command=self.copiar_sem_ponto_virgula, font=("Segoe UI", 12, "bold"), height=35)
        self.btn_copiar_sem.grid(row=0, column=0, sticky="ew")

        self.btn_copiar_com = ctk.CTkButton(self.frame_botoes_resultado, text='📋 Copiar com ";"', command=self.copiar_com_ponto_virgula, font=("Segoe UI", 12, "bold"), height=35)
        self.btn_copiar_com.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        self.frame_acoes = ctk.CTkFrame(self.container)
        self.frame_acoes.pack(fill="x", pady=(15, 0))

        self.btn_iq09 = ctk.CTkButton(self.frame_acoes, text='🖥️ ABRIR NA IQ09 (Tabela sem ";")', command=self.abrir_na_iq09, font=("Segoe UI", 13, "bold"), height=40)
        self.btn_iq09.pack(side="left", fill="x", expand=True)

        self.btn_limpar = ctk.CTkButton(self.frame_acoes, text="🧹 Limpar", command=self.limpar, font=("Segoe UI", 12, "bold"), width=120, height=40)
        self.btn_limpar.pack(side="left", padx=(10, 0))

    def _set_texto(self, widget, texto):
        widget.delete("1.0", tk.END)
        widget.insert("1.0", texto)

    def _finalizar_iq09(self):
        self.executando_iq09 = False
        self.btn_iq09.configure(state="normal")

    def _status(self, texto):
        if self.label_status_global:
            # CustomTkinter label uses 'text' or configure(text=)
            self.label_status_global.configure(text=f" {texto}")

    def _status_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._status(texto))

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

def construir_aba_extrator(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrator(parent_frame, label_status_global, paletas, obter_tema_atual)
