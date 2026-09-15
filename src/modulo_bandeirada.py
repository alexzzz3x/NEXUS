import os
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

from modulo_Sap import verificar_sap_aberto, set_sap_ocupado
from utils_os import obter_caminho_script

def obter_script_bandeirada():
    return obter_caminho_script("SCRIPT DA BANDEIRADA.vbs")

SCRIPT_BANDEIRADA = obter_script_bandeirada()

class AbaBandeirada:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("🚩 Aba BANDEIRADA pronta.")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for widget in [self.txt_recursos, self.txt_log]:
            widget.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)

        self.btn_iniciar.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_limpar_log.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        
        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_log.configure(text_color=c["texto"])
        self.container.configure(fg_color="transparent")
        self.frame_botoes.configure(fg_color="transparent")

    def iniciar_bandeirada(self):
        if self.executando:
            messagebox.showwarning("Aviso", "A bandeirada já está em execução.")
            return

        recursos = [linha.strip() for linha in self.txt_recursos.get("1.0", tk.END).splitlines() if linha.strip()]
        if not recursos:
            messagebox.showwarning("Aviso", "Insira ao menos uma equipe/recurso.")
            return

        script_bandeirada = obter_script_bandeirada()
        if not os.path.exists(script_bandeirada):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_bandeirada}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para executar a Bandeirada.\nAbra o SAP e tente novamente.")
            return

        self.executando = True
        self.btn_iniciar.configure(state="disabled")
        self._status("Bandeirada em execução...")
        self._log("\n🏁 Iniciando bandeirada inteligente...")

        thread = threading.Thread(target=self._executar_recursos, args=(recursos,), daemon=True)
        thread.start()

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("🚩 Log limpo.")

    def _executar_recursos(self, recursos):
        set_sap_ocupado(True)
        try:
            for recurso in recursos:
                self._log_threadsafe(f"\n⏳ Pesquisando ORDEM para: {recurso}...")
                script_temp = self._criar_script_temporario(recurso)
                
                try:
                    self._log_threadsafe("🔍 Executando consulta no SAP para obter a ordem do recurso...")

                    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                    resultado = subprocess.run(
                        ["cscript.exe", "//nologo", script_temp],
                        capture_output=True, text=True, timeout=180, creationflags=creationflags,
                    )

                    saida = (resultado.stdout or "").strip()
                    if saida:
                        for linha in saida.splitlines():
                            self._log_threadsafe(linha.strip())

                    if resultado.returncode == 0:
                        self._log_threadsafe(f"✅ Bandeirada finalizada para: {recurso}")
                    else:
                        erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                        self._log_threadsafe(f"⚠️ Falha na bandeirada de {recurso}: {erro}")
                finally:
                    try: os.remove(script_temp)
                    except OSError: pass

            self._status_threadsafe("Bandeirada finalizada.")
            self._log_threadsafe("\n✅ Processo de bandeirada concluído.")
        except subprocess.TimeoutExpired:
            self._log_threadsafe("⚠️ Tempo limite atingido ao executar o script VBS.")
            self._status_threadsafe("Bandeirada interrompida por tempo limite.")
        except Exception as e:
            self._log_threadsafe(f"⚠️ Erro inesperado: {e}")
            self._status_threadsafe("Erro na bandeirada.")
        finally:
            set_sap_ocupado(False)
            self.parent_frame.after(0, self._finalizar_execucao)

    def _criar_script_temporario(self, recurso):
        script_bandeirada = obter_script_bandeirada()
        with open(script_bandeirada, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        conteudo = self._inserir_reset_sap(conteudo)
        conteudo = re.sub(
            r'(recurso\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtDY_PARNR"\)\.text\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtDY_PARNR"\)\.caretPosition\s*=\s*)\d+',
            lambda match: f"{match.group(1)}{len(recurso)}",
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = conteudo.replace("{RECURSO}", recurso).replace("{EQUIPE}", recurso)
        conteudo = self._inserir_captura_ordem(conteudo)
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/txtP_AUFNR"\)\.text\s*=\s*)".*?"',
            r"\1ordemBandeirada",
            conteudo, count=1, flags=re.IGNORECASE,
        )

        fd, caminho = tempfile.mkstemp(prefix="bandeirada_", suffix=".vbs", text=True)
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

    def _inserir_captura_ordem(self, conteudo):
        if "ordemBandeirada" in conteudo: return conteudo
        ponto_ancora = 'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu'
        bloco_captura = (
            'Dim ordemBandeirada\n'
            'ordemBandeirada = ""\n'
            'Dim gridBandeirada\n'
            'Set gridBandeirada = Nothing\n'
            'On Error Resume Next\n'
            'Set gridBandeirada = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")\n'
            'On Error GoTo 0\n'
            'If Not gridBandeirada Is Nothing Then\n'
            '    On Error Resume Next\n'
            '    gridBandeirada.setCurrentCell -1, "GSTRP"\n'
            '    gridBandeirada.selectColumn "GSTRP"\n'
            '    session.findById("wnd[0]/tbar[1]/btn[40]").press\n'
            '    WScript.Sleep 300\n'
            '    Set gridBandeirada = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")\n'
            '    ordemBandeirada = Trim(gridBandeirada.GetCellValue(0, "AUFNR"))\n'
            '    If Err.Number <> 0 Or ordemBandeirada = "" Then\n'
            '        Err.Clear\n'
            '        ordemBandeirada = Trim(gridBandeirada.GetCellValue(0, "ORDERID"))\n'
            '    End If\n'
            '    On Error GoTo 0\n'
            'End If\n'
            'If Trim(ordemBandeirada) = "" Then\n'
            '    Dim txtOrdemDet\n'
            '    Set txtOrdemDet = Nothing\n'
            '    On Error Resume Next\n'
            '    Set txtOrdemDet = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR")\n'
            '    If txtOrdemDet Is Nothing Then\n'
            '        Set txtOrdemDet = session.findById("wnd[0]/usr/txtCAUFVD-AUFNR")\n'
            '    End If\n'
            '    If Not txtOrdemDet Is Nothing Then\n'
            '        ordemBandeirada = Trim(txtOrdemDet.text)\n'
            '    End If\n'
            '    On Error GoTo 0\n'
            'End If\n'
            'If Trim(ordemBandeirada) = "" Then\n'
            '    WScript.Echo "Nenhuma ordem encontrada para o recurso pesquisado."\n'
            '    WScript.Quit 2\n'
            'End If\n'
            'WScript.Echo "Ordem encontrada: " & ordemBandeirada & ". Bandeirando..."\n'
        )
        if ponto_ancora in conteudo:
            substituicao = bloco_captura + 'On Error Resume Next\n' + ponto_ancora
            return conteudo.replace(ponto_ancora, substituicao, 1)
        return bloco_captura + conteudo

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame)
        self.container.pack(fill="both", expand=True, padx=40, pady=24)

        self.lbl_titulo = ctk.CTkLabel(self.container, text="🚩 Insira os recursos/equipes (um por linha):", font=("Segoe UI", 14, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 10))

        self.txt_recursos = ctk.CTkTextbox(self.container, height=140, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_recursos.pack(fill="x", pady=(0, 15))

        self.frame_botoes = ctk.CTkFrame(self.container)
        self.frame_botoes.pack(fill="x", pady=(0, 18))

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="🏁 Iniciar processo de bandeirada", command=self.iniciar_bandeirada, font=("Segoe UI", 13, "bold"), height=40)
        self.btn_iniciar.pack(side="left", fill="x", expand=True)

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🧹 Limpar log", command=self.limpar_log, font=("Segoe UI", 12, "bold"), width=120, height=40)
        self.btn_limpar_log.pack(side="left", padx=(15, 0))

        self.lbl_log = ctk.CTkLabel(self.container, text="📋 Log de execução:", font=("Segoe UI", 12, "bold"))
        self.lbl_log.pack(anchor="w")

        self.txt_log = ctk.CTkTextbox(self.container, height=200, font=("Segoe UI", 12), wrap="word", state="disabled", corner_radius=6)
        self.txt_log.pack(fill="both", expand=True, pady=(6, 0))

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_iniciar.configure(state="normal")

    def _log(self, texto):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, texto + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._log(texto))

    def _status(self, texto):
        if self.label_status_global:
            self.label_status_global.configure(text=f" {texto}")

    def _status_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._status(texto))

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

def construir_aba_bandeirada(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaBandeirada(parent_frame, label_status_global, paletas, obter_tema_atual)
