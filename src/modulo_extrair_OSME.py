import os
import re
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

import pythoncom
import win32com.client
import win32con
import win32gui

from modulo_Sap import verificar_sap_aberto, set_sap_ocupado

DEFAULT_PALETAS = {
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60"}
}

from utils_os import obter_diretorio_executavel, obter_caminho_recurso, obter_caminho_script
from utils_metrics import incrementar_metrica

def obter_script_osme():
    return obter_caminho_script("Sciprt OSME2.vbs")

SCRIPT_OSME = obter_script_osme()

class AbaExtrairOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("📄 Aba Extrator OSME pronta.")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for widget in [self.txt_recursos, self.txt_log]:
            widget.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)

        self.btn_extrair.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_limpar_log.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        self.check_imprimir_excel.configure(fg_color=c["accent"], text_color=c["texto"])
        
        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_log.configure(text_color=c["texto"])
        self.container.configure(fg_color="transparent")
        self.frame_botoes.configure(fg_color="transparent")

    def iniciar_extracao(self):
        if self.executando:
            messagebox.showwarning("Aviso", "A extração OSME já está em execução.")
            return

        recursos = [linha.strip() for linha in self.txt_recursos.get("1.0", tk.END).splitlines() if linha.strip()]
        if not recursos:
            messagebox.showwarning("Aviso", "Insira ao menos uma equipe/recurso.")
            return

        script_osme = obter_script_osme()
        if not os.path.exists(script_osme):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_osme}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para realizar a extração OSME.\nAbra o SAP e tente novamente.")
            return

        self.executando = True
        self.btn_extrair.configure(state="disabled")
        self._status("Extração OSME em execução...")
        self._log("\n📄 Iniciando extração OSME...")

        thread = threading.Thread(target=self._executar_recursos, args=(recursos,), daemon=True)
        thread.start()

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("🧹 Log limpo.")

    def _executar_recursos(self, recursos):
        set_sap_ocupado(True)
        try:
            for recurso in recursos:
                self._log_threadsafe(f"\n⏳ Consultando OSME para: {recurso}...")
                script_temp = self._criar_script_temporario(recurso)
                
                try:
                    self._log_threadsafe("🔍 Executando extração OSME no SAP...")

                    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                    resultado = subprocess.run(
                        ["cscript.exe", "//nologo", script_temp],
                        capture_output=True, text=True, timeout=240, creationflags=creationflags,
                    )

                    saida = (resultado.stdout or "").strip()
                    if saida:
                        for linha in saida.splitlines():
                            self._log_threadsafe(linha.strip())

                    if resultado.returncode == 0:
                        if self.imprimir_excel_var.get():
                            self._log_threadsafe("🖨️ Enviando planilha do Excel para impressão...")
                            self._imprimir_excel_aberto()
                        self._log_threadsafe(f"✅ Extração OSME finalizada para: {recurso}")
                        incrementar_metrica("osmes_extraidos", 1)
                    else:
                        erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                        self._log_threadsafe(f"⚠️ Falha na extração OSME de {recurso}: {erro}")
                finally:
                    try: os.remove(script_temp)
                    except OSError: pass

            self._status_threadsafe("Extração OSME finalizada.")
            self._log_threadsafe("\n✅ Processo OSME concluído.")
        except subprocess.TimeoutExpired:
            self._log_threadsafe("⚠️ Tempo limite atingido ao executar o script OSME.")
            self._status_threadsafe("Extração OSME interrompida por tempo limite.")
        except Exception as e:
            self._log_threadsafe(f"⚠️ Erro inesperado: {e}")
            self._status_threadsafe("Erro na extração OSME.")
        finally:
            set_sap_ocupado(False)
            self.parent_frame.after(0, self._finalizar_execucao)

    def _criar_script_temporario(self, recurso):
        script_osme = obter_script_osme()
        with open(script_osme, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
        conteudo = self._inserir_reset_sap(conteudo)
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.text\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.caretPosition\s*=\s*)\d+',
            lambda match: f"{match.group(1)}{len(recurso)}",
            conteudo, count=1, flags=re.IGNORECASE,
        )
        fd, caminho = tempfile.mkstemp(prefix="osme_", suffix=".vbs", text=True)
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

    def _imprimir_excel_aberto(self):
        pythoncom.CoInitialize()
        self._log_threadsafe("🖨️ Focando na janela do Excel para impressão...")
        time.sleep(2.5)

        hwnd_excel = self._encontrar_hwnd_excel()
        if not hwnd_excel:
            self._log_threadsafe("⚠️ Janela do Excel não foi localizada na tela.")
            return

        try:
            win32gui.ShowWindow(hwnd_excel, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd_excel)
            time.sleep(0.8)
        except Exception as e_win:
            pass

        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            if excel:
                excel.DisplayAlerts = False
                excel.Visible = True
                if hasattr(excel, "ProtectedViewWindows") and excel.ProtectedViewWindows.Count > 0:
                    for i in range(1, excel.ProtectedViewWindows.Count + 1):
                        excel.ProtectedViewWindows(i).Edit()
                    time.sleep(1.0)
                wb = excel.ActiveWorkbook or excel.Workbooks(excel.Workbooks.Count)
                if wb:
                    sheet = wb.ActiveSheet or wb.Sheets(1)
                    sheet.PageSetup.Orientation = 2
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = False
        except Exception as e_com:
            pass

        try:
            win32gui.SetForegroundWindow(hwnd_excel)
            time.sleep(0.5)
        except Exception:
            pass

        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys("%p")
            time.sleep(0.4)
            shell.SendKeys("o")
            time.sleep(0.4)
            shell.SendKeys("{DOWN}")
            time.sleep(0.4)
            shell.SendKeys("{ENTER}")
            time.sleep(0.8)
            shell.SendKeys("^p")
            time.sleep(1.5)
            shell.SendKeys("{ENTER}")
            self._log_threadsafe("✅ Planilha configurada em PAISAGEM e impressa com sucesso!")
        except Exception as e_keys:
            self._log_threadsafe(f"⚠️ Erro ao enviar atalho no teclado: {e_keys}")

    def _encontrar_hwnd_excel(self, tentativas=20, intervalo=0.5):
        for _ in range(tentativas):
            janelas = []
            def enum_cb(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd):
                    cls = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                    if "XLMAIN" in cls or "Excel" in title or "XLS" in title or "ALV" in title:
                        janelas.append(hwnd)
            win32gui.EnumWindows(enum_cb, None)
            if janelas: return janelas[0]
            time.sleep(intervalo)
        return None

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame)
        self.container.pack(fill="both", expand=True, padx=40, pady=24)

        self.lbl_titulo = ctk.CTkLabel(self.container, text="📄 Insira os recursos/equipes para extrair OSME (um por linha):", font=("Segoe UI", 14, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 10))

        self.txt_recursos = ctk.CTkTextbox(self.container, height=140, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_recursos.pack(fill="x", pady=(0, 15))

        self.frame_botoes = ctk.CTkFrame(self.container)
        self.frame_botoes.pack(fill="x", pady=(0, 18))

        self.btn_extrair = ctk.CTkButton(self.frame_botoes, text="📄 Iniciar extração OSME", command=self.iniciar_extracao, font=("Segoe UI", 13, "bold"), height=40)
        self.btn_extrair.pack(side="left", fill="x", expand=True)

        self.imprimir_excel_var = tk.BooleanVar(value=True)
        self.check_imprimir_excel = ctk.CTkCheckBox(self.frame_botoes, text="🖨️ Imprimir Excel ao finalizar", variable=self.imprimir_excel_var, font=("Segoe UI", 12, "bold"))
        self.check_imprimir_excel.pack(side="left", padx=(15, 0))

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🧹 Limpar log", command=self.limpar_log, font=("Segoe UI", 12, "bold"), width=120, height=40)
        self.btn_limpar_log.pack(side="left", padx=(15, 0))

        self.lbl_log = ctk.CTkLabel(self.container, text="📋 Log de execução:", font=("Segoe UI", 12, "bold"))
        self.lbl_log.pack(anchor="w")

        self.txt_log = ctk.CTkTextbox(self.container, height=200, font=("Segoe UI", 12), wrap="word", state="disabled", corner_radius=6)
        self.txt_log.pack(fill="both", expand=True, pady=(6, 0))

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_extrair.configure(state="normal")

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

def construir_aba_extrair_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrairOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
