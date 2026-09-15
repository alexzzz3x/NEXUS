import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

# Adicionar pasta raiz ao path para importar 'modulos'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from modulos import transferencia_osmes_final as transferencia_osmes
except ImportError:
    transferencia_osmes = None

from utils_metrics import incrementar_metrica
from utils_logger import log_error

try:
    from plyer import notification
except ImportError:
    notification = None

DEFAULT_PALETAS = {
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60"}
}

class AbaTransferenciaOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("📦 Aba Transferência de OSME pronta.")

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Conecta Blue"))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for widget in [self.txt_equipes, self.txt_log]:
            widget.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)

        self.btn_iniciar.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_limpar_log.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        
        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_log.configure(text_color=c["texto"])
        self.container.configure(fg_color="transparent")
        self.frame_botoes.configure(fg_color="transparent")

    def _status(self, msg):
        if self.label_status_global:
            self.label_status_global.configure(text=msg)

    def _status_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._status(msg))

    def _log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._log(msg))

    def iniciar_transferencia(self):
        if self.executando:
            messagebox.showwarning("Aviso", "O processo já está em execução.")
            return

        if not transferencia_osmes:
            messagebox.showerror("Erro", "Módulo 'transferencia_osmes_final' não encontrado na pasta 'modulos'.")
            return

        linhas = self.txt_equipes.get("1.0", tk.END).strip().splitlines()
        equipes = [l.strip().upper() for l in linhas if l.strip()]
        
        if not equipes:
            messagebox.showwarning("Aviso", "Insira pelo menos um código da equipe.")
            return

        processo_completo = getattr(transferencia_osmes, "MOVIMENTA_SAP", False)
        if processo_completo:
            confirmado = messagebox.askyesno(
                "Confirmar contabilização SAP",
                "Este processo inicia na ZUDWM_237 e registra de verdade os "
                "movimentos Z62 e Z61 no SAP, alterando estoque e ordens.\n\n"
                f"Fila: {len(equipes)} equipes\n\n"
                "A Z61 só inicia depois que o documento Z62 for capturado. "
                "Deseja continuar?"
            )
            if not confirmado:
                self._log("Execução cancelada pelo usuário.")
                return

        self.executando = True
        self.btn_iniciar.configure(state="disabled")
        self.txt_equipes.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_bar.pack(side="left", padx=15, fill="x", expand=True)
        self._status("Transferência de OSME em execução...")
        self._log(f"\n=============================================\n🏁 Iniciando fila de transferência para {len(equipes)} OSMEs...")

        def _run():
            sucessos = 0
            for idx, equipe in enumerate(equipes):
                try:
                    self._log_threadsafe(f"\n⏳ Processando {idx+1}/{len(equipes)}: {equipe}")
                    def on_seriais(serials):
                        self.parent_frame.after(0, self.mostrar_seriais, serials)
                    
                    transferencia_osmes.executar(equipe, self._log_threadsafe, serial_callback=on_seriais)
                    sucessos += 1
                except Exception as e:
                    log_error(f"ERRO OSME {equipe}: {e}")
                    self._log_threadsafe(f"ERRO na execução de {equipe}: {e}")
                
                progresso = (idx + 1) / len(equipes)
                self.parent_frame.after(0, lambda p=progresso: self.progress_bar.set(p))
            
            # Atualiza métricas
            if sucessos > 0:
                incrementar_metrica("osmes_transferidos", sucessos)
            
            # Notificação nativa
            if notification:
                try:
                    notification.notify(
                        title="Conecta Hub 9.0",
                        message=f"Fila concluída! {sucessos} OSMEs processados com sucesso.",
                        app_name="Conecta Hub",
                        timeout=5
                    )
                except: pass
                
            self.parent_frame.after(0, lambda: self._finalizar(sucessos, len(equipes)))

        threading.Thread(target=_run, daemon=True).start()

    def _finalizar(self, sucessos, total):
        self.executando = False
        self.btn_iniciar.configure(state="normal")
        self.txt_equipes.configure(state="normal")
        self.progress_bar.pack_forget()
        self._status("Pronto | Transferência de OSME")
        self._log(f"\n✅ Fila concluída! {sucessos} de {total} OSMEs transferidos com sucesso.")

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("📦 Log limpo.")

    def mostrar_seriais(self, serials):
        if not serials:
            return
        
        # CTkToplevel para manter o visual premium
        win = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        win.title("Materiais em OSMEs Vencidas")
        win.geometry("820x450")
        win.transient(self.parent_frame.winfo_toplevel())
        
        c = self._cores(self.obter_tema_atual())
        win.configure(fg_color=c.get("bg", "#1a202c"))
        
        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        txt_seriais = ctk.CTkTextbox(frame, font=("Consolas", 12), wrap="none", fg_color=c.get("input", "#2d3748"), text_color=c.get("texto", "white"))
        txt_seriais.pack(fill="both", expand=True)
        
        texto_completo = "\n".join(str(item) for item in serials)
        txt_seriais.insert("1.0", texto_completo)
        txt_seriais.configure(state="disabled")

        def copiar_todos():
            win.clipboard_clear()
            win.clipboard_append(texto_completo)
            win.update()
            messagebox.showinfo("Copiado", "Seriais copiados para a área de transferência!", parent=win)

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=10)
        
        btn_copiar = ctk.CTkButton(btn_frame, text="📋 Copiar Todos", command=copiar_todos, font=("Segoe UI", 12, "bold"), fg_color=c.get("accent", "#f57c00"))
        btn_copiar.pack(side="left", padx=5)
        
        btn_fechar = ctk.CTkButton(btn_frame, text="Fechar", command=win.destroy, font=("Segoe UI", 12, "bold"), fg_color=c.get("container", "#4a5568"))
        btn_fechar.pack(side="right", padx=5)

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)

        self.lbl_titulo = ctk.CTkLabel(self.container, text="📦 Transferência de OSME - ZUDWM_237", font=("Segoe UI", 16, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 10))

        frame_input = ctk.CTkFrame(self.container, fg_color="transparent")
        frame_input.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(frame_input, text="Códigos das Equipes\n(Um por linha):", font=("Segoe UI", 13)).pack(side="left", padx=(0, 10))
        
        self.txt_equipes = ctk.CTkTextbox(frame_input, font=("Consolas", 13), width=200, height=80, corner_radius=6)
        self.txt_equipes.pack(side="left")

        self.frame_botoes = ctk.CTkFrame(self.container)
        self.frame_botoes.pack(fill="x", pady=(0, 15))

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="▶ Executar Fila SAP", command=self.iniciar_transferencia, font=("Segoe UI", 13, "bold"), height=40, corner_radius=6)
        self.btn_iniciar.pack(side="left")
        
        self.progress_bar = ctk.CTkProgressBar(self.frame_botoes)

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🗑 Limpar Log", command=self.limpar_log, font=("Segoe UI", 13, "bold"), height=40, corner_radius=6)
        self.btn_limpar_log.pack(side="right")

        self.lbl_log = ctk.CTkLabel(self.container, text="Log de Execução", font=("Segoe UI", 13, "bold"))
        self.lbl_log.pack(anchor="w", pady=(5, 5))

        self.txt_log = ctk.CTkTextbox(self.container, font=("Consolas", 12), wrap="word", corner_radius=6)
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.configure(state="disabled")

def construir_aba_transferencia_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaTransferenciaOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
