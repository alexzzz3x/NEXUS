# -*- coding: utf-8 -*-
"""
Interface Gráfica: Mudança de Recurso (Conecta Hub 9.0)
Permite a transferência integral de materiais e seriais entre duas equipes (Origem -> Destino):
- Centro: EMPR | Depósito: CL03 | Recebedor: EMPR 1CL03
- Suporte a temas dinâmicos (PALETAS)
- Execução não-bloqueante em thread dedicada
- Confirmação de segurança antes da contabilização
"""

import os
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk

# Garante acesso aos módulos na raiz do projeto (desenvolvimento e PyInstaller frozen)
if getattr(sys, "frozen", False):
    _RAIZ_PROJETO = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
else:
    _RAIZ_PROJETO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ_PROJETO not in sys.path:
    sys.path.insert(0, _RAIZ_PROJETO)

try:
    from modulos import mudanca_recurso
except ImportError:
    mudanca_recurso = None

from utils_metrics import incrementar_metrica
from utils_logger import log_error, log_info
from utils_ui import show_toast, play_sound

try:
    from plyer import notification
except ImportError:
    notification = None

DEFAULT_PALETAS = {
    "Nexus": {
        "bg": "#090d16",
        "container": "#111827",
        "texto": "#f8fafc",
        "accent": "#00d2ff",
        "hover": "#0284c7",
        "input": "#1a2436",
    },
    "Azul Delray": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
    },
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
    }
}


class AbaMudancaRecurso:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False
        self.ultimos_itens = []

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("🔄 Aba Mudança de Recurso pronta para operação.")

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

    def _status(self, msg):
        if self.label_status_global:
            try:
                self.label_status_global.configure(text=msg)
            except Exception:
                pass

    def _status_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._status(msg))

    def _log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, msg)
        if not msg.endswith("\n"):
            self.txt_log.insert(tk.END, "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._log(msg))

    def _construir_interface(self):
        # Container principal com bordas arredondadas
        self.container = ctk.CTkFrame(self.parent_frame, corner_radius=10)
        self.container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Cabeçalho da Aba
        self.frame_header = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_header.pack(fill=tk.X, padx=15, pady=(15, 10))

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_header,
            text="🔄 Mudança & Transferência de Recursos (ZUDWM_237 ➡️ MIGO ➡️ IW39)",
            font=("Segoe UI", 16, "bold")
        )
        self.lbl_titulo.pack(side=tk.LEFT)

        # Divisão em duas colunas: Esquerda (Entrada e Ações), Direita (Terminal de Log)
        self.frame_corpo = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_corpo.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # Coluna Esquerda: Formulário de Equipes
        self.col_esquerda = ctk.CTkFrame(self.frame_corpo, corner_radius=8)
        self.col_esquerda.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10), pady=0)
        self.col_esquerda.configure(width=380)
        self.col_esquerda.pack_propagate(False)

        # Card de Inputs
        self.lbl_sec_equipes = ctk.CTkLabel(
            self.col_esquerda,
            text="Identificação das Equipes",
            font=("Segoe UI", 13, "bold")
        )
        self.lbl_sec_equipes.pack(anchor="w", padx=15, pady=(15, 5))

        # Equipe Origem (Antiga)
        self.lbl_origem = ctk.CTkLabel(self.col_esquerda, text="Equipe de Origem (Antiga):", font=("Segoe UI", 12))
        self.lbl_origem.pack(anchor="w", padx=15, pady=(5, 2))
        
        self.ent_origem = ctk.CTkEntry(
            self.col_esquerda,
            placeholder_text="Ex: CSC395",
            font=("Consolas", 13, "bold"),
            height=36
        )
        self.ent_origem.pack(fill=tk.X, padx=15, pady=(0, 10))

        # Equipe Destino (Nova)
        self.lbl_destino = ctk.CTkLabel(self.col_esquerda, text="Equipe de Destino (Nova):", font=("Segoe UI", 12))
        self.lbl_destino.pack(anchor="w", padx=15, pady=(5, 2))

        self.ent_destino = ctk.CTkEntry(
            self.col_esquerda,
            placeholder_text="Ex: CVC396",
            font=("Consolas", 13, "bold"),
            height=36
        )
        self.ent_destino.pack(fill=tk.X, padx=15, pady=(0, 8))

        # Ordem Destino (Opcional / Fallback) com Botão Testar IW39
        self.lbl_ordem = ctk.CTkLabel(self.col_esquerda, text="Ordem Destino (opcional / fallback):", font=("Segoe UI", 12))
        self.lbl_ordem.pack(anchor="w", padx=15, pady=(2, 2))

        self.frame_ordem_row = ctk.CTkFrame(self.col_esquerda, fg_color="transparent")
        self.frame_ordem_row.pack(fill=tk.X, padx=15, pady=(0, 15))

        self.ent_ordem = ctk.CTkEntry(
            self.frame_ordem_row,
            placeholder_text="Auto na IW39 se vazio",
            font=("Consolas", 12),
            height=36
        )
        self.ent_ordem.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.btn_validar_ordem = ctk.CTkButton(
            self.frame_ordem_row,
            text="🔍 Testar IW39",
            width=105,
            height=36,
            command=self.acao_validar_ordem_destino,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6
        )
        self.btn_validar_ordem.pack(side=tk.RIGHT)

        # Painel de Parâmetros Fixos Padrão
        self.frame_badges = ctk.CTkFrame(self.col_esquerda, corner_radius=6)
        self.frame_badges.pack(fill=tk.X, padx=15, pady=(0, 15))

        self.lbl_badge_info = ctk.CTkLabel(
            self.frame_badges,
            text="📌 Padrão de Contabilização:\nCentro: EMPR  |  Depósito: CL03\nRecebedor: EMPR 1CL03  |  Operação: 10",
            font=("Segoe UI", 11, "italic"),
            justify="center"
        )
        self.lbl_badge_info.pack(padx=10, pady=8)

        # Botões de Ação
        self.btn_consultar = ctk.CTkButton(
            self.col_esquerda,
            text="🔍 Consultar Saldo Origem",
            command=self.acao_consultar_origem,
            font=("Segoe UI", 12, "bold"),
            height=38,
            corner_radius=6
        )
        self.btn_consultar.pack(fill=tk.X, padx=15, pady=(0, 8))

        self.btn_executar = ctk.CTkButton(
            self.col_esquerda,
            text="▶ Executar Mudança de Recurso",
            command=self.acao_executar_mudanca,
            font=("Segoe UI", 13, "bold"),
            height=42,
            corner_radius=6
        )
        self.btn_executar.pack(fill=tk.X, padx=15, pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(self.col_esquerda, height=12)
        self.progress_bar.pack(fill=tk.X, padx=15, pady=(0, 4))
        self.progress_bar.set(0)

        self.lbl_progresso = ctk.CTkLabel(
            self.col_esquerda,
            text="Aguardando início...",
            font=("Segoe UI", 11),
            text_color="#94a3b8",
            wraplength=340
        )
        self.lbl_progresso.pack(fill=tk.X, padx=15, pady=(0, 8))

        # Painel inferior de ações de log
        self.frame_acoes_log = ctk.CTkFrame(self.col_esquerda, fg_color="transparent")
        self.frame_acoes_log.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)

        self.btn_recuperar_z61 = ctk.CTkButton(
            self.frame_acoes_log,
            text="⚡ Concluir Saída Z61 (Recuperação)",
            command=self.abrir_janela_recuperacao_z61,
            font=("Segoe UI", 11, "bold"),
            height=34,
            corner_radius=6
        )
        self.btn_recuperar_z61.pack(fill=tk.X, pady=(0, 6))

        self.btn_salvar_log = ctk.CTkButton(
            self.frame_acoes_log,
            text="💾 Salvar Log (.txt)",
            command=self.salvar_log_txt,
            font=("Segoe UI", 11, "bold"),
            height=34,
            corner_radius=6
        )
        self.btn_salvar_log.pack(fill=tk.X, pady=(0, 6))

        self.btn_limpar_log = ctk.CTkButton(
            self.frame_acoes_log,
            text="🗑 Limpar Terminal",
            command=self.limpar_log,
            font=("Segoe UI", 11),
            height=30,
            corner_radius=6
        )
        self.btn_limpar_log.pack(fill=tk.X)

        # Coluna Direita: Terminal de Execução
        self.col_direita = ctk.CTkFrame(self.frame_corpo, corner_radius=8)
        self.col_direita.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0, pady=0)

        self.frame_log_header = ctk.CTkFrame(self.col_direita, fg_color="transparent")
        self.frame_log_header.pack(fill=tk.X, padx=15, pady=(10, 5))

        self.lbl_log_header = ctk.CTkLabel(
            self.frame_log_header,
            text="Terminal de Execução e Auditoria SAP",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_log_header.pack(side=tk.LEFT)

        self.btn_salvar_log_top = ctk.CTkButton(
            self.frame_log_header,
            text="💾 Baixar Log (.txt)",
            command=self.salvar_log_txt,
            font=("Segoe UI", 11, "bold"),
            height=28,
            width=135,
            corner_radius=6
        )
        self.btn_salvar_log_top.pack(side=tk.RIGHT)

        self.txt_log = ctk.CTkTextbox(
            self.col_direita,
            font=("Consolas", 11),
            wrap="word",
            corner_radius=6
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.txt_log.configure(state="disabled")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))
        self.container.configure(fg_color=c["container"])

        self.col_esquerda.configure(fg_color=c.get("input", "#1f3a60"))
        self.col_direita.configure(fg_color=c.get("input", "#1f3a60"))
        self.frame_badges.configure(fg_color=c.get("container", "#172a45"))

        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_sec_equipes.configure(text_color=c["texto"])
        self.lbl_origem.configure(text_color=c["texto"])
        self.lbl_destino.configure(text_color=c["texto"])
        if hasattr(self, "lbl_ordem"):
            self.lbl_ordem.configure(text_color=c["texto"])
        self.lbl_badge_info.configure(text_color=c.get("accent", "#f57c00"))
        self.lbl_log_header.configure(text_color=c["texto"])

        self.ent_origem.configure(
            fg_color=c.get("container", "#172a45"),
            text_color=c["texto"],
            border_color=c.get("accent", "#f57c00")
        )
        self.ent_destino.configure(
            fg_color=c.get("container", "#172a45"),
            text_color=c["texto"],
            border_color=c.get("accent", "#f57c00")
        )
        if hasattr(self, "ent_ordem"):
            self.ent_ordem.configure(
                fg_color=c.get("container", "#172a45"),
                text_color=c["texto"],
                border_color=c.get("accent", "#f57c00")
            )
        if hasattr(self, "btn_validar_ordem"):
            self.btn_validar_ordem.configure(
                fg_color=c.get("container", "#172a45"),
                hover_color=c.get("hover", "#e65100"),
                text_color=c["texto"]
            )
        if hasattr(self, "btn_recuperar_z61"):
            self.btn_recuperar_z61.configure(
                fg_color="#dc2626",
                hover_color="#b91c1c",
                text_color="#ffffff"
            )

        self.btn_consultar.configure(
            fg_color=c.get("container", "#172a45"),
            hover_color=c.get("hover", "#e65100"),
            text_color=c["texto"]
        )
        self.btn_executar.configure(
            fg_color=c.get("accent", "#f57c00"),
            hover_color=c.get("hover", "#e65100"),
            text_color="white"
        )
        self.btn_limpar_log.configure(
            fg_color=c.get("container", "#172a45"),
            hover_color=c.get("hover", "#e65100"),
            text_color=c["texto"]
        )
        self.btn_salvar_log.configure(
            fg_color=c.get("container", "#172a45"),
            hover_color=c.get("hover", "#e65100"),
            text_color=c["texto"]
        )
        self.btn_salvar_log_top.configure(
            fg_color=c.get("container", "#172a45"),
            hover_color=c.get("hover", "#e65100"),
            text_color=c["texto"]
        )

        self.progress_bar.configure(progress_color=c.get("accent", "#f57c00"))

        # Terminal sempre em fundo de alto contraste legível
        self.txt_log.configure(
            fg_color="#081120",
            text_color="#d9e4f2"
        )

    def salvar_log_txt(self):
        """Salva o conteúdo completo do terminal de log em um arquivo .txt."""
        conteudo = self.txt_log.get("1.0", tk.END).strip()
        if not conteudo or conteudo == "📦 Terminal limpo.":
            messagebox.showwarning("Aviso", "O terminal está vazio. Não há logs para exportar.", parent=self.parent_frame.winfo_toplevel())
            return

        origem = self.ent_origem.get().strip().upper() or "ORIGEM"
        destino = self.ent_destino.get().strip().upper() or "DESTINO"
        data_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_padrao = f"Log_MudancaRecurso_{origem}_{destino}_{data_str}.txt"

        caminho = filedialog.asksaveasfilename(
            parent=self.parent_frame.winfo_toplevel(),
            title="Salvar Log da Mudança de Recurso (.txt)",
            initialfile=nome_padrao,
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            try:
                with open(caminho, "w", encoding="utf-8") as f:
                    f.write(conteudo)
                messagebox.showinfo("Sucesso", f"Log exportado com sucesso em:\n{caminho}", parent=self.parent_frame.winfo_toplevel())
                show_toast(self.parent_frame, "Log exportado com sucesso!", type="success")
            except Exception as e:
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}", parent=self.parent_frame.winfo_toplevel())

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("📦 Terminal limpo.")

    def _progresso_threadsafe(self, etapa_num, percentual, texto):
        """Atualiza a barra de progresso e rótulo de etapa de forma thread-safe."""
        def _apply():
            try:
                self.progress_bar.set(percentual)
                self.lbl_progresso.configure(text=texto)
                self._status(texto)
            except Exception:
                pass
        self.parent_frame.after(0, _apply)

    def _bloquear_ui(self, bloqueado=True):
        self.executando = bloqueado
        estado = "disabled" if bloqueado else "normal"
        self.btn_consultar.configure(state=estado)
        self.btn_executar.configure(state=estado)
        self.ent_origem.configure(state=estado)
        self.ent_destino.configure(state=estado)
        if hasattr(self, "ent_ordem"):
            self.ent_ordem.configure(state=estado)
        if hasattr(self, "btn_validar_ordem"):
            self.btn_validar_ordem.configure(state=estado)
        if hasattr(self, "btn_recuperar_z61"):
            self.btn_recuperar_z61.configure(state=estado)

        if bloqueado:
            self.progress_bar.set(0.0)
            self.lbl_progresso.configure(text="Iniciando execução...")
        else:
            pass

    def acao_consultar_origem(self):
        """Apenas varre a ZUDWM_237 para exibir os materiais com saldo."""
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe um processo em execução.")
            return

        origem = self.ent_origem.get().strip().upper()
        if not origem:
            messagebox.showwarning("Aviso", "Por favor, informe a Equipe de Origem.")
            return

        if not mudanca_recurso:
            messagebox.showerror("Erro", "Motor 'mudanca_recurso' não encontrado.")
            return

        self._bloquear_ui(True)
        self._status("Consultando saldos na ZUDWM_237...")
        self._log(f"\n🔍 Iniciando consulta de saldos para a equipe {origem}...")

        def _worker():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                sessao = mudanca_recurso._conectar_sap()
                itens = mudanca_recurso.consultar_origem(sessao, origem, self._log_threadsafe)
                
                if itens:
                    self.ultimos_itens = itens
                    self._log_threadsafe(f"\n📊 RESUMO DOS MATERIAIS ENCONTRADOS ({len(itens)} itens):\n")
                    for i, it in enumerate(itens):
                        ser_txt = f" [{len(it['Seriais'])} seriais]" if it.get("Seriais") else ""
                        self._log_threadsafe(f"  {i+1}. {it['Material']} - {it['Descricao']} | Qtd: {it['Quantidade']} {it['Unidade']}{ser_txt} | Ordem: {it['Ordem Origem']}")
                    self._status_threadsafe(f"Consulta concluída: {len(itens)} itens encontrados.")
                    show_toast(self.parent_frame, f"{len(itens)} itens encontrados para {origem}!", type="success")
                    play_sound("done")
                else:
                    self.ultimos_itens = []
                    self._log_threadsafe(f"⚠️ Nenhum material com saldo positivo encontrado para {origem}.")
                    self._status_threadsafe("Nenhum saldo positivo.")
            except Exception as e:
                log_error(f"Erro na consulta de origem {origem}: {e}")
                self._log_threadsafe(f"\n❌ ERRO NA CONSULTA: {e}\n")
                self._status_threadsafe("Erro na consulta.")
                play_sound("error")
                messagebox.showerror("Erro", str(e))
            finally:
                self.parent_frame.after(0, lambda: self._bloquear_ui(False))

        threading.Thread(target=_worker, daemon=True).start()

    def acao_validar_ordem_destino(self):
        """Valida a equipe de destino na IW39 sem realizar nenhuma movimentação contábil."""
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe um processo em execução.")
            return

        destino = self.ent_destino.get().strip().upper()
        if not destino:
            messagebox.showwarning("Aviso", "Por favor, informe a Equipe de Destino.")
            return

        if not mudanca_recurso:
            messagebox.showerror("Erro", "Motor 'mudanca_recurso' não encontrado.")
            return

        self._bloquear_ui(True)
        self._status(f"Consultando ordem ativa para {destino} na IW39...")
        self._log(f"\n🔎 Testando consulta na IW39 para o recurso de destino '{destino}'...")

        def _worker():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                sessao = mudanca_recurso._conectar_sap()
                ordem = mudanca_recurso.buscar_ordem_destino_iw39(sessao, destino, self._log_threadsafe)

                def _atualizar():
                    self.ent_ordem.delete(0, tk.END)
                    self.ent_ordem.insert(0, ordem)
                    show_toast(self.parent_frame, f"Ordem {ordem} localizada para {destino}!", type="success")
                    play_sound("done")

                self.parent_frame.after(0, _atualizar)
                self._status_threadsafe(f"Ordem localizada: {ordem}")
            except Exception as e:
                log_error(f"Erro ao validar ordem de destino {destino}: {e}")
                self._log_threadsafe(f"\n❌ AVISO IW39: {e}\n")
                self._status_threadsafe("Ordem não localizada na IW39.")
                play_sound("error")
                messagebox.showwarning("Aviso IW39", f"Não foi possível localizar ordem ativa para {destino}:\n{e}")
            finally:
                self.parent_frame.after(0, lambda: self._bloquear_ui(False))

        threading.Thread(target=_worker, daemon=True).start()

    def acao_executar_mudanca(self):
        """Executa o ciclo completo de transferência Z62 -> IW39 -> Z61."""
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe um processo em execução.")
            return

        origem = self.ent_origem.get().strip().upper()
        destino = self.ent_destino.get().strip().upper()
        ordem_manual = self.ent_ordem.get().strip() if hasattr(self, "ent_ordem") else ""

        if not origem or not destino:
            messagebox.showwarning("Aviso", "Informe a Equipe de Origem e a Equipe de Destino.")
            return

        if origem == destino:
            messagebox.showwarning("Aviso", "A Equipe de Origem e a Equipe de Destino devem ser diferentes.")
            return

        if not mudanca_recurso:
            messagebox.showerror("Erro", "Motor 'mudanca_recurso' não encontrado.")
            return

        desc_ordem = f"Ordem informada: {ordem_manual}" if ordem_manual else "Ordem ativa será pré-validada na IW39 antes do Z62"
        confirmado = messagebox.askyesno(
            "Confirmar Mudança de Recurso SAP",
            f"Você confirma a transferência integral de recursos no SAP?\n\n"
            f"📤 Origem: {origem} (Estorno Z62 das ordens antigas)\n"
            f"📥 Destino: {destino} ({desc_ordem})\n\n"
            f"Centro: EMPR | Depósito: CL03 | Recebedor: EMPR 1CL03\n\n"
            f"🛡️ SEGURANÇA: A ordem de destino será 100% validada antes de realizar qualquer estorno no SAP.\n"
            f"Deseja prosseguir?",
            parent=self.parent_frame.winfo_toplevel()
        )
        if not confirmado:
            self._log("Execução cancelada pelo usuário.")
            return

        self._bloquear_ui(True)
        self._status(f"Mudança de recurso: {origem} ➡️ {destino} em execução...")

        def _worker():
            try:
                def on_seriais(seriais):
                    self.parent_frame.after(0, lambda: self.mostrar_seriais(seriais))

                resultado = mudanca_recurso.executar_mudanca_recurso(
                    recurso_antigo=origem,
                    recurso_novo=destino,
                    log_callback=self._log_threadsafe,
                    serial_callback=on_seriais,
                    progresso_callback=self._progresso_threadsafe,
                    ordem_destino_manual=ordem_manual
                )

                if resultado.get("status") == "sucesso":
                    incrementar_metrica("mudancas_recursos", 1)
                    self._progresso_threadsafe(4, 1.00, "🏆 Concluído (100%): Transferência realizada com sucesso!")
                    self._status_threadsafe(f"Pronto | Mudança {origem} ➡️ {destino} finalizada!")
                    show_toast(self.parent_frame, "Mudança de Recurso realizada com sucesso!", type="success")
                    play_sound("done")

                    if notification:
                        try:
                            notification.notify(
                                title="Conecta Hub 9.0",
                                message=f"Mudança de recurso {origem} ➡️ {destino} concluída com êxito!",
                                app_name="Conecta Hub",
                                timeout=6
                            )
                        except Exception:
                            pass
                else:
                    self._status_threadsafe(f"Concluído com aviso: {resultado.get('mensagem', '')}")

            except Exception as e:
                log_error(f"Erro na mudança de recurso {origem} -> {destino}: {e}")
                self._progresso_threadsafe(0, 0.0, f"❌ Falha: {e}")
                self._status_threadsafe("Erro na mudança de recurso.")
                play_sound("error")
                messagebox.showerror("Erro na Transferência SAP", str(e))
            finally:
                self.parent_frame.after(0, lambda: self._bloquear_ui(False))

        threading.Thread(target=_worker, daemon=True).start()

    def abrir_janela_recuperacao_z61(self):
        """Abre janela para concluir a Saída MIGO Z61 em casos onde o estorno Z62 já foi realizado."""
        win = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        win.title("⚡ Recuperação / Conclusão de Saída Z61")
        win.geometry("680x520")
        win.transient(self.parent_frame.winfo_toplevel())

        c = self._cores(self.obter_tema_atual())
        win.configure(fg_color=c.get("bg", "#0a192f"))

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=15)

        ctk.CTkLabel(
            frame,
            text="⚡ Concluir Saída Z61 no Destino (Modo de Recuperação)",
            font=("Segoe UI", 14, "bold"),
            text_color="#ef4444"
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            frame,
            text="Utilize quando o estorno Z62 já tiver sido efetuado na origem e faltar apenas lançar a saída Z61 no destino.",
            font=("Segoe UI", 11),
            text_color="#94a3b8"
        ).pack(anchor="w", pady=(0, 10))

        # Linha Equipe e Ordem
        f_top = ctk.CTkFrame(frame, fg_color="transparent")
        f_top.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_top, text="Recurso Destino:", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 5))
        ent_rec = ctk.CTkEntry(f_top, width=120, font=("Consolas", 12, "bold"))
        ent_rec.pack(side="left", padx=(0, 15))
        ent_rec.insert(0, self.ent_destino.get().strip().upper())

        ctk.CTkLabel(f_top, text="Ordem Destino:", font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 5))
        ent_ord = ctk.CTkEntry(f_top, width=160, font=("Consolas", 12, "bold"))
        ent_ord.pack(side="left", padx=(0, 10))
        if hasattr(self, "ent_ordem") and self.ent_ordem.get().strip():
            ent_ord.insert(0, self.ent_ordem.get().strip())

        # Materiais / Itens
        ctk.CTkLabel(
            frame,
            text="Itens a Imputar (Material;Qtd;Seriais separados por vírgula):",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", pady=(5, 2))

        txt_itens = ctk.CTkTextbox(frame, font=("Consolas", 11), height=180)
        txt_itens.pack(fill="both", expand=True, pady=(0, 10))

        # Preenche se tivermos itens na memória da última consulta
        if hasattr(self, "ultimos_itens") and self.ultimos_itens:
            linhas_exemplo = []
            for it in self.ultimos_itens:
                sers = ",".join(it.get("Seriais", []))
                linhas_exemplo.append(f"{it['Material']};{it['Quantidade']};{sers}")
            txt_itens.insert("1.0", "\n".join(linhas_exemplo))
        else:
            txt_itens.insert("1.0", "# Formato por linha: COD_MATERIAL;QUANTIDADE;SERIAL1,SERIAL2\n# Exemplo:\n# 510388;2;18398041,18398042\n")

        def executar_recuperacao():
            rec = ent_rec.get().strip().upper()
            ord_d = ent_ord.get().strip()
            if not rec or not ord_d:
                messagebox.showwarning("Aviso", "Informe o Recurso Destino e a Ordem Destino.", parent=win)
                return

            conteudo = txt_itens.get("1.0", tk.END).strip().splitlines()
            itens_para_saida = []
            for linha in conteudo:
                linha = linha.strip()
                if not linha or linha.startswith("#"):
                    continue
                partes = linha.split(";")
                if len(partes) < 2:
                    continue
                mat = partes[0].strip()
                qtd = partes[1].strip()
                sers = [s.strip() for s in partes[2].split(",") if s.strip()] if len(partes) > 2 else []
                itens_para_saida.append({
                    "Material": mat,
                    "Quantidade": qtd,
                    "Seriais": sers,
                    "Ordem Origem": ord_d
                })

            if not itens_para_saida:
                messagebox.showwarning("Aviso", "Nenhum item válido informado na lista.", parent=win)
                return

            conf = messagebox.askyesno(
                "Confirmar Saída Z61",
                f"Deseja executar a MIGO Z61 para o recurso {rec} na ordem {ord_d} com {len(itens_para_saida)} item(ns)?",
                parent=win
            )
            if not conf:
                return

            win.destroy()
            self._bloquear_ui(True)
            self._log(f"\n⚡ Iniciando Saída Z61 Direta de Recuperação para {rec} na ordem {ord_d}...")

            def _w_rec():
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    res = mudanca_recurso.executar_apenas_saida_z61(
                        recurso_destino=rec,
                        ordem_destino=ord_d,
                        itens=itens_para_saida,
                        log_callback=self._log_threadsafe,
                        progresso_callback=self._progresso_threadsafe
                    )
                    self._status_threadsafe(f"Saída Z61 gerada com sucesso: {res.get('doc_z61')}")
                    show_toast(self.parent_frame, f"Saída Z61 gerada: {res.get('doc_z61')}!", type="success")
                    play_sound("done")
                except Exception as ex:
                    log_error(f"Erro na recuperação Z61: {ex}")
                    self._log_threadsafe(f"\n❌ ERRO NA SAÍDA Z61: {ex}\n")
                    play_sound("error")
                    messagebox.showerror("Erro na Saída Z61", str(ex))
                finally:
                    self.parent_frame.after(0, lambda: self._bloquear_ui(False))

            threading.Thread(target=_w_rec, daemon=True).start()

        f_b = ctk.CTkFrame(frame, fg_color="transparent")
        f_b.pack(fill="x", pady=(5, 0))

        ctk.CTkButton(
            f_b,
            text="▶ Executar MIGO Z61",
            command=executar_recuperacao,
            font=("Segoe UI", 12, "bold"),
            fg_color="#ef4444",
            hover_color="#dc2626",
            height=36
        ).pack(side="left")

        ctk.CTkButton(
            f_b,
            text="Cancelar",
            command=win.destroy,
            font=("Segoe UI", 11),
            height=36
        ).pack(side="right")

    def mostrar_seriais(self, serials):
        """Abre janela para conferência e cópia dos seriais transferidos."""
        if not serials:
            return

        win = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        win.title("Números de Série Transferidos")
        win.geometry("650x400")
        win.transient(self.parent_frame.winfo_toplevel())

        c = self._cores(self.obter_tema_atual())
        win.configure(fg_color=c.get("bg", "#0a192f"))

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        lbl = ctk.CTkLabel(frame, text=f"Total de seriais: {len(serials)}", font=("Segoe UI", 12, "bold"))
        lbl.pack(anchor="w", pady=(0, 5))

        txt = ctk.CTkTextbox(frame, font=("Consolas", 11), wrap="none")
        txt.pack(fill="both", expand=True)
        texto_seriais = "\n".join(serials)
        txt.insert("1.0", texto_seriais)
        txt.configure(state="disabled")

        def copiar():
            win.clipboard_clear()
            win.clipboard_append(texto_seriais)
            win.update()
            messagebox.showinfo("Copiado", "Seriais copiados com sucesso!", parent=win)

        btn_box = ctk.CTkFrame(win, fg_color="transparent")
        btn_box.pack(fill="x", padx=15, pady=10)

        ctk.CTkButton(btn_box, text="📋 Copiar Seriais", command=copiar, font=("Segoe UI", 11, "bold")).pack(side="left")
        ctk.CTkButton(btn_box, text="Fechar", command=win.destroy, font=("Segoe UI", 11)).pack(side="right")


def construir_aba_mudanca_recurso(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaMudancaRecurso(parent_frame, label_status_global, paletas, obter_tema_atual)
