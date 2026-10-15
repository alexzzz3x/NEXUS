import os
import sys
import re
import json
import datetime
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk

# Adicionar pasta raiz ao path para importar 'modulos'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from modulos import transferencia_osmes_final as transferencia_osmes
except ImportError:
    transferencia_osmes = None

from utils_metrics import incrementar_metrica
from utils_logger import log_error
from utils_os import obter_caminho_config

try:
    from plyer import notification
except ImportError:
    notification = None

DEFAULT_PALETAS = {
    "Nexus": {"bg": "#090d16", "container": "#111827", "texto": "#f8fafc", "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"},
    "Azul Delray": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"},
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"}
}

class AbaTransferenciaOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._renderizar_historico()
        self._log("📦 Aba Transferência de OSME pronta.")

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        inp = c.get("input", "#1a2436")
        texto = c.get("texto", "#f8fafc")
        accent = c.get("accent", "#00d2ff")
        hover = c.get("hover", "#0284c7")

        self.parent_frame.configure(fg_color=bg)

        for widget in [self.txt_equipes, self.txt_log, getattr(self, "entry_data_lancamento", None)]:
            if widget:
                widget.configure(fg_color=inp, text_color=texto, border_color=container, border_width=1)

        if hasattr(self, "lbl_dt_lcto"):
            self.lbl_dt_lcto.configure(text_color=texto)

        if hasattr(self, "chk_forcar_cl04"):
            self.chk_forcar_cl04.configure(text_color=texto, fg_color=accent, hover_color=hover)

        self.btn_iniciar.configure(fg_color=accent, hover_color=hover, text_color="#090d16" if "Light" not in tema_nome else "#ffffff")
        self.btn_limpar_log.configure(fg_color=container, hover_color=hover, text_color=texto)
        self.btn_limpar_hist.configure(fg_color=inp, hover_color=hover, text_color=texto)

        self.lbl_titulo.configure(text_color=texto)
        self.lbl_log.configure(text_color=texto)
        self.lbl_hist_tit.configure(text_color=accent)
        self.lbl_hist_hint.configure(text_color="#94a3b8")

        self.card_hist.configure(fg_color=container)

        estilo = ttk.Style()
        estilo.configure(
            "HistoricoTransfOSME.Treeview",
            background=container,
            foreground=texto,
            fieldbackground=inp,
            bordercolor=container,
            rowheight=26,
            relief="flat"
        )
        estilo.configure(
            "HistoricoTransfOSME.Treeview.Heading",
            background=inp,
            foreground=accent,
            font=("Segoe UI", 9, "bold")
        )
        estilo.map(
            "HistoricoTransfOSME.Treeview",
            background=[("selected", accent)],
            foreground=[("selected", "white")]
        )

    # ---------------------------------------------------------
    # Persistência de Histórico Recente
    # ---------------------------------------------------------
    def _caminho_historico(self):
        return obter_caminho_config("historico_transf_osme.json")

    def _carregar_historico(self):
        caminho = self._caminho_historico()
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    if isinstance(dados, list):
                        return dados
            except Exception:
                pass
        return []

    def _salvar_historico(self, itens):
        caminho = self._caminho_historico()
        try:
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(itens[:50], f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _adicionar_historico(self, equipe, status):
        hora = datetime.datetime.now().strftime("%H:%M")
        itens = self._carregar_historico()
        itens.insert(0, {
            "equipe": equipe,
            "horario": hora,
            "status": status
        })
        self._salvar_historico(itens)
        self._renderizar_historico(itens)

    def _limpar_historico(self):
        self._salvar_historico([])
        self._renderizar_historico([])

    def _renderizar_historico(self, itens=None):
        if itens is None:
            itens = self._carregar_historico()
        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)
        for d in itens:
            self.tree_hist.insert("", "end", values=(d.get("equipe", "-"), d.get("horario", "-"), d.get("status", "-")))

    def _ao_clique_duplo_historico(self, event):
        selecionado = self.tree_hist.selection()
        if selecionado:
            valores = self.tree_hist.item(selecionado[0], "values")
            if valores:
                eq = valores[0]
                conteudo_atual = self.txt_equipes.get("1.0", tk.END).strip()
                if conteudo_atual:
                    if eq not in conteudo_atual.splitlines():
                        self.txt_equipes.insert(tk.END, f"\n{eq}")
                else:
                    self.txt_equipes.insert(tk.END, eq)

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

    def obter_data_lancamento(self):
        """Retorna a data de lançamento informada na interface no formato dd.mm.aaaa, ou '' se em branco. Retorna None se inválida."""
        val = self.entry_data_lancamento.get().strip() if hasattr(self, "entry_data_lancamento") else ""
        if not val:
            return ""
        padrao = r"^(\d{2})\.(\d{2})\.(\d{4})$"
        m = re.match(padrao, val)
        if not m:
            messagebox.showerror(
                "Data de Lançamento Inválida",
                f"A data de lançamento informada ('{val}') está em formato incorreto.\n\n"
                f"Use o formato dd.mm.aaaa (ex: 01.10.2026) ou deixe o campo vazio para utilizar a data padrão de hoje.",
                parent=self.parent_frame.winfo_toplevel()
            )
            return None
        dia, mes, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            datetime.date(ano, mes, dia)
        except ValueError:
            messagebox.showerror(
                "Data Inválida",
                f"A data de lançamento '{val}' não corresponde a uma data válida de calendário.\n\n"
                f"Por favor, verifique o dia, mês e ano digitados.",
                parent=self.parent_frame.winfo_toplevel()
            )
            return None
        return val

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

        data_lcto = self.obter_data_lancamento()
        if data_lcto is None:
            return

        processo_completo = getattr(transferencia_osmes, "MOVIMENTA_SAP", False)
        if processo_completo:
            txt_dt = f"• Data de Lançamento (MIGO): {data_lcto} (Fechamento SAP)\n" if data_lcto else "• Data de Lançamento (MIGO): Hoje (Automático)\n"
            confirmado = messagebox.askyesno(
                "Confirmar contabilização SAP",
                "Este processo inicia na ZUDWM_237 e registra de verdade os "
                "movimentos Z62 e Z61 no SAP, alterando estoque e ordens.\n\n"
                f"Fila: {len(equipes)} equipes\n"
                f"{txt_dt}\n"
                "A Z61 só inicia depois que o documento Z62 for capturado. "
                "Deseja continuar?"
            )
            if not confirmado:
                self._log("Execução cancelada pelo usuário.")
                return

        forcar_cl04 = bool(self.chk_forcar_cl04.get()) if hasattr(self, "chk_forcar_cl04") else False

        self.executando = True
        self.btn_iniciar.configure(state="disabled")
        self.txt_equipes.configure(state="disabled")
        if hasattr(self, "entry_data_lancamento"):
            self.entry_data_lancamento.configure(state="disabled")
        if hasattr(self, "chk_forcar_cl04"):
            self.chk_forcar_cl04.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_bar.pack(side="left", padx=10, fill="x", expand=True)
        self._status("Transferência de OSME em execução...")
        if data_lcto:
            self._log(f"📅 Data de Lançamento configurada: {data_lcto}")
        if forcar_cl04:
            self._log("🏢 Opção ativa: Forçar Depósito CL04 / EMPR 1CL04 para toda a fila")
        self._log(f"\n=============================================\n🏁 Iniciando fila de transferência para {len(equipes)} OSMEs...")

        def _run():
            sucessos = 0
            for idx, equipe in enumerate(equipes):
                try:
                    self._log_threadsafe(f"\n⏳ Processando {idx+1}/{len(equipes)}: {equipe}")
                    eh_rcl = equipe.strip().upper().startswith("RCL") or forcar_cl04
                    if eh_rcl:
                        self._log_threadsafe("⚡ Rota Clandestina: Depósito CL04 | Recebedor EMPR 1CL04")
                    def on_seriais(serials):
                        self.parent_frame.after(0, self.mostrar_seriais, serials)

                    transferencia_osmes.executar(
                        equipe, self._log_threadsafe, serial_callback=on_seriais, data_lancamento=data_lcto, forcar_cl04=forcar_cl04
                    )
                    sucessos += 1
                    self.parent_frame.after(0, lambda eq=equipe: self._adicionar_historico(eq, "✅ Concluído"))
                except Exception as e:
                    log_error(f"ERRO OSME {equipe}: {e}")
                    self._log_threadsafe(f"ERRO na execução de {equipe}: {e}")
                    self.parent_frame.after(0, lambda eq=equipe: self._adicionar_historico(eq, "❌ Falha"))

                progresso = (idx + 1) / len(equipes)
                self.parent_frame.after(0, lambda p=progresso: self.progress_bar.set(p))

            if sucessos > 0:
                try:
                    incrementar_metrica("osmes_transferidos", sucessos)
                except Exception:
                    pass

            if notification:
                try:
                    notification.notify(
                        title="N.E.X.U.S",
                        message=f"Fila concluída! {sucessos} OSMEs processados com sucesso.",
                        app_name="N.E.X.U.S",
                        timeout=5
                    )
                except: pass

            self.parent_frame.after(0, lambda: self._finalizar(sucessos, len(equipes)))

        threading.Thread(target=_run, daemon=True).start()

    def _finalizar(self, sucessos, total):
        self.executando = False
        self.btn_iniciar.configure(state="normal")
        self.txt_equipes.configure(state="normal")
        if hasattr(self, "entry_data_lancamento"):
            self.entry_data_lancamento.configure(state="normal")
        if hasattr(self, "chk_forcar_cl04"):
            self.chk_forcar_cl04.configure(state="normal")
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

        win = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        win.title("Materiais em OSMEs Vencidas")
        win.geometry("820x450")
        win.transient(self.parent_frame.winfo_toplevel())

        c = self._cores(self.obter_tema_atual())
        win.configure(fg_color=c.get("bg", "#090d16"))

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        txt_seriais = ctk.CTkTextbox(frame, font=("Consolas", 12), wrap="none", fg_color=c.get("input", "#1a2436"), text_color=c.get("texto", "white"))
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

        btn_copiar = ctk.CTkButton(btn_frame, text="📋 Copiar Todos", command=copiar_todos, font=("Segoe UI", 12, "bold"), fg_color=c.get("accent", "#00d2ff"))
        btn_copiar.pack(side="left", padx=5)

        btn_fechar = ctk.CTkButton(btn_frame, text="Fechar", command=win.destroy, font=("Segoe UI", 12, "bold"), fg_color=c.get("container", "#111827"))
        btn_fechar.pack(side="right", padx=5)

    # ---------------------------------------------------------
    # Construção da Interface (2 Colunas: Operação & Histórico)
    # ---------------------------------------------------------
    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=16)

        self.container.grid_columnconfigure(0, weight=6)  # Coluna Esquerda: 60%
        self.container.grid_columnconfigure(1, weight=4)  # Coluna Direita: 40%
        self.container.grid_rowconfigure(0, weight=1)

        # 1. COLUNA ESQUERDA: Execução & Logs
        col_esq = ctk.CTkFrame(self.container, fg_color="transparent")
        col_esq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.lbl_titulo = ctk.CTkLabel(col_esq, text="📦 Códigos das Equipes (um por linha):", font=("Segoe UI", 13, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 6))

        self.txt_equipes = ctk.CTkTextbox(col_esq, height=110, font=("Consolas", 13), wrap="none", corner_radius=8)
        self.txt_equipes.pack(fill="x", pady=(0, 10))

        # Campo Data de Lançamento (Fechamento SAP / Opcional)
        f_dt_lcto = ctk.CTkFrame(col_esq, fg_color="transparent")
        f_dt_lcto.pack(fill="x", pady=(0, 10))

        self.lbl_dt_lcto = ctk.CTkLabel(
            f_dt_lcto,
            text="📅 Data Lçto. (Fechamento SAP):",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_dt_lcto.pack(side="left", padx=(0, 8))

        self.entry_data_lancamento = ctk.CTkEntry(
            f_dt_lcto,
            width=160,
            placeholder_text="dd.mm.aaaa (opcional)",
            font=("Consolas", 12)
        )
        self.entry_data_lancamento.pack(side="left")

        self.lbl_dt_lcto_hint = ctk.CTkLabel(
            f_dt_lcto,
            text="* Deixe vazio para usar a data de hoje",
            font=("Segoe UI", 10, "italic"),
            text_color="#94a3b8"
        )
        self.lbl_dt_lcto_hint.pack(side="left", padx=(10, 0))

        # Checkbox Forçar Depósito CL04
        self.chk_forcar_cl04 = ctk.CTkCheckBox(
            col_esq,
            text="🏢 Forçar Depósito CL04 / EMPR 1CL04 (Clandestinas RCL)",
            font=("Segoe UI", 11, "bold"),
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=4
        )
        self.chk_forcar_cl04.pack(anchor="w", pady=(0, 10))

        self.frame_botoes = ctk.CTkFrame(col_esq, fg_color="transparent")
        self.frame_botoes.pack(fill="x", pady=(0, 12))

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="▶ Executar Fila SAP", command=self.iniciar_transferencia, font=("Segoe UI", 13, "bold"), height=38, corner_radius=8)
        self.btn_iniciar.pack(side="left")

        self.progress_bar = ctk.CTkProgressBar(self.frame_botoes)

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🧹 Limpar Log", command=self.limpar_log, font=("Segoe UI", 11, "bold"), width=100, height=38, corner_radius=8)
        self.btn_limpar_log.pack(side="right")

        self.lbl_log = ctk.CTkLabel(col_esq, text="📋 Log de Execução:", font=("Segoe UI", 12, "bold"))
        self.lbl_log.pack(anchor="w", pady=(0, 4))

        self.txt_log = ctk.CTkTextbox(col_esq, font=("Consolas", 11), wrap="word", state="disabled", corner_radius=8)
        self.txt_log.pack(fill="both", expand=True)

        # 2. COLUNA DIREITA: Painel de Histórico Recente
        self.card_hist = ctk.CTkFrame(self.container, corner_radius=10)
        self.card_hist.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        frame_hist_header = ctk.CTkFrame(self.card_hist, fg_color="transparent")
        frame_hist_header.pack(fill="x", padx=12, pady=(12, 8))

        self.lbl_hist_tit = ctk.CTkLabel(frame_hist_header, text="🕒 Histórico de Transferências", font=("Segoe UI", 13, "bold"))
        self.lbl_hist_tit.pack(side="left")

        self.btn_limpar_hist = ctk.CTkButton(
            frame_hist_header,
            text="🗑️ Limpar",
            width=70,
            height=26,
            font=("Segoe UI", 10, "bold"),
            command=self._limpar_historico,
            corner_radius=6
        )
        self.btn_limpar_hist.pack(side="right")

        # Tabela Treeview de Histórico
        frame_tree = ctk.CTkFrame(self.card_hist, fg_color="transparent")
        frame_tree.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        colunas = ("equipe", "horario", "status")
        self.tree_hist = ttk.Treeview(
            frame_tree,
            columns=colunas,
            show="headings",
            style="HistoricoTransfOSME.Treeview",
            selectmode="browse"
        )
        self.tree_hist.heading("equipe", text="Equipe")
        self.tree_hist.heading("horario", text="Hora")
        self.tree_hist.heading("status", text="Status")

        self.tree_hist.column("equipe", width=120, anchor="center")
        self.tree_hist.column("horario", width=70, anchor="center")
        self.tree_hist.column("status", width=120, anchor="center")

        scroll_y = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree_hist.yview)
        self.tree_hist.configure(yscrollcommand=scroll_y.set)

        self.tree_hist.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.tree_hist.bind("<Double-1>", self._ao_clique_duplo_historico)

        self.lbl_hist_hint = ctk.CTkLabel(
            self.card_hist,
            text="💡 Duplo clique em um item para inseri-lo no campo.",
            font=("Segoe UI", 10, "italic")
        )
        self.lbl_hist_hint.pack(padx=12, pady=(0, 8))

def construir_aba_transferencia_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaTransferenciaOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
