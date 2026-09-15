# -*- coding: utf-8 -*-
"""
Aba do Módulo de Relatório de Fichas Plataforma (Manhã x Tarde) para o Conecta Hub.
Permite carregar dados brutos, pré-visualizar indicadores e tabela resumo,
copiar para e-mail/Teams com 1 clique e mandar pro Excel gerando o arquivo idêntico ao modelo oficial.
"""

import os
import sys
import threading
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

# Adiciona pasta raiz ao path para importar 'modulos'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from modulos import compilador_plataforma
except ImportError:
    compilador_plataforma = None

try:
    import win32clipboard
except ImportError:
    win32clipboard = None

from utils_metrics import incrementar_metrica
from utils_logger import log_error, log_info
from utils_ui import show_toast, play_sound


DEFAULT_PALETAS = {
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#f57c00",
        "hover": "#e65100",
        "input": "#1f3a60"
    }
}


def obter_palette_dict(tema_nome, paletas):
    c = paletas.get(tema_nome, paletas.get("Conecta Blue", DEFAULT_PALETAS["Conecta Blue"]))
    is_light = "Light" in tema_nome or c.get("bg", "").lower() in ["#ffffff", "#f3f6fb"]

    return {
        "navy": c.get("container", "#07183F"),
        "navy_soft": c.get("input", "#0D285F"),
        "blue": c.get("accent", "#1746A2"),
        "blue_hover": c.get("hover", "#103A8C"),
        "blue_pale": "#EAF1FF" if is_light else c.get("input", "#1f3a60"),
        "orange": "#F47721",
        "orange_hover": "#D95C0B",
        "orange_pale": "#FFF0E5",
        "page": c.get("bg", "#F3F6FA"),
        "surface": "#FFFFFF" if is_light else c.get("container", "#172a45"),
        "surface_alt": "#F8FAFD" if is_light else c.get("input", "#1f3a60"),
        "border": c.get("container", "#DCE4EF"),
        "text": "#15213A" if is_light else c.get("texto", "#ffffff"),
        "muted": "#69758A",
        "success": "#149B68",
        "success_pale": "#E8F7F1",
        "danger": "#D92D20",
        "danger_pale": "#FDECEA",
        "terminal": "#081120",
        "terminal_text": "#D9E4F2"
    }


class AbaRelatorioPlataforma:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False
        self.resultado_atual = None
        self.caminho_ultimo_excel = None

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("📊 Módulo de Relatório Plataforma (Manhã x Tarde) inicializado.")
        self._verificar_arquivo_padrao()

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Conecta Blue"))

    def _log(self, mensagem):
        try:
            self.txt_terminal.configure(state="normal")
            self.txt_terminal.insert("end", f"{mensagem}\n")
            self.txt_terminal.see("end")
            self.txt_terminal.configure(state="disabled")
        except Exception:
            pass

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame, corner_radius=10)
        self.container.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # 1. CABEÇALHO
        self.frame_header = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_header.pack(fill=tk.X, padx=15, pady=(10, 4))

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_header,
            text="📈 Relatório de Fichas Plataforma (Manhã x Tarde)",
            font=("Segoe UI", 16, "bold")
        )
        self.lbl_titulo.pack(side=tk.LEFT)

        self.lbl_desc = ctk.CTkLabel(
            self.frame_header,
            text="Consolidação e exportação oficial (idêntica ao modelo Exemplo.xlsm)",
            font=("Segoe UI", 11, "italic")
        )
        self.lbl_desc.pack(side=tk.LEFT, padx=15)

        # 2. BARRA DE ARQUIVO E CONTROLES
        self.frame_top_bar = ctk.CTkFrame(self.container, corner_radius=8)
        self.frame_top_bar.pack(fill=tk.X, padx=15, pady=(0, 6))

        self.lbl_origem = ctk.CTkLabel(
            self.frame_top_bar,
            text="Arquivo Bruto:",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_origem.pack(side=tk.LEFT, padx=(12, 6), pady=8)

        self.entry_arquivo = ctk.CTkEntry(
            self.frame_top_bar,
            placeholder_text="Selecione a planilha de dados brutos (.xlsx, .xlsm, .csv)...",
            font=("Segoe UI", 11),
            width=360
        )
        self.entry_arquivo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), pady=8)

        self.btn_procurar = ctk.CTkButton(
            self.frame_top_bar,
            text="📂 Procurar",
            width=90,
            font=("Segoe UI", 11, "bold"),
            command=self._selecionar_arquivo
        )
        self.btn_procurar.pack(side=tk.LEFT, padx=3, pady=8)

        self.btn_abrir_pasta_data = ctk.CTkButton(
            self.frame_top_bar,
            text="📁 Pasta Data",
            width=90,
            font=("Segoe UI", 11, "bold"),
            command=self._abrir_pasta_data
        )
        self.btn_abrir_pasta_data.pack(side=tk.LEFT, padx=3, pady=8)

        self.btn_processar = ctk.CTkButton(
            self.frame_top_bar,
            text="⚡ Compilar em Tela",
            width=130,
            font=("Segoe UI", 11, "bold"),
            command=self._iniciar_processamento_async
        )
        self.btn_processar.pack(side=tk.LEFT, padx=(3, 12), pady=8)

        # 3. CARDS DE KPIS
        self.frame_kpis = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_kpis.pack(fill=tk.X, padx=15, pady=(0, 6))
        self.frame_kpis.columnconfigure((0, 1, 2, 3), weight=1, uniform="kpi")

        self.card_total = self._criar_card_kpi(self.frame_kpis, 0, "📋 TOTAL FICHAS", "0", "Aguardando arquivo")
        self.card_manha = self._criar_card_kpi(self.frame_kpis, 1, "☀️ ENCERRADAS MANHÃ", "0", "0% do total")
        self.card_tarde = self._criar_card_kpi(self.frame_kpis, 2, "🌙 ENCERRADAS TARDE", "0", "0% do total")
        self.card_rapidas = self._criar_card_kpi(self.frame_kpis, 3, "⚡ RÁPIDAS MANHÃ (≤2H)", "0", "Média: 0 min")

        # 4. TABELA RESUMO PREVIEW
        self.frame_central = ctk.CTkFrame(self.container, corner_radius=8)
        self.frame_central.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 6))

        self.frame_titulo_tabela = ctk.CTkFrame(self.frame_central, height=34, corner_radius=6, fg_color="#0F3252")
        self.frame_titulo_tabela.pack(fill=tk.X, padx=8, pady=(6, 4))
        self.lbl_titulo_tabela = ctk.CTkLabel(
            self.frame_titulo_tabela,
            text="FICHAS POR SEGMENTO | ENCERRAMENTO MANHA x TARDE",
            font=("Segoe UI", 12, "bold"),
            text_color="#FFFFFF"
        )
        self.lbl_titulo_tabela.pack(expand=True, pady=4)

        self.frame_tree = tk.Frame(self.frame_central, bg="#172a45")
        self.frame_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        colunas = ("segmento", "total", "manha", "pct_manha", "tarde", "pct_tarde")
        self.tree = ttk.Treeview(self.frame_tree, columns=colunas, show="headings", height=5)
        
        self.tree.heading("segmento", text="SEGMENTO")
        self.tree.heading("total", text="TOTAL FICHAS")
        self.tree.heading("manha", text="ENCERRADAS MANHA")
        self.tree.heading("pct_manha", text="%")
        self.tree.heading("tarde", text="ENCERRADAS A TARDE")
        self.tree.heading("pct_tarde", text="%")

        self.tree.column("segmento", width=180, anchor="w")
        self.tree.column("total", width=120, anchor="center")
        self.tree.column("manha", width=140, anchor="center")
        self.tree.column("pct_manha", width=90, anchor="center")
        self.tree.column("tarde", width=150, anchor="center")
        self.tree.column("pct_tarde", width=90, anchor="center")

        self.tree_scroll = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("total_row", background="#0F3252", foreground="#FFFFFF", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("even_row", background="#F8FAFC", foreground="#0F172A", font=("Segoe UI", 10))
        self.tree.tag_configure("odd_row", background="#FFFFFF", foreground="#0F172A", font=("Segoe UI", 10))

        # 5. TERMINAL DE ATIVIDADE
        self.frame_terminal_container = ctk.CTkFrame(self.container, height=80, corner_radius=6)
        self.frame_terminal_container.pack(fill=tk.X, padx=15, pady=(0, 6))

        self.txt_terminal = ctk.CTkTextbox(
            self.frame_terminal_container,
            height=70,
            font=("Consolas", 10),
            corner_radius=6
        )
        self.txt_terminal.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.txt_terminal.configure(state="disabled")

        # 6. BARRA DE AÇÕES INFERIOR
        self.frame_acoes = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_acoes.pack(fill=tk.X, padx=15, pady=(0, 8))

        # BOTÃO PRINCIPAL EM DESTAQUE (MANDA PRO EXCEL)
        self.btn_manda_pro_excel = ctk.CTkButton(
            self.frame_acoes,
            text="🚀 Mandar pro Excel (Gerar & Abrir)",
            font=("Segoe UI", 12, "bold"),
            command=self._mandar_pro_excel_direto,
            fg_color="#107C41",  # Verde Excel
            hover_color="#0B5A2F",
            height=36,
            width=260
        )
        self.btn_manda_pro_excel.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_copiar = ctk.CTkButton(
            self.frame_acoes,
            text="📋 Copiar Tabela (E-mail / Teams)",
            font=("Segoe UI", 11, "bold"),
            command=self._copiar_para_clipboard,
            height=36,
            width=210
        )
        self.btn_copiar.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_salvar_como = ctk.CTkButton(
            self.frame_acoes,
            text="💾 Salvar Como...",
            font=("Segoe UI", 11, "bold"),
            command=self._exportar_excel,
            height=36,
            width=130
        )
        self.btn_salvar_como.pack(side=tk.LEFT, padx=(0, 10))

        self.lbl_status_local = ctk.CTkLabel(
            self.frame_acoes,
            text="Pronto para compilar e mandar pro Excel.",
            font=("Segoe UI", 10, "italic")
        )
        self.lbl_status_local.pack(side=tk.RIGHT, padx=10)

    def _criar_card_kpi(self, parent, coluna, titulo, valor_padrao, subtexto_padrao):
        card = ctk.CTkFrame(parent, corner_radius=8)
        card.grid(row=0, column=coluna, padx=4, pady=0, sticky="nsew")

        lbl_t = ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 9, "bold"), text_color="#94A3B8")
        lbl_t.pack(anchor="w", padx=10, pady=(6, 0))

        lbl_v = ctk.CTkLabel(card, text=valor_padrao, font=("Segoe UI", 18, "bold"))
        lbl_v.pack(anchor="w", padx=10, pady=(0, 0))

        lbl_s = ctk.CTkLabel(card, text=subtexto_padrao, font=("Segoe UI", 9, "italic"), text_color="#64748B")
        lbl_s.pack(anchor="w", padx=10, pady=(0, 6))

        return {"frame": card, "lbl_t": lbl_t, "lbl_v": lbl_v, "lbl_s": lbl_s}

    def _abrir_pasta_data(self):
        pasta_data = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
        os.makedirs(pasta_data, exist_ok=True)
        try:
            os.startfile(pasta_data)
            self._log(f"📁 Pasta de dados aberta no Explorer: {pasta_data}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta data:\n{e}")

    def _verificar_arquivo_padrao(self):
        """Detecta automaticamente arquivo recente em data, Downloads ou Desktop."""
        pasta_data = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
        caminhos_sugeridos = [
            os.path.join(pasta_data, "DADOS_BRUTOS_PLATAFORMA.xlsx"),
            os.path.join(os.path.expanduser("~"), "Downloads", "exemplo.xlsm"),
            r"C:\Users\anderson.vieira\Desktop\Planilha do pedro\PLANILHA_PLATAFORMA_AUTOMATICA.xlsm",
            r"C:\Users\anderson.vieira\Desktop\pedidoscl04.xlsx",
        ]

        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        if os.path.exists(pasta_downloads):
            try:
                for f in sorted(os.listdir(pasta_downloads), key=lambda x: os.path.getmtime(os.path.join(pasta_downloads, x)) if os.path.exists(os.path.join(pasta_downloads, x)) else 0, reverse=True):
                    if f.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")) and any(k in f.lower() for k in ("exemplo", "ficha", "plataforma", "medicao", "relat")):
                        caminhos_sugeridos.insert(0, os.path.join(pasta_downloads, f))
                        break
            except Exception:
                pass

        for p in caminhos_sugeridos:
            if os.path.exists(p):
                self.entry_arquivo.delete(0, "end")
                self.entry_arquivo.insert(0, os.path.normpath(p))
                self._log(f"📁 Arquivo detectado pronto para uso: {os.path.basename(p)}")
                break

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Dados Brutos da Plataforma",
            filetypes=[
                ("Planilhas Excel e CSV", "*.xlsx *.xlsm *.xls *.csv"),
                ("Pastas de Trabalho Excel (*.xlsx)", "*.xlsx"),
                ("Planilhas com Macro (*.xlsm)", "*.xlsm"),
                ("Arquivos CSV (*.csv)", "*.csv"),
                ("Todos os Arquivos", "*.*")
            ]
        )
        if caminho:
            self.entry_arquivo.delete(0, "end")
            self.entry_arquivo.insert(0, os.path.normpath(caminho))
            self._log(f"📂 Arquivo selecionado: {os.path.basename(caminho)}")

    def _obter_caminho_valido(self):
        caminho = self.entry_arquivo.get().strip()
        if not caminho or not os.path.exists(caminho):
            self._verificar_arquivo_padrao()
            caminho = self.entry_arquivo.get().strip()
        return caminho

    def _mandar_pro_excel_direto(self):
        """Compila e abre diretamente o relatório no Excel em 1 clique."""
        if self.executando:
            return

        caminho_bruto = self._obter_caminho_valido()
        if not caminho_bruto or not os.path.exists(caminho_bruto):
            messagebox.showwarning("Arquivo não selecionado", "Por favor, selecione um arquivo de dados brutos clicando em '📂 Procurar'.")
            return

        self.executando = True
        self.btn_manda_pro_excel.configure(state="disabled", text="⏳ Gerando Excel...")
        self.lbl_status_local.configure(text="Compilando e gerando relatório completo...")
        self._log(f"⚡ Mandando pro Excel com base em: {os.path.basename(caminho_bruto)}...")

        threading.Thread(target=self._executar_manda_pro_excel, args=(caminho_bruto,), daemon=True).start()

    def _executar_manda_pro_excel(self, caminho_bruto):
        try:
            if compilador_plataforma is None:
                raise ImportError("Módulo compilador_plataforma não encontrado.")

            # Processa os dados
            df_bruto = compilador_plataforma.carregar_dados_brutos(caminho_bruto)
            resultado = compilador_plataforma.processar_fichas(df_bruto)
            self.resultado_atual = resultado

            # Define caminho de saída na Área de Trabalho com timestamp
            nome_arquivo = f"Relatorio_Plataforma_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
            caminho_destino = os.path.join(os.path.expanduser("~"), "Desktop", nome_arquivo)

            # Gera a planilha usando o template oficial completo
            saida_final = compilador_plataforma.exportar_relatorio_excel(resultado, caminho_destino)
            self.caminho_ultimo_excel = saida_final

            # Atualiza a tela
            self.parent_frame.after(0, lambda: self._atualizar_ui_sucesso(resultado, caminho_bruto))

            # Abre imediatamente no Excel
            self.parent_frame.after(100, lambda: self._abrir_arquivo_excel(saida_final))

        except Exception as e:
            log_error(f"Erro ao mandar pro Excel: {e}")
            self.parent_frame.after(0, lambda erro=str(e): self._atualizar_ui_erro(erro))

    def _abrir_arquivo_excel(self, caminho):
        try:
            os.startfile(caminho)
            self._log(f"🚀 Relatório oficial aberto no Excel com sucesso:\n   {os.path.basename(caminho)}")
            show_toast(self.container, "Relatório aberto no Excel com sucesso!", duration=3000, type="success")
            play_sound("done")
        except Exception as e:
            try:
                subprocess.Popen(["cmd", "/c", "start", "", caminho], shell=True)
                self._log(f"🚀 Relatório aberto no Excel via comando:\n   {os.path.basename(caminho)}")
            except Exception as e2:
                self._log(f"⚠️ Não foi possível abrir o Excel automaticamente: {e2}")
                messagebox.showinfo("Arquivo Gerado", f"O relatório foi gerado com sucesso em:\n{caminho}\n\nAbra-o diretamente na sua Área de Trabalho.")

    def _iniciar_processamento_async(self):
        if self.executando:
            return
        caminho = self._obter_caminho_valido()
        if not caminho or not os.path.exists(caminho):
            messagebox.showwarning("Arquivo Inválido", "Por favor, selecione um arquivo válido de dados brutos (.xlsx, .xlsm ou .csv).")
            return

        self.executando = True
        self.btn_processar.configure(state="disabled", text="⏳ Processando...")
        self.lbl_status_local.configure(text="Lendo e compilando dados brutos...")
        self._log(f"⏳ Compilando dados em tela: {os.path.basename(caminho)}...")

        threading.Thread(target=self._executar_processamento, args=(caminho,), daemon=True).start()

    def _executar_processamento(self, caminho):
        try:
            if compilador_plataforma is None:
                raise ImportError("Módulo compilador_plataforma não foi carregado.")

            df_bruto = compilador_plataforma.carregar_dados_brutos(caminho)
            resultado = compilador_plataforma.processar_fichas(df_bruto)
            self.resultado_atual = resultado

            self.parent_frame.after(0, lambda: self._atualizar_ui_sucesso(resultado, caminho))
        except Exception as e:
            log_error(f"Erro ao processar dados da plataforma: {e}")
            self.parent_frame.after(0, lambda erro=str(e): self._atualizar_ui_erro(erro))

    def _atualizar_ui_sucesso(self, resultado, caminho):
        self.executando = False
        self.btn_processar.configure(state="normal", text="⚡ Compilar em Tela")
        self.btn_manda_pro_excel.configure(state="normal", text="🚀 Mandar pro Excel (Gerar & Abrir)")
        self.lbl_status_local.configure(text="Compilação concluída com sucesso!")

        metricas = resultado["metricas"]
        df_resumo = resultado["df_resumo"]

        self.card_total["lbl_v"].configure(text=f"{metricas['total_fichas']}")
        self.card_total["lbl_s"].configure(text=f"{metricas['total_finalizadas']} finalizadas")

        pct_m = (metricas['finalizadas_manha'] / metricas['total_fichas'] * 100) if metricas['total_fichas'] else 0
        self.card_manha["lbl_v"].configure(text=f"{metricas['finalizadas_manha']}")
        self.card_manha["lbl_s"].configure(text=f"{pct_m:.0f}% das fichas")

        pct_t = (metricas['finalizadas_tarde'] / metricas['total_fichas'] * 100) if metricas['total_fichas'] else 0
        self.card_tarde["lbl_v"].configure(text=f"{metricas['finalizadas_tarde']}")
        self.card_tarde["lbl_s"].configure(text=f"{pct_t:.0f}% das fichas")

        self.card_rapidas["lbl_v"].configure(text=f"{metricas['rapidas_manha']}")
        self.card_rapidas["lbl_s"].configure(text=f"Média: {metricas['tempo_medio_rapidas']} min")

        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, r in df_resumo.iterrows():
            seg = str(r["SEGMENTO"]).strip()
            is_total = (seg == "TOTAL")
            tot = int(r["TOTAL FICHAS"])
            m = int(r["ENCERRADAS MANHA"])
            pm = f"{r['% MANHA'] * 100:.0f}%"
            t = int(r["ENCERRADAS A TARDE"])
            pt = f"{r['% TARDE'] * 100:.0f}%"

            tag = "total_row" if is_total else ("even_row" if idx % 2 == 0 else "odd_row")
            self.tree.insert("", "end", values=(seg, tot, m, pm, t, pt), tags=(tag,))

        self._log(f"✅ {metricas['total_fichas']} fichas consolidadas: Manhã: {metricas['finalizadas_manha']} ({pct_m:.0f}%) | Tarde: {metricas['finalizadas_tarde']} ({pct_t:.0f}%)")

    def _atualizar_ui_erro(self, erro_msg):
        self.executando = False
        self.btn_processar.configure(state="normal", text="⚡ Compilar em Tela")
        self.btn_manda_pro_excel.configure(state="normal", text="🚀 Mandar pro Excel (Gerar & Abrir)")
        self.lbl_status_local.configure(text="Erro durante o processamento.")
        self._log(f"❌ Erro: {erro_msg}")
        messagebox.showerror("Erro de Processamento", f"Não foi possível processar o relatório:\n\n{erro_msg}")
        show_toast(self.container, "Erro ao processar dados.", duration=3000, type="error")

    def _exportar_excel(self):
        """Salva uma cópia do relatório em local escolhido pelo usuário."""
        caminho_bruto = self._obter_caminho_valido()
        if not self.resultado_atual:
            if caminho_bruto and os.path.exists(caminho_bruto):
                try:
                    df = compilador_plataforma.carregar_dados_brutos(caminho_bruto)
                    self.resultado_atual = compilador_plataforma.processar_fichas(df)
                    self._atualizar_ui_sucesso(self.resultado_atual, caminho_bruto)
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível compilar os dados: {e}")
                    return
            else:
                messagebox.showinfo("Sem Dados", "Selecione o arquivo de dados brutos primeiro.")
                return

        caminho_sugerido = os.path.join(
            os.path.expanduser("~"), "Desktop",
            f"Relatorio_Plataforma_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
        )

        caminho_salvar = filedialog.asksaveasfilename(
            title="Salvar Relatório Excel Oficial",
            defaultextension=".xlsm",
            initialfile=os.path.basename(caminho_sugerido),
            filetypes=[
                ("Planilha Habilitada para Macro (*.xlsm)", "*.xlsm"),
                ("Pasta de Trabalho do Excel (*.xlsx)", "*.xlsx")
            ]
        )

        if not caminho_salvar:
            return

        try:
            saida = compilador_plataforma.exportar_relatorio_excel(self.resultado_atual, caminho_salvar)
            self.caminho_ultimo_excel = saida
            self._log(f"💾 Relatório oficial salvo com sucesso em:\n   {saida}")
            show_toast(self.container, "Excel oficial gerado com sucesso!", duration=3000, type="success")
            play_sound("done")
        except Exception as e:
            log_error(f"Erro ao exportar Excel: {e}")
            messagebox.showerror("Erro ao Exportar", f"Ocorreu um erro ao salvar o arquivo Excel:\n\n{e}")

    def _copiar_para_clipboard(self):
        caminho_bruto = self._obter_caminho_valido()
        if not self.resultado_atual:
            if caminho_bruto and os.path.exists(caminho_bruto):
                try:
                    df = compilador_plataforma.carregar_dados_brutos(caminho_bruto)
                    self.resultado_atual = compilador_plataforma.processar_fichas(df)
                    self._atualizar_ui_sucesso(self.resultado_atual, caminho_bruto)
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível compilar os dados: {e}")
                    return
            else:
                messagebox.showinfo("Sem Dados", "Selecione o arquivo de dados brutos primeiro.")
                return

        try:
            html = compilador_plataforma.gerar_html_clipboard(self.resultado_atual)
            texto = compilador_plataforma.gerar_texto_clipboard(self.resultado_atual)

            copiado = False
            if win32clipboard:
                try:
                    CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
                    header = (
                        "Version:0.9\r\n"
                        "StartHTML:00000071\r\n"
                        f"EndHTML:{71 + len(html):08d}\r\n"
                        "StartFragment:00000071\r\n"
                        f"EndFragment:{71 + len(html):08d}\r\n"
                    )
                    payload = (header + html).encode("utf-8")
                    win32clipboard.OpenClipboard(0)
                    try:
                        win32clipboard.EmptyClipboard()
                        win32clipboard.SetClipboardData(CF_HTML, payload)
                        win32clipboard.SetClipboardText(texto)
                        copiado = True
                    finally:
                        win32clipboard.CloseClipboard()
                except Exception as e:
                    log_error(f"Erro win32clipboard: {e}")

            if not copiado:
                self.parent_frame.clipboard_clear()
                self.parent_frame.clipboard_append(texto)
                self.parent_frame.update()

            self._log("📋 Tabela formatada copiada para a Área de Transferência!")
            show_toast(self.container, "Tabela copiada! Cole direto no Outlook/Teams (Ctrl+V)", duration=3500, type="success")
            play_sound("done")
        except Exception as e:
            messagebox.showerror("Erro ao Copiar", f"Não foi possível copiar para a área de transferência: {e}")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        palette = obter_palette_dict(tema_nome, self.paletas)

        self.container.configure(fg_color=palette["surface"])
        self.frame_top_bar.configure(fg_color=palette["surface_alt"])
        self.frame_central.configure(fg_color=palette["surface_alt"])
        self.frame_terminal_container.configure(fg_color=palette["surface_alt"])

        self.lbl_titulo.configure(text_color=palette["text"])
        self.lbl_desc.configure(text_color=palette["muted"])
        self.lbl_origem.configure(text_color=palette["text"])
        self.lbl_status_local.configure(text_color=palette["muted"])

        self.entry_arquivo.configure(
            fg_color=palette["blue_pale"] if "Light" in tema_nome else c.get("input", "#1f3a60"),
            text_color=palette["text"],
            border_color=palette["border"]
        )

        self.btn_procurar.configure(
            fg_color=palette["navy_soft"],
            hover_color=palette["blue_hover"],
            text_color="#ffffff"
        )

        self.btn_abrir_pasta_data.configure(
            fg_color=palette["navy_soft"],
            hover_color=palette["blue_hover"],
            text_color="#ffffff"
        )

        self.btn_processar.configure(
            fg_color=c.get("accent", "#f57c00"),
            hover_color=c.get("hover", "#e65100"),
            text_color="#ffffff"
        )

        self.btn_manda_pro_excel.configure(
            fg_color="#107C41",
            hover_color="#0B5A2F",
            text_color="#ffffff"
        )

        self.btn_copiar.configure(
            fg_color=palette["blue"],
            hover_color=palette["blue_hover"],
            text_color="#ffffff"
        )

        self.btn_salvar_como.configure(
            fg_color=palette["navy_soft"],
            hover_color=palette["blue_hover"],
            text_color="#ffffff"
        )

        for card in [self.card_total, self.card_manha, self.card_tarde, self.card_rapidas]:
            card["frame"].configure(fg_color=palette["blue_pale"] if "Light" in tema_nome else c.get("input", "#1f3a60"))
            card["lbl_v"].configure(text_color=palette["text"])

        self.txt_terminal.configure(
            fg_color=palette["terminal"],
            text_color=palette["terminal_text"]
        )

        style = ttk.Style()
        style.configure(
            "Treeview.Heading",
            background="#0677D9",
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview.Heading", background=[("active", "#005BA1")])


def construir_aba_plataforma(parent_frame, lbl_status=None, paletas=None, obter_tema_atual=None):
    return AbaRelatorioPlataforma(parent_frame, lbl_status, paletas, obter_tema_atual)
