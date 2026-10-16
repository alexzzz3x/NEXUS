# -*- coding: utf-8 -*-
"""
utils_spotlight.py — Spotlight / Command Palette Universal (Ctrl + K) para o N.E.X.U.S.

Permite que o operador execute ações instantâneas de qualquer lugar do sistema:
- Digitar equipe (ex: CSC300) ➔ Saída OSME, Contagem, Bandeirada, Atualizar OSME
- Digitar serial (ex: 18398041) ➔ Consulta IQ09, Saída OSME, Etiqueta
- Digitar nome de módulo (ex: etiquetas, migo, zm02) ➔ Navegação instantânea
"""

import re
import tkinter as tk
import customtkinter as ctk


class SpotlightModal:
    _instancia = None

    @classmethod
    def abrir(cls, root, paleta, on_navigate, on_exec_acao=None):
        if cls._instancia and cls._instancia.winfo_exists():
            cls._instancia.lift()
            cls._instancia.entry_busca.focus_set()
            return cls._instancia
        cls._instancia = cls(root, paleta, on_navigate, on_exec_acao)
        return cls._instancia

    def __init__(self, root, paleta, on_navigate, on_exec_acao=None):
        self.root = root
        self.paleta = paleta or {}
        self.on_navigate = on_navigate
        self.on_exec_acao = on_exec_acao

        self.janela = ctk.CTkToplevel(root)
        self.janela.title("N.E.X.U.S Spotlight")
        self.janela.overrideredirect(True)
        self.janela.attributes("-topmost", True)

        largura, altura = 580, 420
        x = root.winfo_rootx() + (root.winfo_width() - largura) // 2
        y = root.winfo_rooty() + max(40, (root.winfo_height() - altura) // 3)
        self.janela.geometry(f"{largura}x{altura}+{x}+{y}")

        bg = self.paleta.get("bg", "#090d16")
        container = self.paleta.get("container", "#111827")
        accent = self.paleta.get("accent", "#00d2ff")
        texto = self.paleta.get("texto", "#f8fafc")

        self.card_main = ctk.CTkFrame(
            self.janela,
            corner_radius=12,
            fg_color=container,
            border_width=2,
            border_color=accent
        )
        self.card_main.pack(fill="both", expand=True)

        # Topo com Campo de Busca
        f_top = ctk.CTkFrame(self.card_main, fg_color="transparent")
        f_top.pack(fill="x", padx=14, pady=(12, 6))

        self.lbl_ico = ctk.CTkLabel(f_top, text="⚡", font=("Segoe UI", 16))
        self.lbl_ico.pack(side="left", padx=(4, 8))

        self.entry_busca = ctk.CTkEntry(
            f_top,
            placeholder_text="Digite uma equipe (ex: CSC300), serial, módulo ou ação...",
            font=("Segoe UI", 13),
            height=38,
            corner_radius=8
        )
        self.entry_busca.pack(side="left", fill="x", expand=True)
        self.entry_busca.bind("<KeyRelease>", self._ao_digitar)
        self.entry_busca.bind("<Return>", self._ao_pressionar_enter)
        self.entry_busca.bind("<Down>", lambda e: self._mover_selecao(1))
        self.entry_busca.bind("<Up>", lambda e: self._mover_selecao(-1))

        self.btn_fechar = ctk.CTkButton(
            f_top,
            text="✕",
            width=32,
            height=32,
            font=("Segoe UI", 12, "bold"),
            fg_color="transparent",
            text_color="#94a3b8",
            command=self.fechar
        )
        self.btn_fechar.pack(side="right", padx=(8, 0))

        # Divisor sutil
        div = ctk.CTkFrame(self.card_main, height=1, fg_color="#1e293b")
        div.pack(fill="x", padx=10, pady=(2, 6))

        # Lista de Ações / Resultados
        self.scroll_resultados = ctk.CTkScrollableFrame(self.card_main, corner_radius=6, fg_color="transparent")
        self.scroll_resultados.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # Rodapé com dicas de atalhos
        f_bot = ctk.CTkFrame(self.card_main, height=26, fg_color="transparent")
        f_bot.pack(fill="x", padx=14, pady=(0, 8))

        lbl_dica = ctk.CTkLabel(
            f_bot,
            text="[Enter] Executar   •   [↑/↓] Navegar   •   [Esc] Fechar",
            font=("Segoe UI", 9),
            text_color="#64748b"
        )
        lbl_dica.pack(side="left")

        self.itens_renderizados = []
        self.indice_selecionado = 0

        # Eventos globais da janela
        self.janela.bind("<Escape>", lambda e: self.fechar())
        self.janela.bind("<FocusOut>", self._ao_perder_foco)

        # Carrega itens iniciais e foca
        self.janela.after(50, self.entry_busca.focus_set)
        self._ao_digitar()

    def _ao_perder_foco(self, event):
        # Fecha se o clique foi fora da janela
        try:
            x, y = self.janela.winfo_pointerxy()
            wx = self.janela.winfo_rootx()
            wy = self.janela.winfo_rooty()
            ww = self.janela.winfo_width()
            wh = self.janela.winfo_height()
            if not (wx <= x <= wx + ww and wy <= y <= wy + wh):
                self.fechar()
        except Exception:
            pass

    def fechar(self):
        try:
            self.janela.destroy()
        except Exception:
            pass
        SpotlightModal._instancia = None

    def _ao_digitar(self, event=None):
        termo = self.entry_busca.get().strip().upper()
        self.itens_renderizados.clear()
        self.indice_selecionado = 0

        for w in self.scroll_resultados.winfo_children():
            w.destroy()

        acoes = self._gerar_acoes_para_termo(termo)
        if not acoes:
            lbl_vazio = ctk.CTkLabel(
                self.scroll_resultados,
                text="Nenhuma ação correspondente encontrada.",
                font=("Segoe UI", 11, "italic"),
                text_color="#64748b"
            )
            lbl_vazio.pack(pady=30)
            return

        card_bg = self.paleta.get("card_bg", self.paleta.get("container", "#111827"))

        for i, a in enumerate(acoes):
            row = ctk.CTkFrame(self.scroll_resultados, height=44, corner_radius=6, fg_color=card_bg, cursor="hand2")
            row.pack(fill="x", pady=2, padx=2)
            row.pack_propagate(False)

            lbl_ic = ctk.CTkLabel(row, text=f"  {a['icone']}", font=("Segoe UI", 14), width=34)
            lbl_ic.pack(side="left", padx=(2, 4))

            f_txt = ctk.CTkFrame(row, fg_color="transparent")
            f_txt.pack(side="left", fill="both", expand=True, pady=4)

            lbl_tit = ctk.CTkLabel(f_txt, text=a["titulo"], font=("Segoe UI", 11, "bold"), anchor="w")
            lbl_tit.pack(fill="x")

            lbl_sub = ctk.CTkLabel(f_txt, text=a["subtitulo"], font=("Segoe UI", 9), text_color="#94a3b8", anchor="w")
            lbl_sub.pack(fill="x")

            lbl_tag = ctk.CTkLabel(row, text=f" {a.get('tag', 'Ação')} ", font=("Segoe UI", 9, "bold"), text_color="#38bdf8")
            lbl_tag.pack(side="right", padx=10)

            # Binds de clique
            acao_cmd = a["comando"]
            for w in (row, lbl_ic, f_txt, lbl_tit, lbl_sub, lbl_tag):
                w.bind("<Button-1>", lambda e, cmd=acao_cmd: self._executar(cmd))
                w.bind("<Enter>", lambda e, idx=i: self._definir_selecao(idx))

            self.itens_renderizados.append({"frame": row, "comando": acao_cmd, "lbl_tit": lbl_tit})

        self._atualizar_destaque_selecao()

    def _gerar_acoes_para_termo(self, termo):
        acoes = []
        if not termo:
            # Sugestões rápidas
            acoes.extend([
                {"icone": "📢", "titulo": "O que há de novo?", "subtitulo": "Notas de versão e novas funcionalidades do sistema", "tag": "Novidades", "comando": ("novidades", None)},
                {"icone": "📤", "titulo": "Nova Saída OSME", "subtitulo": "Lançamento de medidores e materiais no SAP MIGO Z61", "tag": "Módulo", "comando": ("nav", "📤 Saída OSME")},
                {"icone": "🚗", "titulo": "Contagem de Carro", "subtitulo": "Balanço físico e auditoria de veículos no SAP ZMMCSE001", "tag": "Módulo", "comando": ("nav", "🚗 Contagem de Carro")},
                {"icone": "🔄", "titulo": "Atualizar OSME", "subtitulo": "Varredura na IW39 e renovação de ordens vencidas", "tag": "Módulo", "comando": ("nav", "🔄 Atualizar OSME")},
                {"icone": "🔢", "titulo": "Extrator & IQ09", "subtitulo": "Pesquisa de medidores e status no SAP IQ09", "tag": "Módulo", "comando": ("nav", "🔢 Extrator Seriais")},
                {"icone": "🎫", "titulo": "Gerador de Etiquetas", "subtitulo": "Geração de etiquetas em PDF com QR Code", "tag": "Módulo", "comando": ("nav", "🎫 Etiquetas")},
                {"icone": "🔍", "titulo": "Consultar MIGO", "subtitulo": "Consulta de equipes clandestinas (RCL001..RCL008)", "tag": "Módulo", "comando": ("nav", "🔍 Consultar MIGO")},
            ])
            return acoes

        # 1. Padrão Equipe (ex: CSC300, CVC312, RCL001, etc.)
        match_eq = re.search(r"\b([A-Z]{3}\d{3}|RCL\d{3})\b", termo)
        if match_eq or len(termo) >= 5 and termo[:3].isalpha() and termo[3:].isdigit():
            eq = match_eq.group(1) if match_eq else termo
            acoes.append({"icone": "📤", "titulo": f"Fazer Saída OSME para {eq}", "subtitulo": f"Preenche o recurso {eq} e prepara lançamento Z61", "tag": "Equipe", "comando": ("saida_eq", eq)})
            acoes.append({"icone": "🚗", "titulo": f"Fazer Contagem de Carro para {eq}", "subtitulo": f"Audita estoque físico de {eq} na ZMMCSE001", "tag": "Equipe", "comando": ("contagem_eq", eq)})
            acoes.append({"icone": "🚩", "titulo": f"Executar Bandeirada para {eq}", "subtitulo": f"Dispara bandeirada no SAP para a equipe {eq}", "tag": "Equipe", "comando": ("bandeirada_eq", eq)})
            acoes.append({"icone": "🔄", "titulo": f"Atualizar OSME de {eq} no SAP", "subtitulo": f"Consulta ordem mais recente na IW39", "tag": "Equipe", "comando": ("nav", "🔄 Atualizar OSME")})

        # 2. Padrão Serial (Numérico com 6+ dígitos ou alfanumérico UQM...)
        if re.search(r"\b\d{6,14}\b", termo) or re.search(r"\bUQM[A-Z0-9]+\b", termo):
            serial = termo.strip()
            acoes.append({"icone": "🔍", "titulo": f"Pesquisar medidor {serial} no SAP IQ09", "subtitulo": "Consulta lote, depósito e status no SAP", "tag": "Serial", "comando": ("iq09_serial", serial)})
            acoes.append({"icone": "📤", "titulo": f"Lançar serial {serial} na Saída OSME", "subtitulo": "Carrega o medidor diretamente para baixa", "tag": "Serial", "comando": ("saida_serial", serial)})
            acoes.append({"icone": "🎫", "titulo": f"Gerar etiqueta para serial {serial}", "subtitulo": "Abre o módulo de etiquetas", "tag": "Serial", "comando": ("nav", "🎫 Etiquetas")})

        # 3. Padrão Módulos por Palavra-Chave
        modulos_catalogo = [
            ("🏠 Início", ["INICIO", "HOME", "DASHBOARD", "PAINEL"]),
            ("📤 Saída OSME", ["SAIDA", "OSME", "MIGO", "Z61", "FORNECIDOS", "BAIXA"]),
            ("🚗 Contagem de Carro", ["CONTAGEM", "CARRO", "VEICULO", "BALANCO", "ZMMCSE001"]),
            ("🔄 Atualizar OSME", ["ATUALIZAR", "VENCIMENTO", "IW39", "VENCIDA"]),
            ("🚩 BANDEIRADA", ["BANDEIRADA", "BANDEIRA", "LIBERAR"]),
            ("🔑 Login SAP", ["LOGIN", "SAP", "CONECTAR", "SENHA"]),
            ("📦 Transf. OSME", ["TRANSF", "TRANSFERENCIA", "Z62"]),
            ("🏢 Contagem Depósito", ["DEPOSITO", "ZM29", "ESTOQUE", "LOTE 0", "ALMOX"]),
            ("🔍 Consultar MIGO", ["CLANDESTINA", "RCL", "CONSULTAR MIGO"]),
            ("⚙️ Entrada ZM02", ["ZM02", "ENTRADA", "RECEBIMENTO"]),
            ("🔄 Mudança Recurso", ["MUDANCA", "RECURSO", "TRANSFERIR EQUIPE"]),
            ("📦 Conferencia", ["CONFERENCIA", "BIPAGEM", "COMPARAR"]),
            ("📋 Fichas", ["FICHAS", "WORKFORCE", "RE"]),
            ("🎫 Etiquetas", ["ETIQUETA", "IMPRIMIR", "QR", "PDF"]),
            ("🔢 Extrator Seriais", ["EXTRATOR", "SERIAL", "IQ09", "EXPANDIR"]),
            ("📈 Relatório Plataforma", ["PLATAFORMA", "RELATORIO", "EXCEL"]),
        ]

        for tab_nome, keywords in modulos_catalogo:
            if any(k in termo for k in keywords) or termo in tab_nome.upper():
                # Evita duplicatas se já adicionado
                if not any(a["comando"] == ("nav", tab_nome) for a in acoes):
                    acoes.append({
                        "icone": tab_nome.split(" ")[0],
                        "titulo": f"Abrir módulo: {tab_nome}",
                        "subtitulo": f"Navegar instantaneamente para a aba {tab_nome}",
                        "tag": "Navegar",
                        "comando": ("nav", tab_nome)
                    })

        # 4. Busca por Novidades / Changelog
        termos_novidades = ["NOVIDADE", "NOVO", "ATT", "VERSAO", "VERSÃO", "UPDATE", "PATCH", "CHANGELOG", "O QUE"]
        if any(tn in termo for tn in termos_novidades):
            acoes.insert(0, {
                "icone": "📢",
                "titulo": "O que há de novo no N.E.X.U.S",
                "subtitulo": "Consulta as últimas notas de versão e novas funcionalidades",
                "tag": "Novidades",
                "comando": ("novidades", None)
            })

        return acoes

    def _mover_selecao(self, delta):
        if not self.itens_renderizados:
            return
        self.indice_selecionado = (self.indice_selecionado + delta) % len(self.itens_renderizados)
        self._atualizar_destaque_selecao()

    def _definir_selecao(self, idx):
        self.indice_selecionado = idx
        self._atualizar_destaque_selecao()

    def _atualizar_destaque_selecao(self):
        accent = self.paleta.get("accent", "#00d2ff")
        card_bg = self.paleta.get("card_bg", self.paleta.get("container", "#111827"))
        hover_bg = self.paleta.get("input", "#1e293b")

        for i, item in enumerate(self.itens_renderizados):
            if i == self.indice_selecionado:
                item["frame"].configure(fg_color=hover_bg, border_width=1, border_color=accent)
            else:
                item["frame"].configure(fg_color=card_bg, border_width=0)

    def _ao_pressionar_enter(self, event=None):
        if self.itens_renderizados and 0 <= self.indice_selecionado < len(self.itens_renderizados):
            cmd = self.itens_renderizados[self.indice_selecionado]["comando"]
            self._executar(cmd)

    def _executar(self, cmd):
        tipo, valor = cmd
        self.fechar()

        if tipo == "nav":
            if self.on_navigate:
                self.on_navigate(valor)
        elif tipo == "novidades":
            try:
                from utils_updater import exibir_novidades
                exibir_novidades(self.root, self.paleta)
            except Exception as e:
                print(f"[Spotlight] Erro ao abrir novidades: {e}")
        elif tipo in ("saida_eq", "contagem_eq", "bandeirada_eq", "iq09_serial", "saida_serial"):
            if self.on_exec_acao:
                self.on_exec_acao(tipo, valor)
            elif self.on_navigate:
                if "saida" in tipo:
                    self.on_navigate("📤 Saída OSME")
                elif "contagem" in tipo:
                    self.on_navigate("🚗 Contagem de Carro")
                elif "bandeirada" in tipo:
                    self.on_navigate("🚩 BANDEIRADA")
                elif "iq09" in tipo:
                    self.on_navigate("🔢 Extrator Seriais")
