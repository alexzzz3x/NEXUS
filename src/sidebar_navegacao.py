"""
sidebar_navegacao.py — Sidebar lateral esquerda retrátil para o N.E.X.U.S.
Componente modular e desacoplado, com categorias agrupadas (accordion),
sistema de favoritos persistente (⭐) sem arquivos soltos no Desktop,
transição instantânea e estável (zero tearing/artefatos),
modo compacto (68px) com tooltips flutuantes e modo expandido (230px).
"""

import os
import sys
import json
import tkinter as tk
import customtkinter as ctk

from utils_os import obter_caminho_config


class TooltipFlutuante:
    """Micro-janela flutuante moderna para exibir o nome do módulo no modo recolhido."""

    def __init__(self):
        self.tip_window = None

    def exibir(self, widget, texto, paleta):
        self.ocultar()
        if not texto:
            return
        try:
            x = widget.winfo_rootx() + widget.winfo_width() + 10
            y = widget.winfo_rooty() + (widget.winfo_height() // 2) - 14
        except Exception:
            return

        self.tip_window = tw = tk.Toplevel(widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)

        bg = paleta.get("container", "#172a45")
        fg = paleta.get("texto", "#ffffff")
        accent = paleta.get("accent", "#f57c00")

        frame = tk.Frame(tw, bg=bg, highlightbackground=accent, highlightthickness=1)
        frame.pack()
        lbl = tk.Label(
            frame,
            text=f" {texto} ",
            font=("Segoe UI", 10, "bold"),
            fg=fg,
            bg=bg,
            padx=8,
            pady=4
        )
        lbl.pack()

    def ocultar(self):
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


class SidebarNavegacao:
    """
    Sidebar lateral de navegação retrátil estilo ChatGPT/VS Code.
    Substitui a barra horizontal de abas, agrupando módulos em categorias lógicas,
    com suporte a Favoritos (⭐), persistência segura em AppData e transição 100% estável.
    """
    LARGURA_EXPANDIDA = 245
    LARGURA_RECOLHIDA = 68

    def __init__(self, parent, notebook, categorias, paletas, obter_tema_atual, on_tab_changed=None, use_grid=True):
        self.parent = parent
        self.notebook = notebook
        self.categorias = categorias
        self.paletas = paletas
        self.obter_tema_atual = obter_tema_atual
        self.on_tab_changed = on_tab_changed
        self.use_grid = use_grid

        self.expandido = True
        self.animando = False
        self.aba_selecionada = None
        self.tooltip = TooltipFlutuante()

        # Carrega os módulos favoritos persistidos de forma segura no AppData
        self.favoritos = self._carregar_favoritos()

        # Dicionário rápido de busca de metadados dos módulos {tab_id: (icon, nome, cat_nome)}
        self.modulos_map = {}
        for cat in self.categorias:
            for icon, nome, tab_id in cat["itens"]:
                self.modulos_map[tab_id] = (icon, nome, cat["nome"])

        # Guarda se cada grupo de accordion está aberto ou fechado (padrão: aberto)
        self.grupos_abertos = {c["nome"]: True for c in self.categorias}
        self.grupos_abertos["FAVORITOS"] = True

        self.botoes_info = []
        self.categoria_widgets = {}

        # Container lateral fixado na coluna 0 da grade (sem clipping nem corte)
        self.frame_container = ctk.CTkFrame(
            self.parent,
            width=self.LARGURA_EXPANDIDA if self.expandido else self.LARGURA_RECOLHIDA,
            corner_radius=0
        )
        self.frame_container.pack_propagate(False)
        self.frame_container.grid_propagate(False)

        if self.use_grid:
            self.frame_container.grid(row=0, column=0, sticky="nsew")
        else:
            self.frame_container.pack(side="left", fill="y")

        # Topo da sidebar com botão toggle indicativo (◀/▶)
        self.frame_topo = ctk.CTkFrame(self.frame_container, height=44, corner_radius=0, fg_color="transparent")
        self.frame_topo.pack(fill="x", padx=4, pady=(6, 2))

        self.btn_toggle = ctk.CTkButton(
            self.frame_topo,
            text="◀",
            width=34,
            height=32,
            font=("Segoe UI", 13, "bold"),
            corner_radius=8,
            command=self.alternar_expansao
        )
        self.btn_toggle.pack(side="left", padx=(6, 8), pady=4)
        self.btn_toggle.bind("<Enter>", lambda e: self.tooltip.exibir(self.btn_toggle, "Recolher menu" if self.expandido else "Expandir menu", self.paletas.get(self.obter_tema_atual(), {})))
        self.btn_toggle.bind("<Leave>", lambda e: self.tooltip.ocultar())

        self.lbl_menu_title = ctk.CTkLabel(
            self.frame_topo,
            text="NAVEGAÇÃO",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_menu_title.pack(side="left", fill="x", expand=True, pady=4)

        # Barra de Busca Rápida (Filtro instantâneo no topo)
        self.frame_busca = ctk.CTkFrame(self.frame_container, height=32, corner_radius=6, fg_color="transparent")
        self.entry_busca = ctk.CTkEntry(
            self.frame_busca,
            placeholder_text="🔍 Buscar módulo...",
            font=("Segoe UI", 11),
            height=28,
            corner_radius=6
        )
        self.entry_busca.pack(fill="x", expand=True, padx=4, pady=2)
        self.entry_busca.bind("<KeyRelease>", self._ao_filtrar_modulos)

        try:
            self.parent.winfo_toplevel().bind("<Control-f>", lambda e: self._focar_busca())
        except Exception:
            pass

        # Scrollable frame para suportar qualquer quantidade de módulos
        self.scroll = ctk.CTkScrollableFrame(self.frame_container, corner_radius=0, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=2, pady=2)

        self._construir_itens()
        self.aplicar_tema(self.obter_tema_atual())
        self._atualizar_layout()


    # ---------------------------------------------------------
    # Persistência Segura de Favoritos (AppData / data raiz)
    # ---------------------------------------------------------
    def _obter_caminho_favoritos(self):
        """Retorna o caminho seguro no AppData para nunca salvar solto no Desktop."""
        return obter_caminho_config("favoritos_sidebar.json")

    def _carregar_favoritos(self):
        """Carrega conjunto de tab_ids favoritos do arquivo JSON."""
        caminho = self._obter_caminho_favoritos()
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    if isinstance(dados, list):
                        return set(dados)
            except Exception as e:
                print(f"[Sidebar] Erro ao carregar favoritos: {e}")
        return set()

    def _salvar_favoritos(self):
        """Salva conjunto de tab_ids favoritos no arquivo JSON."""
        caminho = self._obter_caminho_favoritos()
        try:
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(list(self.favoritos), f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"[Sidebar] Erro ao salvar favoritos: {e}")

    def alternar_favorito(self, tab_id):
        """Adiciona ou remove um módulo dos favoritos e atualiza a interface."""
        if tab_id in self.favoritos:
            self.favoritos.remove(tab_id)
        else:
            self.favoritos.add(tab_id)

        self._salvar_favoritos()
        self._reconstruir_itens_favoritos()
        self._atualizar_estrelas()
        self._atualizar_layout()

    # ---------------------------------------------------------
    # Construção dos Elementos
    # ---------------------------------------------------------
    def _construir_itens(self):
        """Constrói a seção de Favoritos e as categorias normais."""
        # A. SEÇÃO DE FAVORITOS (Fixa no topo)
        self.cat_header_fav = ctk.CTkFrame(self.scroll, height=28, corner_radius=6, fg_color="transparent", cursor="hand2")
        self.lbl_arrow_fav = ctk.CTkLabel(self.cat_header_fav, text="▾", width=18, font=("Segoe UI", 11, "bold"))
        self.lbl_arrow_fav.pack(side="left", padx=(4, 2))

        self.lbl_titulo_fav = ctk.CTkLabel(self.cat_header_fav, text="⭐ FAVORITOS", font=("Segoe UI", 9, "bold"), anchor="w")
        self.lbl_titulo_fav.pack(side="left", fill="x", expand=True)

        self.cat_header_fav.bind("<Button-1>", lambda e: self._toggle_grupo("FAVORITOS"))
        self.lbl_arrow_fav.bind("<Button-1>", lambda e: self._toggle_grupo("FAVORITOS"))
        self.lbl_titulo_fav.bind("<Button-1>", lambda e: self._toggle_grupo("FAVORITOS"))

        self.frame_fav_itens = ctk.CTkFrame(self.scroll, fg_color="transparent")
        self.divider_fav = ctk.CTkFrame(self.scroll, height=1, fg_color="#334155")

        self.categoria_widgets["FAVORITOS"] = {
            "header": self.cat_header_fav,
            "arrow": self.lbl_arrow_fav,
            "titulo": self.lbl_titulo_fav,
            "itens_frame": self.frame_fav_itens
        }

        self._reconstruir_itens_favoritos()

        # B. CATEGORIAS NORMAIS
        for cat in self.categorias:
            nome_cat = cat["nome"]

            cat_header = ctk.CTkFrame(self.scroll, height=28, corner_radius=6, fg_color="transparent", cursor="hand2")
            lbl_arrow = ctk.CTkLabel(cat_header, text="▾", width=18, font=("Segoe UI", 11, "bold"))
            lbl_arrow.pack(side="left", padx=(4, 2))

            lbl_titulo = ctk.CTkLabel(cat_header, text=nome_cat, font=("Segoe UI", 9, "bold"), anchor="w")
            lbl_titulo.pack(side="left", fill="x", expand=True)

            cat_header.bind("<Button-1>", lambda e, n=nome_cat: self._toggle_grupo(n))
            lbl_arrow.bind("<Button-1>", lambda e, n=nome_cat: self._toggle_grupo(n))
            lbl_titulo.bind("<Button-1>", lambda e, n=nome_cat: self._toggle_grupo(n))

            itens_frame = ctk.CTkFrame(self.scroll, fg_color="transparent")

            self.categoria_widgets[nome_cat] = {
                "header": cat_header,
                "arrow": lbl_arrow,
                "titulo": lbl_titulo,
                "itens_frame": itens_frame
            }

            for icon, nome, tab_id in cat["itens"]:
                row = ctk.CTkFrame(itens_frame, fg_color="transparent", height=36)

                btn = ctk.CTkButton(
                    row,
                    text=f"  {icon}   {nome}",
                    anchor="w",
                    height=36,
                    corner_radius=8,
                    font=("Segoe UI", 10, "bold"),
                    command=lambda t=tab_id: self.selecionar_aba(t, trocar_notebook=True)
                )
                btn.bind("<Enter>", lambda e, b=btn, n=nome: self._on_btn_enter(b, n))
                btn.bind("<Leave>", lambda e, b=btn: self._on_btn_leave(b))
                btn.bind("<Button-3>", lambda e, t=tab_id: self._mostrar_menu_contexto(e, t))

                btn_fav = ctk.CTkButton(
                    row,
                    text="⭐" if tab_id in self.favoritos else "☆",
                    width=28,
                    height=34,
                    corner_radius=6,
                    font=("Segoe UI", 12),
                    fg_color="transparent",
                    text_color="#f59e0b" if tab_id in self.favoritos else "#64748b",
                    command=lambda t=tab_id: self.alternar_favorito(t)
                )

                self.botoes_info.append({
                    "tab_id": tab_id,
                    "btn": btn,
                    "btn_fav": btn_fav,
                    "icon": icon,
                    "nome": nome,
                    "cat": nome_cat,
                    "row": row,
                    "itens_frame": itens_frame
                })

    def _reconstruir_itens_favoritos(self):
        """Regenera os botões na seção ⭐ FAVORITOS."""
        for w in self.frame_fav_itens.winfo_children():
            w.destroy()

        self.botoes_info = [b for b in self.botoes_info if b.get("cat") != "FAVORITOS"]

        if not self.favoritos:
            lbl_vazio = ctk.CTkLabel(
                self.frame_fav_itens,
                text="(Clique na estrela ☆ para favoritar)",
                font=("Segoe UI", 9, "italic"),
                text_color="#64748b"
            )
            lbl_vazio.pack(padx=8, pady=3)
            return

        for tab_id in list(self.favoritos):
            if tab_id not in self.modulos_map:
                continue
            icon, nome, _ = self.modulos_map[tab_id]

            row = ctk.CTkFrame(self.frame_fav_itens, fg_color="transparent", height=36)

            btn = ctk.CTkButton(
                row,
                text=f"  {icon}   {nome}",
                anchor="w",
                height=36,
                corner_radius=8,
                font=("Segoe UI", 10, "bold"),
                command=lambda t=tab_id: self.selecionar_aba(t, trocar_notebook=True, expandir_categoria=False)
            )
            btn.bind("<Enter>", lambda e, b=btn, n=nome: self._on_btn_enter(b, n))
            btn.bind("<Leave>", lambda e, b=btn: self._on_btn_leave(b))
            btn.bind("<Button-3>", lambda e, t=tab_id: self._mostrar_menu_contexto(e, t))

            btn_fav = ctk.CTkButton(
                row,
                text="⭐",
                width=28,
                height=34,
                corner_radius=6,
                font=("Segoe UI", 12),
                fg_color="transparent",
                text_color="#f59e0b",
                command=lambda t=tab_id: self.alternar_favorito(t)
            )

            self.botoes_info.append({
                "tab_id": tab_id,
                "btn": btn,
                "btn_fav": btn_fav,
                "icon": icon,
                "nome": nome,
                "cat": "FAVORITOS",
                "row": row,
                "itens_frame": self.frame_fav_itens
            })

    def _atualizar_estrelas(self):
        """Sincroniza o estado visual das estrelas (⭐ ou ☆) em todos os botões."""
        for info in self.botoes_info:
            btn_fav = info.get("btn_fav")
            if btn_fav:
                tab_id = info["tab_id"]
                is_fav = tab_id in self.favoritos
                btn_fav.configure(
                    text="⭐" if is_fav else "☆",
                    text_color="#f59e0b" if is_fav else "#64748b"
                )

    def _mostrar_menu_contexto(self, event, tab_id):
        """Menu de contexto no clique com o botão direito."""
        paleta = self.paletas.get(self.obter_tema_atual(), {})
        bg = paleta.get("container", "#172a45")
        fg = paleta.get("texto", "#ffffff")

        menu = tk.Menu(self.parent, tearoff=0, bg=bg, fg=fg, activebackground=paleta.get("accent", "#f57c00"))
        is_fav = tab_id in self.favoritos
        if is_fav:
            menu.add_command(label="❌ Remover dos Favoritos", command=lambda: self.alternar_favorito(tab_id))
        else:
            menu.add_command(label="⭐ Adicionar aos Favoritos", command=lambda: self.alternar_favorito(tab_id))

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _ao_filtrar_modulos(self, event=None):
        """Filtra módulos em tempo real conforme o operador digita na busca."""
        termo = self.entry_busca.get().strip().lower()
        if not termo:
            self._atualizar_layout()
            return

        cats_com_match = set()

        for info in self.botoes_info:
            nome = info["nome"].lower()
            cat = info["cat"]
            corresponde = (termo in nome) or (termo in cat.lower()) or (termo in info["tab_id"].lower())
            if corresponde:
                info["row"].pack(fill="x", padx=4, pady=1)
                cats_com_match.add(cat)
            else:
                info["row"].pack_forget()

        # Atualiza a visibilidade das seções na ordem correta
        if "FAVORITOS" in cats_com_match:
            self.cat_header_fav.pack(fill="x", padx=4, pady=(4, 2))
            self.frame_fav_itens.pack(fill="x", padx=2, pady=(0, 4))
            self.divider_fav.pack(fill="x", padx=8, pady=(4, 6))
        else:
            self.cat_header_fav.pack_forget()
            self.frame_fav_itens.pack_forget()
            self.divider_fav.pack_forget()

        for cat in self.categorias:
            nome_cat = cat["nome"]
            cw = self.categoria_widgets.get(nome_cat)
            if not cw:
                continue
            if nome_cat in cats_com_match:
                cw["header"].pack(fill="x", padx=4, pady=(6, 2))
                cw["itens_frame"].pack(fill="x", padx=2, pady=(0, 2))
            else:
                cw["header"].pack_forget()
                cw["itens_frame"].pack_forget()

    def _focar_busca(self):
        """Atalho de teclado (Ctrl+F) para focar no campo de busca da sidebar."""
        if not self.expandido:
            self.alternar_expansao(animar=True)
        try:
            self.entry_busca.focus_set()
        except Exception:
            pass

    def _toggle_grupo(self, nome_cat):
        """Alterna a visibilidade dos itens de uma categoria no modo expandido."""
        if not self.expandido:
            return
        self.grupos_abertos[nome_cat] = not self.grupos_abertos.get(nome_cat, True)
        self._atualizar_layout()

    # ---------------------------------------------------------
    # ---------------------------------------------------------
    # Alternância de Expansão com Animação Suave (Ease-Out)
    # ---------------------------------------------------------
    def _animar_largura(self, alvo):
        """Interpolação de largura suave (Ease-Out) sem engasgos."""
        self.animando = True
        try:
            atual = self.frame_container.winfo_width()
        except Exception:
            self.animando = False
            return

        diff = alvo - atual
        if abs(diff) <= 2:
            self.frame_container.configure(width=alvo)
            self.animando = False
            return

        step = int(diff * 0.45)
        if step == 0:
            step = 2 if diff > 0 else -2
        nova = atual + step
        self.frame_container.configure(width=nova)
        self.parent.after(14, lambda: self._animar_largura(alvo))

    def alternar_expansao(self, animar=False):
        """Alterna instantaneamente e com estabilidade total entre os modos Expandido (245px) e Recolhido (68px)."""
        if self.animando:
            return
        self.tooltip.ocultar()
        self.expandido = not self.expandido

        alvo = self.LARGURA_EXPANDIDA if self.expandido else self.LARGURA_RECOLHIDA
        self.frame_container.configure(width=alvo)
        self._atualizar_layout()

    def _atualizar_layout(self):
        """Ajusta dimensões, textos e disposição dos componentes conforme o estado."""
        paleta = self.paletas.get(self.obter_tema_atual(), {})
        bg_container = paleta.get("container", "#111827")

        # 1. Desempacota todos os blocos do scroll para garantir ordem estrita (Header -> Itens)
        self.cat_header_fav.pack_forget()
        self.frame_fav_itens.pack_forget()
        self.divider_fav.pack_forget()
        for cw in self.categoria_widgets.values():
            cw["header"].pack_forget()
            cw["itens_frame"].pack_forget()

        if self.expandido:
            self.frame_container.configure(width=self.LARGURA_EXPANDIDA)
            self.btn_toggle.pack_forget()
            self.lbl_menu_title.pack_forget()

            self.btn_toggle.pack(side="left", padx=(6, 8), pady=4)
            self.lbl_menu_title.pack(side="left", fill="x", expand=True, pady=4)
            self.btn_toggle.configure(text="◀")

            if hasattr(self, "frame_busca"):
                self.frame_busca.pack(fill="x", padx=6, pady=(0, 4))

            # Seção de Favoritos
            self.cat_header_fav.pack(fill="x", padx=4, pady=(4, 2))
            esta_aberto_fav = self.grupos_abertos.get("FAVORITOS", True)
            self.lbl_arrow_fav.configure(text="▾" if esta_aberto_fav else "▸")
            if esta_aberto_fav:
                self.frame_fav_itens.pack(fill="x", padx=2, pady=(0, 4))
            self.divider_fav.pack(fill="x", padx=8, pady=(4, 6))

            # Categorias Normais (na ordem original de self.categorias)
            for cat in self.categorias:
                nome_cat = cat["nome"]
                cw = self.categoria_widgets.get(nome_cat)
                if not cw:
                    continue
                cw["header"].pack(fill="x", padx=4, pady=(6, 2))
                esta_aberto = self.grupos_abertos.get(nome_cat, True)
                cw["arrow"].configure(text="▾" if esta_aberto else "▸")
                if esta_aberto:
                    cw["itens_frame"].pack(fill="x", padx=2, pady=(0, 2))

            # Botões com texto completo e estrela à direita
            for info in self.botoes_info:
                btn = info["btn"]
                btn.pack_forget()

                btn_fav = info.get("btn_fav")
                if btn_fav:
                    btn_fav.pack_forget()

                row = info["row"]
                row.pack_forget()
                row.configure(fg_color=bg_container, height=36)
                row.pack_propagate(True)
                row.pack(fill="x", padx=4, pady=1)

                btn.configure(
                    text=f"  {info['icon']}   {info['nome']}",
                    anchor="w",
                    font=("Segoe UI", 10, "bold")
                )
                if btn_fav:
                    btn_fav.pack(side="right", padx=(0, 2))
                btn.pack(side="left", fill="x", expand=True, padx=(2, 2))

        else:
            self.frame_container.configure(width=self.LARGURA_RECOLHIDA)
            self.lbl_menu_title.pack_forget()
            if hasattr(self, "frame_busca"):
                self.frame_busca.pack_forget()

            self.btn_toggle.pack_forget()
            self.btn_toggle.pack(side="top", padx=4, pady=4)
            self.btn_toggle.configure(text="▶")

            # No modo recolhido, cabeçalhos de grupos ficam ocultos
            self.frame_fav_itens.pack(fill="x", padx=2, pady=1)
            self.divider_fav.pack(fill="x", padx=6, pady=4)

            for cat in self.categorias:
                nome_cat = cat["nome"]
                cw = self.categoria_widgets.get(nome_cat)
                if cw:
                    cw["itens_frame"].pack(fill="x", padx=2, pady=1)

            # No modo recolhido, centraliza os botoes com apenas o icone
            for info in self.botoes_info:
                btn_fav = info.get("btn_fav")
                if btn_fav:
                    btn_fav.pack_forget()

                btn = info["btn"]
                btn.pack_forget()

                row = info["row"]
                row.pack_forget()
                row.configure(width=44, height=36, fg_color=bg_container)
                row.pack_propagate(False)
                row.pack(fill="none", expand=False, pady=1)

                btn.configure(
                    text=info["icon"],
                    height=34
                )
                btn.pack(fill="both", expand=True)

        self._atualizar_destaque_botoes()
        self._atualizar_estrelas()



    # ---------------------------------------------------------
    # Seleção de Abas & Destaque
    # ---------------------------------------------------------
    def vincular_notebook(self, notebook):
        """Conecta o CTkTabview à sidebar."""
        self.notebook = notebook

    def selecionar_aba(self, tab_id, trocar_notebook=True, expandir_categoria=False):
        """Seleciona uma aba, garantindo destaque e navegação sem forçar abertura de acordeons fechados."""
        self.aba_selecionada = tab_id

        # Só expande a categoria pai se for explicitamente solicitado (evita abrir acordions fechados)
        if expandir_categoria and self.expandido and tab_id in self.modulos_map:
            cat = self.modulos_map[tab_id][2]
            if not self.grupos_abertos.get(cat, True):
                self.grupos_abertos[cat] = True
                self._atualizar_layout()

        if trocar_notebook and self.notebook:
            try:
                self.notebook.set(tab_id)
            except Exception as e:
                print(f"[Sidebar] Erro ao trocar aba: {e}")

        self._atualizar_destaque_botoes()
        if self.on_tab_changed:
            self.on_tab_changed(tab_id)

    def _atualizar_destaque_botoes(self):
        """Aplica o destaque visual no item ativo em todas as suas ocorrências."""
        paleta = self.paletas.get(self.obter_tema_atual(), {})
        accent = paleta.get("accent", "#f57c00")
        hover = paleta.get("hover", "#e65100")
        inp = paleta.get("input", "#1f3a60")
        texto = paleta.get("texto", "#ffffff")

        for info in self.botoes_info:
            b = info["btn"]
            if info["tab_id"] == self.aba_selecionada:
                b.configure(fg_color=accent, hover_color=hover, text_color="white")
            else:
                b.configure(fg_color="transparent", hover_color=inp, text_color=texto)

    def _on_btn_enter(self, btn, nome):
        """Mostra tooltip somente quando a sidebar estiver recolhida."""
        if not self.expandido:
            paleta = self.paletas.get(self.obter_tema_atual(), {})
            self.tooltip.exibir(btn, nome, paleta)

    def _on_btn_leave(self, btn):
        """Oculta o tooltip ao sair do botão."""
        self.tooltip.ocultar()

    # ---------------------------------------------------------
    # Temas Dinâmicos
    # ---------------------------------------------------------
    def aplicar_tema(self, tema_nome):
        """Reage a mudanças de tema da aplicação inteira em tempo real."""
        paleta = self.paletas.get(tema_nome, {})
        bg_container = paleta.get("container", "#111827")
        fg_texto = paleta.get("texto", "#f8fafc")
        accent = paleta.get("accent", "#00d2ff")
        hover = paleta.get("hover", "#0284c7")
        inp = paleta.get("input", "#1a2436")
        subtle = paleta.get("scrollbar_hover", "#38bdf8")

        self.frame_container.configure(fg_color=bg_container)
        self.frame_topo.configure(fg_color=bg_container)
        self.scroll.configure(fg_color=bg_container)

        self.btn_toggle.configure(fg_color=inp, hover_color=accent, text_color=fg_texto)
        self.lbl_menu_title.configure(text_color=accent)

        for nome_cat, cw in self.categoria_widgets.items():
            cw["arrow"].configure(text_color=accent)
            cw["titulo"].configure(text_color=subtle)
            cw["header"].configure(fg_color=bg_container)
            cw["itens_frame"].configure(fg_color=bg_container)

        if hasattr(self, "frame_fav_itens"):
            self.frame_fav_itens.configure(fg_color=bg_container)

        for info in self.botoes_info:
            info["row"].configure(fg_color=bg_container)

        if hasattr(self, "divider_fav"):
            self.divider_fav.configure(fg_color=inp)

        if hasattr(self, "entry_busca"):
            self.entry_busca.configure(fg_color=inp, text_color=fg_texto, border_color=inp)

        self._atualizar_destaque_botoes()
        self._atualizar_estrelas()

