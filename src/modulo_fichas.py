import os
import re
import json
import time
import datetime as _dt
import threading
import webbrowser
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
except ImportError:
    webdriver = None
    ChromeOptions = None
    ChromeService = None


from utils_workforce import (
    _obter_caminho_config,
    _carregar_banco_nomes,
    _carregar_dicionario_materiais,
    carregar_credenciais_workforce,
    executar_fluxo_workforce,
)
from utils_metrics import incrementar_metrica


class AbaFichas:
    """Formulário para preparar e preencher automaticamente uma Ficha de Solicitação com Materiais no WorkForce."""

    URL_NOVA_FICHA = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFichaDetail.aspx?MedicaoFichaId=0"
    URL_LOGIN = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFicha.aspx"
    DEFAULT_USUARIO = "19595"
    DEFAULT_SENHA = "979648"

    def __init__(self, parent_frame, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.paletas = paletas or {}
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.driver = None
        self.materiais_lista = []
        self.catalogo_materiais = {}
        self.mapa_nomes = {}
        self._carregar_credenciais()
        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

        threading.Thread(target=self._carregar_dados_async, daemon=True).start()

    def _carregar_dados_async(self):
        self.catalogo_materiais = _carregar_dicionario_materiais()
        self.mapa_nomes = _carregar_banco_nomes()

    def _carregar_credenciais(self):
        self.login_usuario = self.DEFAULT_USUARIO
        self.login_senha = self.DEFAULT_SENHA
        caminho = _obter_caminho_config()
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    self.login_usuario = dados.get("usuario", self.DEFAULT_USUARIO)
                    self.login_senha = dados.get("senha", self.DEFAULT_SENHA)
            except Exception:
                pass
        else:
            self._salvar_credenciais(self.DEFAULT_USUARIO, self.DEFAULT_SENHA)

    def _salvar_credenciais(self, usuario, senha):
        self.login_usuario = str(usuario).strip()
        self.login_senha = str(senha).strip()
        caminho = _obter_caminho_config()
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump({"usuario": self.login_usuario, "senha": self.login_senha}, f, indent=4)
        except Exception:
            pass

    def _cores(self, tema):
        return self.paletas.get(tema, self.paletas.get("Nexus", {
            "bg": "#090d16", "container": "#111827", "texto": "#f8fafc",
            "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"
        }))

    def _construir_interface(self):
        # Topo
        self.frame_topo = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_topo.pack(fill="x", padx=25, pady=(15, 6))

        self.lbl_titulo = ctk.CTkLabel(self.frame_topo, text="📋 Nova Ficha de Solicitação & Materiais",
                                       font=("Segoe UI", 20, "bold"))
        self.lbl_titulo.pack(side="left")

        self.btn_config_login = ctk.CTkButton(
            self.frame_topo, text="🔑 Credenciais WorkForce", width=170, height=32,
            command=self._modal_credenciais, font=("Segoe UI", 11, "bold")
        )
        self.btn_config_login.pack(side="right")

        self.lbl_info = ctk.CTkLabel(
            self.parent_frame,
            text="Preencha os dados da ficha e adicione os materiais a serem fornecidos para salvar automaticamente no WorkForce.",
            font=("Segoe UI", 11)
        )
        self.lbl_info.pack(anchor="w", padx=25, pady=(0, 10))

        # Container principal dividido em 2 colunas
        self.container_principal = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.container_principal.pack(fill="both", expand=True, padx=25, pady=0)
        self.container_principal.grid_columnconfigure(0, weight=4, minsize=420)
        self.container_principal.grid_columnconfigure(1, weight=5, minsize=460)
        self.container_principal.grid_rowconfigure(0, weight=1)

        # ================== COLUNA ESQUERDA: DADOS DA FICHA ==================
        self.card_dados = ctk.CTkFrame(self.container_principal, corner_radius=10)
        self.card_dados.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        self.card_dados.grid_columnconfigure(0, weight=1)
        self.card_dados.grid_columnconfigure(1, weight=1)

        lbl_tit_dados = ctk.CTkLabel(self.card_dados, text="📝 Dados Gerais da Ficha", font=("Segoe UI", 13, "bold"))
        lbl_tit_dados.grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 6))

        def label(parent, text, row, col, sticky="w"):
            w = ctk.CTkLabel(parent, text=text, font=("Segoe UI", 11, "bold"))
            w.grid(row=row, column=col, sticky=sticky, padx=14, pady=(6, 2))
            return w

        # Data & Base
        label(self.card_dados, "Data (DD-MM-AAAA)", 1, 0)
        label(self.card_dados, "Base", 1, 1)
        self.campo_data = ctk.CTkEntry(self.card_dados, height=30)
        self.campo_data.grid(row=2, column=0, sticky="ew", padx=(14, 6), pady=(0, 8))
        self.campo_data.insert(0, _dt.date.today().strftime("%d-%m-%Y"))

        self.campo_base = ctk.CTkComboBox(
            self.card_dados, height=30,
            values=["SANTO ANDRE", "MAUA", "NORTE", "INTERLAGOS", "SALVADOR", "SAO PAULO", "OUTRA"]
        )
        self.campo_base.grid(row=2, column=1, sticky="ew", padx=(6, 14), pady=(0, 8))
        self.campo_base.set("SANTO ANDRE")

        # RE Almoxarifados
        label(self.card_dados, "RE Resp. Almox. *", 3, 0)
        label(self.card_dados, "RE Almox. Entregue", 3, 1)
        self.campo_re_almox = ctk.CTkEntry(self.card_dados, height=30, placeholder_text="Ex.: 19595")
        self.campo_re_almox.grid(row=4, column=0, sticky="ew", padx=(14, 6), pady=(0, 8))
        self.campo_re_almox.insert(0, self.login_usuario)

        self.campo_re_almox_entregue = ctk.CTkEntry(self.card_dados, height=30, placeholder_text="Ex.: 19595")
        self.campo_re_almox_entregue.grid(row=4, column=1, sticky="ew", padx=(6, 14), pady=(0, 8))
        self.campo_re_almox_entregue.insert(0, self.login_usuario)

        # RE Eletricistas
        label(self.card_dados, "RE Eletricista 1 *", 5, 0)
        label(self.card_dados, "RE Eletricista 2", 5, 1)
        self.campo_re_eletricista_1 = ctk.CTkEntry(self.card_dados, height=30, placeholder_text="Ex.: 21475")
        self.campo_re_eletricista_1.grid(row=6, column=0, sticky="ew", padx=(14, 6), pady=(0, 8))

        self.campo_re_eletricista_2 = ctk.CTkEntry(self.card_dados, height=30, placeholder_text="Ex.: 22739")
        self.campo_re_eletricista_2.grid(row=6, column=1, sticky="ew", padx=(6, 14), pady=(0, 8))

        # Recurso
        label(self.card_dados, "Recurso", 7, 0)
        self.campo_recurso = ctk.CTkEntry(self.card_dados, height=30, placeholder_text="Ex.: CDC424")
        self.campo_recurso.grid(row=8, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 8))

        # Observação
        label(self.card_dados, "Observação", 9, 0)
        self.campo_observacao = ctk.CTkTextbox(self.card_dados, height=75)
        self.campo_observacao.grid(row=10, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 14))

        # ================== COLUNA DIREITA: MATERIAIS ==================
        self.card_materiais = ctk.CTkFrame(self.container_principal, corner_radius=10)
        self.card_materiais.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)

        # Título Materiais
        f_header_mat = ctk.CTkFrame(self.card_materiais, fg_color="transparent")
        f_header_mat.pack(fill="x", padx=14, pady=(12, 4))
        self.lbl_tit_mat = ctk.CTkLabel(f_header_mat, text="📦 Materiais a Fornecer", font=("Segoe UI", 13, "bold"))
        self.lbl_tit_mat.pack(side="left")

        self.lbl_tot_itens = ctk.CTkLabel(f_header_mat, text="0 itens adicionados", font=("Segoe UI", 10, "italic"))
        self.lbl_tot_itens.pack(side="right")

        # Entrada rápida (Código + Qtd + Botão)
        f_input_rapido = ctk.CTkFrame(self.card_materiais, fg_color="transparent")
        f_input_rapido.pack(fill="x", padx=14, pady=(4, 6))

        self.campo_mat_cod = ctk.CTkEntry(f_input_rapido, height=32, placeholder_text="Código (ex: 510388)", width=170)
        self.campo_mat_cod.pack(side="left", padx=(0, 6))
        self.campo_mat_cod.bind("<Return>", lambda e: self.campo_mat_qtd.focus_set())

        self.campo_mat_qtd = ctk.CTkEntry(f_input_rapido, height=32, placeholder_text="Qtd", width=75)
        self.campo_mat_qtd.pack(side="left", padx=(0, 6))
        self.campo_mat_qtd.insert(0, "1")
        self.campo_mat_qtd.bind("<Return>", lambda e: self._adicionar_material_rapido())

        self.btn_add_mat = ctk.CTkButton(
            f_input_rapido, text="➕ Adicionar", height=32, width=100,
            command=self._adicionar_material_rapido, font=("Segoe UI", 11, "bold")
        )
        self.btn_add_mat.pack(side="left")

        # Entrada em Lote / Colar Excel
        f_lote_ctrl = ctk.CTkFrame(self.card_materiais, fg_color="transparent")
        f_lote_ctrl.pack(fill="x", padx=14, pady=(4, 2))

        self.lbl_lote_dica = ctk.CTkLabel(
            f_lote_ctrl,
            text="📋 Ou cole várias linhas (Excel ou texto: '510388 12'):",
            font=("Segoe UI", 10)
        )
        self.lbl_lote_dica.pack(side="left")

        self.btn_limpar_todos = ctk.CTkButton(
            f_lote_ctrl, text="🗑️ Limpar Lista", height=24, width=90,
            command=self._limpar_materiais, font=("Segoe UI", 10)
        )
        self.btn_limpar_todos.pack(side="right")

        f_lote_box = ctk.CTkFrame(self.card_materiais, fg_color="transparent")
        f_lote_box.pack(fill="x", padx=14, pady=(2, 6))

        self.campo_lote_texto = ctk.CTkTextbox(f_lote_box, height=50, font=("Consolas", 10))
        self.campo_lote_texto.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_importar_lote = ctk.CTkButton(
            f_lote_box, text="Importar\nLote", height=50, width=80,
            command=self._importar_materiais_lote, font=("Segoe UI", 11, "bold")
        )
        self.btn_importar_lote.pack(side="right")

        # Lista visual de materiais
        self.scroll_materiais = ctk.CTkScrollableFrame(self.card_materiais, height=140, corner_radius=6)
        self.scroll_materiais.pack(fill="both", expand=True, padx=14, pady=(4, 12))

        # ================== BARRA INFERIOR: AÇÕES ==================
        self.bar_acoes = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.bar_acoes.pack(fill="x", padx=25, pady=(12, 6))

        self.btn_abrir = ctk.CTkButton(
            self.bar_acoes, text="🌐 Preencher e Salvar Ficha no WorkForce",
            command=self.abrir_ficha, height=44, font=("Segoe UI", 13, "bold")
        )
        self.btn_abrir.pack(side="left", padx=(0, 16))

        self.chk_salvar_auto = ctk.CTkCheckBox(
            self.bar_acoes, text="Salvar automaticamente no WorkForce ao finalizar",
            font=("Segoe UI", 11, "bold")
        )
        self.chk_salvar_auto.pack(side="left")
        self.chk_salvar_auto.select()

        self.lbl_status = ctk.CTkLabel(self.parent_frame, text="Pronto para abrir nova ficha.", font=("Segoe UI", 11))
        self.lbl_status.pack(anchor="w", padx=25, pady=(2, 10))

    def _adicionar_material_rapido(self):
        cod = self.campo_mat_cod.get().strip()
        qtd_str = self.campo_mat_qtd.get().strip()

        if not cod:
            messagebox.showwarning("Aviso", "Informe o código do material.", parent=self.parent_frame)
            self.campo_mat_cod.focus_set()
            return

        try:
            qtd = int(qtd_str) if qtd_str else 1
            if qtd <= 0:
                qtd = 1
        except ValueError:
            messagebox.showwarning("Aviso", "A quantidade deve ser um número inteiro.", parent=self.parent_frame)
            self.campo_mat_qtd.focus_set()
            return

        self._adicionar_item_lista(cod, qtd)
        self.campo_mat_cod.delete(0, "end")
        self.campo_mat_qtd.delete(0, "end")
        self.campo_mat_qtd.insert(0, "1")
        self.campo_mat_cod.focus_set()

    def _importar_materiais_lote(self):
        texto = self.campo_lote_texto.get("1.0", "end").strip()
        if not texto:
            messagebox.showinfo("Aviso", "Cole linhas com código e quantidade na caixa de texto.", parent=self.parent_frame)
            return

        adicionados = 0
        linhas = texto.splitlines()
        for linha in linhas:
            linha = linha.strip()
            if not linha:
                continue
            partes = re.split(r'[\t;, -]+', linha)
            if partes and partes[0].strip():
                cod = partes[0].strip()
                qtd = 1
                if len(partes) > 1 and partes[1].strip().isdigit():
                    qtd = int(partes[1].strip())
                self._adicionar_item_lista(cod, qtd, renderizar_imediato=False)
                adicionados += 1

        self.campo_lote_texto.delete("1.0", "end")
        self._renderizar_lista_materiais()
        if adicionados > 0:
            messagebox.showinfo("Importação", f"{adicionados} item(ns) importado(s) com sucesso!", parent=self.parent_frame)

    def _adicionar_item_lista(self, codigo, qtd, renderizar_imediato=True):
        for i, (c, q) in enumerate(self.materiais_lista):
            if c == codigo:
                self.materiais_lista[i] = (c, q + qtd)
                if renderizar_imediato:
                    self._renderizar_lista_materiais()
                return

        self.materiais_lista.append((codigo, qtd))
        if renderizar_imediato:
            self._renderizar_lista_materiais()

    def _remover_material(self, idx):
        if 0 <= idx < len(self.materiais_lista):
            self.materiais_lista.pop(idx)
            self._renderizar_lista_materiais()

    def _limpar_materiais(self):
        self.materiais_lista.clear()
        self._renderizar_lista_materiais()

    def _renderizar_lista_materiais(self):
        for w in self.scroll_materiais.winfo_children():
            w.destroy()

        tema = self.obter_tema_atual()
        c = self._cores(tema)

        tot_unidades = sum(q for _, q in self.materiais_lista)
        self.lbl_tot_itens.configure(text=f"{len(self.materiais_lista)} item(ns) | Total: {tot_unidades} un")

        if not self.materiais_lista:
            lbl_vazio = ctk.CTkLabel(
                self.scroll_materiais,
                text="Nenhum material adicionado ainda. Insira acima ou cole do Excel.",
                font=("Segoe UI", 10, "italic"), text_color="gray"
            )
            lbl_vazio.pack(pady=20)
            return

        for idx, (cod, qtd) in enumerate(self.materiais_lista):
            desc = self.catalogo_materiais.get(cod, "")
            texto_item = f"Código: {cod}"
            if desc:
                d_curta = desc if len(desc) <= 30 else desc[:28] + "..."
                texto_item += f" ({d_curta})"

            row_frame = ctk.CTkFrame(self.scroll_materiais, fg_color=c["input"], height=30, corner_radius=6)
            row_frame.pack(fill="x", pady=2, padx=4)

            lbl_cod = ctk.CTkLabel(row_frame, text=texto_item, font=("Segoe UI", 10, "bold"), text_color=c["texto"])
            lbl_cod.pack(side="left", padx=8)

            lbl_qtd = ctk.CTkLabel(
                row_frame, text=f"Fornecido: {qtd} un",
                font=("Segoe UI", 10, "bold"), text_color=c["accent"]
            )
            lbl_qtd.pack(side="right", padx=(0, 6))

            btn_del = ctk.CTkButton(
                row_frame, text="✕", width=22, height=22,
                fg_color="#d32f2f", hover_color="#b71c1c", text_color="white",
                font=("Segoe UI", 10, "bold"), command=lambda i=idx: self._remover_material(i)
            )
            btn_del.pack(side="right", padx=(0, 6))

    def _modal_credenciais(self):
        janela = ctk.CTkToplevel(self.parent_frame)
        janela.title("Credenciais WorkForce")
        janela.geometry("380x260")
        janela.resizable(False, False)
        janela.attributes("-topmost", True)

        tema = self.obter_tema_atual()
        c = self._cores(tema)
        janela.configure(fg_color=c["bg"])

        lbl_tit = ctk.CTkLabel(janela, text="🔑 Configurar Login WorkForce",
                               font=("Segoe UI", 15, "bold"), text_color=c["texto"])
        lbl_tit.pack(pady=(20, 12))

        f_inputs = ctk.CTkFrame(janela, fg_color=c["container"], corner_radius=10)
        f_inputs.pack(fill="x", padx=24, pady=5)

        lbl_u = ctk.CTkLabel(f_inputs, text="Usuário (RE):", font=("Segoe UI", 11, "bold"), text_color=c["texto"])
        lbl_u.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 2))
        ent_u = ctk.CTkEntry(f_inputs, fg_color=c["input"], text_color=c["texto"], border_color=c["container"])
        ent_u.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        ent_u.insert(0, self.login_usuario)

        lbl_s = ctk.CTkLabel(f_inputs, text="Senha:", font=("Segoe UI", 11, "bold"), text_color=c["texto"])
        lbl_s.grid(row=2, column=0, sticky="w", padx=14, pady=(4, 2))
        ent_s = ctk.CTkEntry(f_inputs, show="*", fg_color=c["input"], text_color=c["texto"], border_color=c["container"])
        ent_s.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 14))
        ent_s.insert(0, self.login_senha)
        f_inputs.grid_columnconfigure(0, weight=1)

        def salvar():
            u = ent_u.get().strip()
            s = ent_s.get().strip()
            if not u or not s:
                messagebox.showwarning("Aviso", "Informe o usuário e a senha.", parent=janela)
                return
            self._salvar_credenciais(u, s)
            messagebox.showinfo("Sucesso", "Credenciais salvas com sucesso!", parent=janela)
            janela.destroy()

        btn_salvar = ctk.CTkButton(janela, text="Salvar Credenciais", command=salvar,
                                   fg_color=c["accent"], hover_color=c["hover"], text_color="white",
                                   font=("Segoe UI", 11, "bold"), height=32)
        btn_salvar.pack(pady=(12, 0))

    def _selecionar_dropdown(self, driver, select_element, re_codigo):
        """Seleciona a opção do dropdown combinando o RE ou o nome do Banco de nomes."""
        re_str = str(re_codigo).strip()
        nome_alvo = (self.mapa_nomes.get(re_str, "") if self.mapa_nomes else "").lower().strip()
        partes_nome = [p for p in nome_alvo.split() if len(p) > 2]
        for _ in range(3):
            try:
                s = Select(select_element)
                for opt in s.options:
                    t = opt.text.lower().strip()
                    v = opt.get_attribute("value") or ""
                    if re_str and (re_str in t or re_str == v):
                        s.select_by_visible_text(opt.text)
                        return True
                    if nome_alvo and nome_alvo in t:
                        s.select_by_visible_text(opt.text)
                        return True
                    if len(partes_nome) >= 2 and partes_nome[0] in t and partes_nome[-1] in t:
                        s.select_by_visible_text(opt.text)
                        return True
            except Exception:
                time.sleep(0.5)
        return False

    def carregar_dados_ficha(self, dados):
        """Carrega externamente dados na aba Fichas (ex: vindo da Contagem de Carro)."""
        if not dados:
            return
        if "data" in dados and hasattr(self, "campo_data"):
            self.campo_data.delete(0, "end")
            self.campo_data.insert(0, str(dados["data"]))
        if "base" in dados and hasattr(self, "campo_base"):
            self.campo_base.set(str(dados["base"]))
        if "re_almox" in dados and hasattr(self, "campo_re_almox"):
            self.campo_re_almox.delete(0, "end")
            self.campo_re_almox.insert(0, str(dados["re_almox"]))
        if "re_almox_entregue" in dados and hasattr(self, "campo_re_almox_entregue"):
            self.campo_re_almox_entregue.delete(0, "end")
            self.campo_re_almox_entregue.insert(0, str(dados["re_almox_entregue"]))
        if "re_eletricista_1" in dados and hasattr(self, "campo_re_eletricista_1"):
            self.campo_re_eletricista_1.delete(0, "end")
            self.campo_re_eletricista_1.insert(0, str(dados["re_eletricista_1"]))
        if "re_eletricista_2" in dados and hasattr(self, "campo_re_eletricista_2"):
            self.campo_re_eletricista_2.delete(0, "end")
            self.campo_re_eletricista_2.insert(0, str(dados["re_eletricista_2"]))
        if "recurso" in dados and hasattr(self, "campo_recurso"):
            self.campo_recurso.delete(0, "end")
            self.campo_recurso.insert(0, str(dados["recurso"]))
        if "observacao" in dados and hasattr(self, "campo_observacao"):
            self.campo_observacao.delete("1.0", "end")
            self.campo_observacao.insert("1.0", str(dados["observacao"]))
        if "materiais" in dados:
            self.materiais_lista.clear()
            for item in dados["materiais"]:
                try:
                    cod = str(item[0]).strip()
                    qtd = int(item[1])
                    self._adicionar_item_lista(cod, qtd, renderizar_imediato=False)
                except Exception:
                    pass
            self._renderizar_lista_materiais()
        if not dados.get("salvar_auto", False):
            self.chk_salvar_auto.deselect()

    def abrir_ficha(self):
        # 1. Se houver código digitado na entrada rápida e não adicionado, adiciona automaticamente
        cod_pendente = self.campo_mat_cod.get().strip()
        if cod_pendente:
            qtd_str = self.campo_mat_qtd.get().strip()
            try:
                qtd_pendente = int(qtd_str) if qtd_str else 1
                if qtd_pendente <= 0:
                    qtd_pendente = 1
            except ValueError:
                qtd_pendente = 1
            self._adicionar_item_lista(cod_pendente, qtd_pendente)
            self.campo_mat_cod.delete(0, "end")
            self.campo_mat_qtd.delete(0, "end")
            self.campo_mat_qtd.insert(0, "1")

        # 2. Se houver texto colado na caixa de lote e não importado, importa automaticamente
        texto_lote = self.campo_lote_texto.get("1.0", "end").strip()
        if texto_lote:
            self._importar_materiais_lote()

        re_almox = self.campo_re_almox.get().strip()
        re_elet1 = self.campo_re_eletricista_1.get().strip()

        if not re_almox:
            messagebox.showwarning(
                "Campo obrigatório",
                "Por favor, preencha o campo RE Resp. Almox."
            )
            return

        # Se eletricista 1 não foi preenchido, desmarca salvar_auto para que o usuário preencha manualmente no navegador
        salvar_auto = bool(self.chk_salvar_auto.get())
        if not re_elet1:
            salvar_auto = False
            self.chk_salvar_auto.deselect()

        if re_elet1 and re_almox == re_elet1:
            if not messagebox.askyesno(
                "Atenção: Mesmo RE",
                f"O RE do Resp. Almox. e do Eletricista 1 estão iguais ({re_almox}).\n"
                "Geralmente o Eletricista tem um RE diferente do Almoxarife.\n\n"
                "Deseja prosseguir mesmo assim?"
            ):
                return

        if webdriver is None:
            webbrowser.open(self.URL_NOVA_FICHA, new=2)
            try:
                incrementar_metrica("fichas_abertas", 1)
            except Exception:
                pass
            self.lbl_status.configure(text="⚠️ Selenium não instalado; formulário aberto para preenchimento manual.")
            return

        # Ajusta formato da data para DD-MM-YYYY (padrão exigido pelo WorkForce)
        raw_data = self.campo_data.get().strip()
        data_formatada = raw_data.replace("/", "-")

        dados_formulario = {
            "data": data_formatada,
            "base": self.campo_base.get().strip(),
            "re_almox": re_almox,
            "re_almox_entregue": self.campo_re_almox_entregue.get().strip(),
            "re_eletricista_1": re_elet1,
            "re_eletricista_2": self.campo_re_eletricista_2.get().strip(),
            "recurso": self.campo_recurso.get().strip(),
            "observacao": self.campo_observacao.get("1.0", "end").strip(),
            "materiais": list(self.materiais_lista),
            "salvar_auto": bool(self.chk_salvar_auto.get())
        }

        self.btn_abrir.configure(state="disabled")
        self.lbl_status.configure(text="🚀 Iniciando navegador Chrome e efetuando login no WorkForce...")

        threading.Thread(target=self._executar_fluxo, args=(dados_formulario,), daemon=True).start()

    def _executar_fluxo(self, dados):
        try:
            sucesso, res = executar_fluxo_workforce(
                dados=dados,
                callback_status=self._atualizar_status,
                login_usuario=self.login_usuario,
                login_senha=self.login_senha,
                mapa_nomes=self.mapa_nomes
            )
            if sucesso and res:
                self.driver = res
                try:
                    incrementar_metrica("fichas_abertas", 1)
                except Exception:
                    pass
        finally:
            self.parent_frame.after(0, lambda: self.btn_abrir.configure(state="normal"))

    def _atualizar_status(self, texto):
        self.parent_frame.after(0, lambda: self.lbl_status.configure(text=texto))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c["bg"])
        self.card_dados.configure(fg_color=c["container"])
        self.card_materiais.configure(fg_color=c["container"])
        self.scroll_materiais.configure(fg_color=c["container"])

        for w in [self.lbl_titulo, self.lbl_info, self.lbl_status, self.lbl_tit_mat, self.lbl_tot_itens, self.lbl_lote_dica]:
            try:
                w.configure(text_color=c["texto"])
            except Exception:
                pass

        campos_texto = [
            self.campo_data, self.campo_re_almox,
            self.campo_re_almox_entregue, self.campo_re_eletricista_1,
            self.campo_re_eletricista_2, self.campo_recurso, self.campo_observacao,
            self.campo_mat_cod, self.campo_mat_qtd, self.campo_lote_texto
        ]
        for w in campos_texto:
            try:
                w.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"])
            except Exception:
                pass

        try:
            self.campo_base.configure(
                fg_color=c["input"], text_color=c["texto"],
                button_color=c["accent"], button_hover_color=c["hover"],
                dropdown_fg_color=c["container"], dropdown_text_color=c["texto"],
                border_color=c["container"]
            )
        except Exception:
            pass

        self.btn_abrir.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_add_mat.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_importar_lote.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_limpar_todos.configure(fg_color="transparent", hover_color=c["input"], text_color="#ef5350")
        self.btn_config_login.configure(fg_color=c["container"], hover_color=c["input"], text_color=c["texto"],
                                        border_width=1, border_color=c["accent"])
        self.chk_salvar_auto.configure(text_color=c["texto"], fg_color=c["accent"], hover_color=c["hover"])

        self._renderizar_lista_materiais()


def construir_aba_fichas(parent_frame, paletas=None, obter_tema_atual=None):
    return AbaFichas(parent_frame, paletas, obter_tema_atual)
