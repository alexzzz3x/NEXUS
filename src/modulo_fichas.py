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


def _obter_caminho_config():
    try:
        from utils_os import obter_diretorio_executavel
        base = obter_diretorio_executavel()
    except Exception:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    pasta_data = os.path.join(base, "data")
    os.makedirs(pasta_data, exist_ok=True)
    return os.path.join(pasta_data, "config_fichas.json")


def _carregar_banco_nomes():
    """Carrega mapa de RE -> Nome completo a partir do Banco de nomes.xlsx."""
    mapa = {}
    try:
        from utils_os import obter_diretorio_executavel, obter_caminho_recurso
        caminhos = [
            os.path.join(obter_diretorio_executavel(), "data", "Banco de nomes.xlsx"),
            obter_caminho_recurso("data", "Banco de nomes.xlsx"),
        ]
        import pandas as pd
        for c in caminhos:
            if os.path.exists(c):
                df = pd.read_excel(c)
                col_re = next((col for col in df.columns if "re" in col.lower()), None)
                col_nome = next((col for col in df.columns if "nome" in col.lower()), None)
                if col_re and col_nome:
                    for _, row in df.iterrows():
                        re_val = str(row[col_re]).strip().split(".")[0]
                        nome_val = str(row[col_nome]).strip()
                        if re_val and nome_val and re_val != "nan" and nome_val != "nan":
                            mapa[re_val] = nome_val
                if mapa:
                    break
    except Exception:
        pass
    return mapa


def _carregar_dicionario_materiais():
    """Tenta carregar o catálogo de materiais para exibição amigável de nomes."""
    catalogo = {}
    try:
        from utils_os import obter_diretorio_executavel, obter_caminho_recurso
        caminhos = [
            os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx"),
            obter_caminho_recurso("data", "banco_de_dados.xlsx"),
            os.path.join(obter_diretorio_executavel(), "data", "Banco de nomes.xlsx"),
            obter_caminho_recurso("data", "Banco de nomes.xlsx"),
        ]
        import pandas as pd
        for c in caminhos:
            if os.path.exists(c):
                df = pd.read_excel(c)
                col_mat = next((col for col in df.columns if "material" in col.lower() or "codigo" in col.lower()), None)
                col_desc = next((col for col in df.columns if "descri" in col.lower() or "nome" in col.lower()), None)
                if col_mat and col_desc:
                    for _, row in df.iterrows():
                        m = str(row[col_mat]).strip().split(".")[0]
                        d = str(row[col_desc]).strip()
                        if m and d and m != "nan" and d != "nan":
                            catalogo[m] = d
                if catalogo:
                    break
    except Exception:
        pass
    return catalogo


class AbaFichas:
    """Formulário para preparar e preencher automaticamente uma Ficha de Solicitação com Materiais no WorkForce."""

    URL_NOVA_FICHA = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFichaDetail.aspx?MedicaoFichaId=0"
    URL_LOGIN = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFicha.aspx"
    DEFAULT_USUARIO = "19595"
    DEFAULT_SENHA = "979648"

    def __init__(self, parent_frame, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.paletas = paletas or {}
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
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
        return self.paletas.get(tema, self.paletas.get("Conecta Blue", {
            "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
            "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60"
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
        nome_alvo = self.mapa_nomes.get(re_str, "").lower()
        for _ in range(3):
            try:
                s = Select(select_element)
                for opt in s.options:
                    t = opt.text.lower().strip()
                    v = opt.get_attribute("value") or ""
                    if (nome_alvo and nome_alvo in t) or (re_str and (re_str in t or re_str == v)):
                        s.select_by_visible_text(opt.text)
                        return True
            except Exception:
                time.sleep(0.5)
        return False

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

        obrigatorios = [
            ("RE Resp. Almox.", re_almox),
            ("RE Eletricista 1", re_elet1)
        ]
        faltantes = [nome for nome, val in obrigatorios if not val]
        if faltantes:
            messagebox.showwarning(
                "Campos obrigatórios",
                f"Por favor, preencha o(s) seguinte(s) campo(s):\n- " + "\n- ".join(faltantes)
            )
            return

        if re_almox == re_elet1:
            if not messagebox.askyesno(
                "Atenção: Mesmo RE",
                f"O RE do Resp. Almox. e do Eletricista 1 estão iguais ({re_almox}).\n"
                "Geralmente o Eletricista tem um RE diferente do Almoxarife.\n\n"
                "Deseja prosseguir mesmo assim?"
            ):
                return

        if webdriver is None:
            webbrowser.open(self.URL_NOVA_FICHA, new=2)
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
        driver = None
        try:
            if ChromeOptions is not None:
                options = ChromeOptions()
            else:
                options = webdriver.ChromeOptions()

            options.add_argument("--start-maximized")
            options.add_experimental_option("detach", True)
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument("--disable-blink-features=AutomationControlled")

            self._atualizar_status("Iniciando Chrome...")
            driver = webdriver.Chrome(options=options)
            self.driver = driver
            wait = WebDriverWait(driver, 25)

            # 1. Tela de Login
            self._atualizar_status("Acessando autenticação do WorkForce...")
            driver.get(self.URL_LOGIN)

            usuario_el = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[id*='wtUserNameInput'], input[type='text']"))
            )
            senha_el = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[id*='wtPasswordInput'], input[type='password']"))
            )

            usuario_el.clear()
            usuario_el.send_keys(self.login_usuario)
            senha_el.clear()
            senha_el.send_keys(self.login_senha)

            btn_login = driver.find_element(
                By.CSS_SELECTOR, "input[id*='wt16'], input[type='submit'][value*='Entrar'], input[type='submit']"
            )

            self._atualizar_status(f"Logando como {self.login_usuario}...")
            btn_login.click()

            wait.until(EC.staleness_of(btn_login))
            time.sleep(1.5)

            # 2. Navegar para a Nova Ficha
            self._atualizar_status("Carregando formulário de Nova Ficha...")
            driver.get(self.URL_NOVA_FICHA)

            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idBase')]")
                )
            )
            time.sleep(1.0)

            # 3. Base (deve ser selecionada primeiro pois o WorkForce recarrega campos via Ajax)
            if dados["base"]:
                self._atualizar_status(f"Configurando Base: {dados['base']}...")
                try:
                    sel_base_el = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idBase')]")
                    s = Select(sel_base_el)
                    alvo = dados["base"].lower()
                    alterou_base = False
                    if alvo not in s.first_selected_option.text.lower():
                        for opt in s.options:
                            opt_txt = opt.text.lower()
                            if alvo in opt_txt or ("santo andr" in opt_txt and "santo andre" in alvo):
                                s.select_by_visible_text(opt.text)
                                alterou_base = True
                                break
                    if alterou_base:
                        # Aguarda o recarregamento Ajax dos funcionários da Base
                        time.sleep(3.0)
                except Exception:
                    pass

            # 4. Preenchimento de Data e Responsáveis
            self._atualizar_status("Preenchendo cabeçalho e responsáveis...")

            # Data obrigatória (DD-MM-YYYY)
            if dados["data"]:
                try:
                    el = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_data')]")
                    el.clear()
                    el.send_keys(dados["data"])
                except Exception:
                    pass

            # Resp Almox (Input + Select)
            if dados["re_almox"]:
                try:
                    inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reRespAlmox') and not(contains(@id, 'reRespAlmox2'))]")
                    inp.clear()
                    inp.send_keys(dados["re_almox"])
                    sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idRespAlmox') and not(contains(@id, 'idRespAlmox2'))]")
                    self._selecionar_dropdown(driver, sel, dados["re_almox"])
                except Exception:
                    pass

            # Resp Almox Entregue (Input + Select)
            if dados["re_almox_entregue"]:
                try:
                    inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reRespAlmox2')]")
                    inp.clear()
                    inp.send_keys(dados["re_almox_entregue"])
                    sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idRespAlmox2')]")
                    self._selecionar_dropdown(driver, sel, dados["re_almox_entregue"])
                except Exception:
                    pass

            # Eletricista 1 (Input + Select)
            if dados["re_eletricista_1"]:
                try:
                    inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reEletricista1')]")
                    inp.clear()
                    inp.send_keys(dados["re_eletricista_1"])
                    sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idEletricista1')]")
                    self._selecionar_dropdown(driver, sel, dados["re_eletricista_1"])
                except Exception:
                    pass

            # Eletricista 2 (Input + Select)
            if dados["re_eletricista_2"]:
                try:
                    inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reEletricista2')]")
                    inp.clear()
                    inp.send_keys(dados["re_eletricista_2"])
                    sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idEletricista2')]")
                    self._selecionar_dropdown(driver, sel, dados["re_eletricista_2"])
                except Exception:
                    pass

            # Recurso
            if dados["recurso"]:
                try:
                    el = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_recurso')]")
                    el.clear()
                    el.send_keys(dados["recurso"])
                except Exception:
                    pass

            # Observação
            if dados["observacao"]:
                try:
                    obs_els = driver.find_elements(By.XPATH, "//textarea[not(contains(@id, 'ModalConsulta'))]")
                    if obs_els:
                        obs_els[0].clear()
                        obs_els[0].send_keys(dados["observacao"])
                except Exception:
                    pass

            # 5. Preenchimento dos Materiais na tabela (Pedido e Fornecido)
            materiais = dados.get("materiais", [])
            if materiais:
                self._atualizar_status(f"Preenchendo {len(materiais)} material(is) na tabela...")
                time.sleep(1.0)

                js_preencher = """
                    var materiais = arguments[0];
                    var table = document.querySelector("table[id*='AlmoxarifadoMaterialMedicaoTable']");
                    if (!table) return {sucesso: false, erro: "Tabela não encontrada"};

                    var rows = table.querySelectorAll("tbody tr, tr");
                    var preenchidos = [];
                    var nao_encontrados = [];

                    for (var m = 0; m < materiais.length; m++) {
                        var codAlvo = String(materiais[m][0]).trim();
                        var qtdAlvo = String(materiais[m][1]).trim();
                        var achou = false;

                        for (var i = 0; i < rows.length; i++) {
                            var r = rows[i];
                            var cells = r.querySelectorAll("td");
                            if (cells.length >= 6) {
                                var txtCod = cells[1].textContent.trim();
                                if (txtCod === codAlvo || txtCod.indexOf(codAlvo) !== -1) {
                                    // Cell 4: Pedido, Cell 5: Fornecido
                                    var inps = r.querySelectorAll("input[type='number']");
                                    var inpPed = (cells.length > 4 ? cells[4].querySelector("input") : null) || inps[0];
                                    var inpForn = (cells.length > 5 ? cells[5].querySelector("input") : null) || inps[1];

                                    if (inpPed) {
                                        inpPed.focus();
                                        inpPed.value = qtdAlvo;
                                        inpPed.dispatchEvent(new Event('input', {bubbles: true}));
                                        inpPed.dispatchEvent(new Event('change', {bubbles: true}));
                                        inpPed.blur();
                                    }
                                    if (inpForn) {
                                        inpForn.focus();
                                        inpForn.value = qtdAlvo;
                                        inpForn.dispatchEvent(new Event('input', {bubbles: true}));
                                        inpForn.dispatchEvent(new Event('change', {bubbles: true}));
                                        inpForn.blur();
                                    }

                                    r.scrollIntoView({behavior: 'smooth', block: 'center'});
                                    achou = true;
                                    preenchidos.push(codAlvo + " (" + qtdAlvo + " un)");
                                    break;
                                }
                            }
                        }
                        if (!achou) {
                            nao_encontrados.push(codAlvo);
                        }
                    }
                    return {sucesso: true, preenchidos: preenchidos, nao_encontrados: nao_encontrados};
                """

                try:
                    res = driver.execute_script(js_preencher, materiais)
                    if res and res.get("sucesso"):
                        p_list = res.get("preenchidos", [])
                        n_list = res.get("nao_encontrados", [])
                        if n_list:
                            self._atualizar_status(f"Materiais preenchidos: {len(p_list)}. Não encontrados no site: {', '.join(n_list)}")
                        else:
                            self._atualizar_status(f"✅ {len(p_list)} material(is) preenchido(s) com sucesso na tabela!")
                    else:
                        self._atualizar_status(f"⚠️ Atenção ao preencher materiais: {res.get('erro', 'erro desconhecido')}")
                except Exception as e_js:
                    self._atualizar_status(f"⚠️ Falha ao preencher tabela via script: {e_js}")

                time.sleep(1.0)

            # 6. Salvar automaticamente ou deixar para conferência
            if dados.get("salvar_auto", False):
                self._atualizar_status("Gravando e salvando Ficha no WorkForce...")
                try:
                    btn_salvar = driver.find_element(By.XPATH, "//input[@value='Salvar']")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
                    time.sleep(0.5)
                    btn_salvar.click()

                    # Aguarda redirecionamento ou mensagem de sucesso
                    time.sleep(3.5)
                    if "MedicaoFichaId=0" not in driver.current_url:
                        self._atualizar_status("✅ Ficha criada e salva com sucesso no WorkForce!")
                    else:
                        self._atualizar_status("⚠️ Ficha preenchida! Verifique os campos e clique em Salvar no navegador.")
                except Exception as e_save:
                    self._atualizar_status(f"⚠️ Ficha preenchida! Clique em Salvar no navegador ({e_save}).")
            else:
                self._atualizar_status("✅ Ficha e materiais preenchidos com sucesso! Confira e clique em Salvar no navegador.")

        except Exception as erro:
            self._atualizar_status(f"⚠️ Atenção durante o fluxo: {erro}")
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
