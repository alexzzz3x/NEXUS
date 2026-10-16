import os
import io
import sys
import json
import subprocess
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import customtkinter as ctk
import win32api
import win32print

import pandas as pd
from gerador import gerar_pdf_placa_tradicional, gerar_html_para_impressao

def imprimir_direto_na_impressora(deposito, material, descricao, qtd, unidade, lote="", tamanho="Grande", qr_topo=""):
    impressora = win32print.GetDefaultPrinter()
    html_file = gerar_html_para_impressao(deposito, material, descricao, qtd, unidade, lote=lote, tamanho=tamanho, qr_topo=qr_topo)
    os.startfile(html_file)
    return True, impressora

CAMINHO_PLANILHA_CONTROLE_PADRAO = r"\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm"

from utils_os import obter_diretorio_executavel

def obter_diretorio_app():
    return obter_diretorio_executavel()

CONFIG_LOTES = os.path.join(obter_diretorio_app(), "data", "config_lotes.json")
GOOGLE_APPS_SCRIPT_URL_PADRAO = "https://script.google.com/macros/s/AKfycbwfZ0U67WaIoXw9Cy1jz9BISe4X2PFYbx2k-nX34m49JeiFZUuK_BQlTogjw8OICa0h/exec"
CONFIG_GOOGLE_SCRIPT = os.path.join(obter_diretorio_app(), "data", "config_google_script.json")

def carregar_google_script_url():
    try:
        if os.path.exists(CONFIG_GOOGLE_SCRIPT):
            with open(CONFIG_GOOGLE_SCRIPT, "r", encoding="utf-8") as f:
                dados = json.load(f)
            url = str(dados.get("url", "")).strip()
            if url: return url
    except Exception: pass
    return GOOGLE_APPS_SCRIPT_URL_PADRAO

def salvar_google_script_url(url):
    os.makedirs(os.path.dirname(CONFIG_GOOGLE_SCRIPT), exist_ok=True)
    with open(CONFIG_GOOGLE_SCRIPT, "w", encoding="utf-8") as f:
        json.dump({"url": url}, f, ensure_ascii=False, indent=2)

def carregar_caminho_planilha_lotes():
    try:
        if os.path.exists(CONFIG_LOTES):
            with open(CONFIG_LOTES, "r", encoding="utf-8") as f:
                dados = json.load(f)
            caminho = str(dados.get("caminho", "")).strip()
            if caminho: return caminho
    except Exception: pass
    return CAMINHO_PLANILHA_CONTROLE_PADRAO

def salvar_caminho_planilha_lotes(caminho):
    os.makedirs(os.path.dirname(CONFIG_LOTES), exist_ok=True)
    with open(CONFIG_LOTES, "w", encoding="utf-8") as f:
        json.dump({"caminho": caminho}, f, ensure_ascii=False, indent=2)

class AbaEtiquetas:
    def __init__(self, parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status=None):
        self.parent_frame = parent_frame
        self.root = root
        self.arquivo_banco = arquivo_banco
        self.paletas = paletas
        self.obter_tema_atual = obter_tema_atual
        self.label_status = label_status
        self.banco = pd.DataFrame()
        self._carregando_banco = False
        self.janela_sugestoes = None
        self.lista_sugestoes = None
        self.seriais_lote = []

        self._construir_interface()
        self.atualizar_preview()

    def carregar_banco_dados_async(self, callback_final=None):
        if self._carregando_banco:
            return
        self._carregando_banco = True
        self._atualizar_status("Carregando banco de materiais...")

        def _worker():
            sucesso = False
            df_novo = pd.DataFrame()
            try:
                if os.path.exists(self.arquivo_banco):
                    df_novo = pd.read_excel(self.arquivo_banco)
                    df_novo.columns = df_novo.columns.str.strip()
                    if "Material" in df_novo.columns:
                        df_novo["Material"] = df_novo["Material"].astype(str).str.strip()
                    sucesso = True
            except Exception:
                pass

            def _aplicar():
                self._carregando_banco = False
                if sucesso:
                    self.banco = df_novo
                    self._atualizar_status("Banco de materiais pronto.")
                if callback_final:
                    callback_final(sucesso)

            try:
                self.parent_frame.after(0, _aplicar)
            except Exception:
                pass

        import threading
        threading.Thread(target=_worker, daemon=True).start()

    def carregar_banco_dados(self):
        try:
            if os.path.exists(self.arquivo_banco):
                self.banco = pd.read_excel(self.arquivo_banco)
                self.banco.columns = self.banco.columns.str.strip()
                if "Material" in self.banco.columns:
                    self.banco["Material"] = self.banco["Material"].astype(str).str.strip()
                return True
        except Exception:
            pass
        self.banco = pd.DataFrame()
        return False

    def forcar_atualizacao_banco(self):
        self.btn_atualizar_banco.configure(state="disabled")
        def on_done(sucesso):
            try:
                self.btn_atualizar_banco.configure(state="normal")
                if sucesso:
                    messagebox.showinfo("Sucesso", "Banco de dados atualizado com sucesso!", parent=self.parent_frame.winfo_toplevel())
                else:
                    messagebox.showwarning("Aviso", f"Não foi possível ler o arquivo:\n{self.arquivo_banco}", parent=self.parent_frame.winfo_toplevel())
            except Exception:
                pass
        self.carregar_banco_dados_async(on_done)

    def buscar_lote(self):
        lote = self.campo_lote.get().strip()
        if not lote:
            messagebox.showwarning("Aviso", "Digite o número do lote primeiro.")
            return
            
        self.btn_buscar_lote.configure(state="disabled")
        self._atualizar_status("Lendo planilha de controle de seriais. Aguarde...")
        
        import threading
        threading.Thread(target=self._buscar_lote_thread, args=(lote,), daemon=True).start()

    def _buscar_lote_thread(self, lote):
        try:
            caminho_planilha = carregar_caminho_planilha_lotes()
            if not os.path.exists(caminho_planilha):
                self.parent_frame.after(0, lambda: messagebox.showerror("Planilha não encontrada", f"Não encontrei a planilha de lotes neste caminho:\n{caminho_planilha}\n\nClique no botão ... ao lado de Buscar Lote para selecionar o novo arquivo."))
                return

            with open(caminho_planilha, "rb") as f_raw:
                buf = io.BytesIO(f_raw.read())
            try:
                df = pd.read_excel(buf, sheet_name="ESTOQUE", engine="openpyxl")
            finally:
                buf.close()
            df.columns = df.columns.str.strip()
            if "LOTE" not in df.columns:
                self.parent_frame.after(0, lambda: messagebox.showerror("Erro", "Coluna 'LOTE' não encontrada na planilha."))
                return

            df["LOTE"] = df["LOTE"].astype(str).str.strip()
            filtrado = df[df["LOTE"] == lote]
            if filtrado.empty:
                self.parent_frame.after(0, lambda: messagebox.showwarning("Aviso", f"Lote '{lote}' não encontrado na planilha."))
                return

            self.seriais_lote = filtrado["SERIE"].astype(str).tolist()
            qtd = len(self.seriais_lote)

            def atualizar_interface():
                self.campo_qtd.delete(0, tk.END)
                self.campo_qtd.insert(0, str(qtd))

                if "MATERIAL" in filtrado.columns:
                    material = filtrado.iloc[0]["MATERIAL"]
                    self.campo_material.delete(0, tk.END)
                    self.campo_material.insert(0, str(material))
                    self.buscar_material()

                if "DEPOSITO" in filtrado.columns:
                    deposito = str(filtrado.iloc[0]["DEPOSITO"]).strip()
                    if deposito in ["CL03", "CL04", "CS06"]:
                        self.deposito_var.set(deposito)

                self.atualizar_preview()
                self._atualizar_status(f"Lote {lote}: {qtd} seriais encontrados.")
                
            self.parent_frame.after(0, atualizar_interface)
            
        except Exception as e:
            caminho_planilha = carregar_caminho_planilha_lotes()
            self.parent_frame.after(0, lambda e=e, c=caminho_planilha: messagebox.showerror("Erro ao Ler Planilha", f"Não foi possível ler a planilha de controle:\n{c}\n\n{str(e)}"))
        finally:
            self.parent_frame.after(0, lambda: self.btn_buscar_lote.configure(state="normal"))

    def selecionar_planilha_lotes(self):
        caminho = filedialog.askopenfilename(title="Selecionar planilha de controle de lotes", filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not caminho: return
        try:
            salvar_caminho_planilha_lotes(caminho)
            self._atualizar_status("Planilha de lotes atualizada.")
            messagebox.showinfo("Planilha de Lotes", f"Novo caminho salvo:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o caminho da planilha:\n{e}")

    def configurar_google_script_url(self):
        atual = carregar_google_script_url()
        nova_url = simpledialog.askstring("URL Google Apps Script", "Cole a URL /exec do Google Apps Script:", initialvalue=atual, parent=self.root)
        if nova_url is None: return
        nova_url = nova_url.strip()
        if not nova_url:
            messagebox.showwarning("Aviso", "A URL não pode ficar vazia.")
            return
        try:
            salvar_google_script_url(nova_url)
            self._atualizar_status("URL do Google Apps Script atualizada.")
            messagebox.showinfo("Google Apps Script", "URL salva com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar a URL:\n{e}")

    def gerar_doc_google(self):
        if not self.seriais_lote:
            messagebox.showwarning("Aviso", "Busque um lote primeiro antes de gerar o documento.")
            return
        google_script_url = carregar_google_script_url()
        if not google_script_url:
            messagebox.showwarning("Aviso", "Configure a URL do Google Apps Script primeiro.")
            return

        lote = self.campo_lote.get().strip()
        material = self.campo_material.get().strip()
        descricao = self.campo_descricao.get().strip()
        deposito = self.deposito_var.get()
        dados = {"lote": lote, "material": material, "descricao": descricao, "deposito": deposito, "seriais": self.seriais_lote, "quantidade": len(self.seriais_lote)}

        try:
            dados_json = json.dumps(dados).encode("utf-8")
            requisicao = urllib.request.Request(google_script_url, data=dados_json, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(requisicao, timeout=30) as resposta:
                resultado = json.loads(resposta.read().decode("utf-8"))
                link = resultado.get("link", "")
                if link:
                    self.campo_qr_topo.delete(0, tk.END)
                    self.campo_qr_topo.insert(0, link)
                    self.atualizar_preview()
                    self._atualizar_status(f"Google Doc criado: {link}")
                    messagebox.showinfo("Sucesso", f"Documento criado com sucesso!\n\nLink: {link}")
                else:
                    messagebox.showerror("Erro", "Nenhum link retornado pelo Google Apps Script.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar documento:\n{str(e)}")

    def aplicar_tema(self, tema_nome):
        c = self.paletas[tema_nome]
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for campo in [self.campo_material, self.campo_descricao, self.campo_lote, self.campo_qtd, self.campo_qr_topo]:
            campo.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"])

        self.deposito_var.configure(fg_color=c["input"], text_color=c["texto"], button_color=c["container"], border_color=c["container"], dropdown_fg_color=c["input"], dropdown_text_color=c["texto"])
        self.combo_tamanho.configure(fg_color=c["input"], text_color=c["texto"], button_color=c["container"], border_color=c["container"], dropdown_fg_color=c["input"], dropdown_text_color=c["texto"])
        self.combo_unidade.configure(fg_color=c["input"], text_color=c["texto"], button_color=c["container"], border_color=c["container"], dropdown_fg_color=c["input"], dropdown_text_color=c["texto"])
        
        self.check_abrir.configure(fg_color=c["accent"], text_color=c["texto"])
        self.check_imprimir.configure(fg_color=c["accent"], text_color=c["texto"])

        self.btn_gerar.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_atualizar_banco.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_buscar_lote.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_planilha_lotes.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_gerar_doc.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_config_google.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])

        for lbl in self.labels: lbl.configure(text_color=c["texto"])

        self.grupo_config.configure(fg_color="transparent")
        self.grupo_prod.configure(fg_color="transparent")
        self.grupo_inv.configure(fg_color="transparent")
        self.grupo_preview.configure(fg_color="transparent")
        self.coluna_esquerda.configure(fg_color="transparent")
        self.coluna_direita.configure(fg_color="transparent")
        self.frame_opcoes.configure(fg_color="transparent")

    def fechar_sugestoes(self):
        if self.janela_sugestoes and self.janela_sugestoes.winfo_exists(): self.janela_sugestoes.destroy()
        self.janela_sugestoes = None
        self.lista_sugestoes = None

    def verificar_fechar(self, event=None):
        self.root.after(150, self._executar_fechar_se_fora)

    def _executar_fechar_se_fora(self):
        if not self.janela_sugestoes: return
        foco = self.root.focus_get()
        if foco != self.campo_material and foco != self.lista_sugestoes: self.fechar_sugestoes()

    def filtrar_sugestoes(self, event):
        self.atualizar_preview()
        if event.keysym == "Down" and self.janela_sugestoes:
            self.lista_sugestoes.focus_set()
            self.lista_sugestoes.selection_set(0)
            return
        if event.keysym in ["Up", "Return", "Escape", "Tab"]: return

        texto_digitado = self.campo_material.get().strip()
        if not texto_digitado:
            self.fechar_sugestoes()
            self.campo_descricao.delete(0, tk.END)
            return

        if self.banco.empty or "Material" not in self.banco.columns:
            if not self._carregando_banco:
                self.carregar_banco_dados_async()
            self.fechar_sugestoes()
            return

        resultados = self.banco[self.banco["Material"].str.contains(texto_digitado, case=False, na=False)]
        opcoes = list(resultados["Material"].unique())[:8]
        if not opcoes:
            self.fechar_sugestoes()
            return

        c = self.paletas[self.obter_tema_atual()]
        if not self.janela_sugestoes or not self.janela_sugestoes.winfo_exists():
            self.janela_sugestoes = tk.Toplevel(self.root)
            self.janela_sugestoes.wm_overrideredirect(True)
            self.janela_sugestoes.wm_attributes("-topmost", True)

            self.lista_sugestoes = tk.Listbox(
                self.janela_sugestoes, bg=c["input"], fg=c["texto"], selectbackground=c["accent"], selectforeground="white",
                bd=0, highlightthickness=1, highlightbackground=c["accent"], font=("Segoe UI", 12), activestyle="none"
            )
            self.lista_sugestoes.pack(fill="both", expand=True)
            self.lista_sugestoes.bind("<ButtonRelease-1>", self.preencher_selecao)
            self.lista_sugestoes.bind("<Return>", self.preencher_selecao)
            self.lista_sugestoes.bind("<FocusOut>", self.verificar_fechar)

        self.lista_sugestoes.delete(0, tk.END)
        for opcao in opcoes: self.lista_sugestoes.insert(tk.END, opcao)

        x = self.campo_material.winfo_rootx()
        y = self.campo_material.winfo_rooty() + self.campo_material.winfo_height() + 2
        w = self.campo_material.winfo_width()
        h = len(opcoes) * 24 + 4
        self.janela_sugestoes.geometry(f"{w}x{h}+{x}+{y}")

    def preencher_selecao(self, event=None):
        if self.lista_sugestoes and self.lista_sugestoes.curselection():
            selecionado = self.lista_sugestoes.get(self.lista_sugestoes.curselection())
            self.campo_material.delete(0, tk.END)
            self.campo_material.insert(0, selecionado)
            self.buscar_material()
            self.fechar_sugestoes()
            self.campo_qtd.focus()

    def buscar_material(self, event=None):
        self.campo_descricao.delete(0, tk.END)
        if self.banco.empty or "Material" not in self.banco.columns:
            self.atualizar_preview()
            return
        res = self.banco[self.banco["Material"] == str(self.campo_material.get().strip())]
        coluna_descricao = self._obter_coluna_descricao()
        if not res.empty and coluna_descricao: self.campo_descricao.insert(0, res.iloc[0][coluna_descricao])
        self.atualizar_preview()

    def limpar_campos(self):
        self.fechar_sugestoes()
        self.campo_lote.delete(0, tk.END)
        self.campo_qr_topo.delete(0, tk.END)
        self.campo_lote.focus()
        self.atualizar_preview()

    def executar_geracao(self):
        mat_cod = self.campo_material.get().strip()
        lote_val = self.campo_lote.get().strip()
        dep_val = self.deposito_var.get()
        qtd_val = self.campo_qtd.get().strip()
        qr_topo_val = self.campo_qr_topo.get().strip()
        descricao = self.campo_descricao.get().strip()
        unidade = self.combo_unidade.get()
        tamanho = self.combo_tamanho.get()

        erro = self._validar_dados(dep_val)
        if erro:
            messagebox.showwarning("Aviso", erro)
            return

        try:
            arquivo = gerar_pdf_placa_tradicional(dep_val, mat_cod, descricao, qtd_val, unidade, lote=lote_val, tamanho=tamanho, qr_topo=qr_topo_val)
            material_status = mat_cod if mat_cod else "identificação livre"

            if self.check_imprimir_var.get():
                ok, resultado = imprimir_direto_na_impressora(dep_val, mat_cod, descricao, qtd_val, unidade, lote=lote_val, tamanho=tamanho, qr_topo=qr_topo_val)
                if ok: self._atualizar_status(f" Impresso direto em: {resultado}")
                else:
                    html_file = gerar_html_para_impressao(dep_val, mat_cod, descricao, qtd_val, unidade, lote=lote_val, tamanho=tamanho, qr_topo=qr_topo_val)
                    os.startfile(html_file)
                    self._atualizar_status(f" Fallback navegador: {resultado}")
            else:
                if self.check_abrir_var.get(): os.startfile(arquivo)
                self._atualizar_status(f" PDF gerado: {material_status}")

            messagebox.showinfo("Sucesso", "Etiqueta gerada com sucesso!")
            self.limpar_campos()
        except Exception as e:
            messagebox.showerror("Erro ao Gerar", str(e))

    def atualizar_preview(self, event=None):
        self.preview.delete("all")
        dep = self.deposito_var.get()
        tamanho = self.combo_tamanho.get()
        mat = self.campo_material.get().strip()
        desc = self.campo_descricao.get().strip()
        quantidade = self.campo_qtd.get().strip()
        un = self.combo_unidade.get()
        lt = self.campo_lote.get().strip()
        qr_t = self.campo_qr_topo.get().strip()

        mapa_largura = {"Pequena": 250, "Média": 350, "Grande": 450}
        largura = mapa_largura.get(tamanho, 450)
        x1, x2 = 240 - (largura // 2), 240 + (largura // 2)

        self.preview.create_rectangle(x1, 15, x2, 305, fill="white", outline="black", width=2)
        self.preview.create_text(x2 - 10, 30, text=f"TAM: {tamanho.upper()}", font=("Segoe UI", 8, "bold"), fill="gray", anchor="ne")
        self.preview.create_line(x1, 65, x2, 65, fill="black", width=1)
        self.preview.create_text(x1 + 15, 40, text="CONECTA LOG", font=("Segoe UI", 11, "bold"), fill="black", anchor="w")

        if qr_t:
            self.preview.create_rectangle(x2 - 45, 20, x2 - 15, 50, fill="black", outline="black")
            self.preview.create_rectangle(x2 - 41, 24, x2 - 19, 46, fill="white", outline="white")
            self.preview.create_rectangle(x2 - 35, 30, x2 - 25, 40, fill="black", outline="black")
            self.preview.create_text(x2 - 53, 40, text=f"DEP: {dep}", font=("Segoe UI", 12, "bold"), fill="black", anchor="e")
        else:
            self.preview.create_text(x2 - 15, 40, text=f"DEP: {dep}", font=("Segoe UI", 12, "bold"), fill="black", anchor="e")

        self.preview.create_text(240, 120, text=dep, font=("Segoe UI", 48, "bold"), fill="black", anchor="center")
        self.preview.create_text(240, 175, text=mat if mat else "-----", font=("Segoe UI", 22, "bold"), fill="black", anchor="center")

        desc_resumida = desc[:45] + "..." if len(desc) > 45 else desc
        self.preview.create_text(240, 210, text=desc_resumida if desc else "[Aguardando Descrição]", font=("Segoe UI", 11, "bold"), fill="black" if desc else "gray", anchor="center")
        self.preview.create_text(240, 245, text=f"QUANTIDADE: {quantidade} {un}" if quantidade else "QUANTIDADE: -", font=("Segoe UI", 14, "bold"), fill="black", anchor="center")
        if lt: self.preview.create_text(240, 270, text=f"LOTE: {lt}", font=("Segoe UI", 14, "bold"), fill="black", anchor="center")

        if mat:
            for x in range(240 - 80, 240 + 80, 4):
                largura_linha = 3 if x % 3 == 0 else 1
                self.preview.create_line(x, 285, x, 300, fill="black", width=largura_linha)

    def _construir_interface(self):
        self.coluna_esquerda = ctk.CTkFrame(self.parent_frame)
        self.coluna_esquerda.pack(side="left", fill="both", expand=True, padx=5, pady=10)

        self.coluna_direita = ctk.CTkFrame(self.parent_frame)
        self.coluna_direita.pack(side="right", fill="both", padx=5, pady=10)

        self.labels = []

        # Grupo Config
        self.grupo_config = ctk.CTkFrame(self.coluna_esquerda)
        self.grupo_config.pack(fill="x", pady=4)
        lbl1 = ctk.CTkLabel(self.grupo_config, text="Informações de Controle", font=("Segoe UI", 12, "bold"))
        lbl1.grid(row=0, column=0, columnspan=8, pady=4, sticky="w")
        self.labels.append(lbl1)

        lbl2 = ctk.CTkLabel(self.grupo_config, text="Depósito:")
        lbl2.grid(row=1, column=0, pady=4, sticky="w")
        self.labels.append(lbl2)
        
        self.deposito_var = ctk.CTkComboBox(self.grupo_config, values=["CL03", "CL04", "CS06"], width=100, command=self.atualizar_preview)
        self.deposito_var.set("CL03")
        self.deposito_var.grid(row=1, column=1, padx=6)

        lbl3 = ctk.CTkLabel(self.grupo_config, text="Tamanho:")
        lbl3.grid(row=1, column=2, padx=6, sticky="w")
        self.labels.append(lbl3)

        self.combo_tamanho = ctk.CTkComboBox(self.grupo_config, values=["Pequena", "Média", "Grande"], width=100, command=self.atualizar_preview)
        self.combo_tamanho.set("Grande")
        self.combo_tamanho.grid(row=1, column=3, padx=6)

        lbl4 = ctk.CTkLabel(self.grupo_config, text="Nº Lote:")
        lbl4.grid(row=1, column=4, padx=6, sticky="w")
        self.labels.append(lbl4)

        self.campo_lote = ctk.CTkEntry(self.grupo_config, width=120)
        self.campo_lote.grid(row=1, column=5)
        self.campo_lote.bind("<KeyRelease>", self.atualizar_preview)

        self.btn_buscar_lote = ctk.CTkButton(self.grupo_config, text="Buscar Lote", command=self.buscar_lote, width=100)
        self.btn_buscar_lote.grid(row=1, column=6, padx=(6, 2))

        self.btn_planilha_lotes = ctk.CTkButton(self.grupo_config, text="...", command=self.selecionar_planilha_lotes, width=30)
        self.btn_planilha_lotes.grid(row=1, column=7, padx=(0, 6))

        lbl5 = ctk.CTkLabel(self.grupo_config, text="QR Topo Dir:")
        lbl5.grid(row=2, column=0, sticky="w", pady=6)
        self.labels.append(lbl5)

        self.campo_qr_topo = ctk.CTkEntry(self.grupo_config, width=380)
        self.campo_qr_topo.grid(row=2, column=1, columnspan=7, sticky="w", padx=6, pady=6)
        self.campo_qr_topo.bind("<KeyRelease>", self.atualizar_preview)


        # Grupo Prod
        self.grupo_prod = ctk.CTkFrame(self.coluna_esquerda)
        self.grupo_prod.pack(fill="x", pady=4)
        lbl6 = ctk.CTkLabel(self.grupo_prod, text="Identificação do Material", font=("Segoe UI", 12, "bold"))
        lbl6.grid(row=0, column=0, columnspan=5, pady=4, sticky="w")
        self.labels.append(lbl6)

        lbl7 = ctk.CTkLabel(self.grupo_prod, text="Cód. Material:")
        lbl7.grid(row=1, column=0, sticky="w", pady=4)
        self.labels.append(lbl7)

        self.campo_material = ctk.CTkEntry(self.grupo_prod, width=140)
        self.campo_material.grid(row=1, column=1, sticky="w", pady=4, padx=6)
        self.campo_material.bind("<KeyRelease>", self.filtrar_sugestoes)
        self.campo_material.bind("<FocusOut>", self.verificar_fechar)
        self.campo_material.bind("<Return>", self.buscar_material)

        self.btn_atualizar_banco = ctk.CTkButton(self.grupo_prod, text="🔄 Atualizar Banco", command=self.forcar_atualizacao_banco, width=130)
        self.btn_atualizar_banco.grid(row=1, column=2, padx=10, sticky="w")

        self.btn_gerar_doc = ctk.CTkButton(self.grupo_prod, text="Gerar Doc Google", command=self.gerar_doc_google, width=130)
        self.btn_gerar_doc.grid(row=1, column=3, padx=(10, 2), sticky="w")

        self.btn_config_google = ctk.CTkButton(self.grupo_prod, text="URL", command=self.configurar_google_script_url, width=40)
        self.btn_config_google.grid(row=1, column=4, padx=(0, 10), sticky="w")

        lbl8 = ctk.CTkLabel(self.grupo_prod, text="Descrição:")
        lbl8.grid(row=2, column=0, sticky="w", pady=4)
        self.labels.append(lbl8)

        self.campo_descricao = ctk.CTkEntry(self.grupo_prod, width=480)
        self.campo_descricao.grid(row=2, column=1, columnspan=4, sticky="w", pady=4, padx=6)
        self.campo_descricao.bind("<KeyRelease>", self.atualizar_preview)

        # Grupo Inv
        self.grupo_inv = ctk.CTkFrame(self.coluna_esquerda)
        self.grupo_inv.pack(fill="x", pady=4)
        lbl9 = ctk.CTkLabel(self.grupo_inv, text="Dados de Inventário", font=("Segoe UI", 12, "bold"))
        lbl9.grid(row=0, column=0, columnspan=4, pady=4, sticky="w")
        self.labels.append(lbl9)

        lbl10 = ctk.CTkLabel(self.grupo_inv, text="Qtd:")
        lbl10.grid(row=1, column=0, pady=4, sticky="w")
        self.labels.append(lbl10)

        self.campo_qtd = ctk.CTkEntry(self.grupo_inv, width=120)
        self.campo_qtd.grid(row=1, column=1, padx=6)
        self.campo_qtd.bind("<KeyRelease>", self.atualizar_preview)

        lbl11 = ctk.CTkLabel(self.grupo_inv, text="Unidade:")
        lbl11.grid(row=1, column=2, padx=6, sticky="w")
        self.labels.append(lbl11)

        self.combo_unidade = ctk.CTkComboBox(self.grupo_inv, values=["PÇ", "PÇS", "M", "CE", "KG"], width=100, command=self.atualizar_preview)
        self.combo_unidade.set("PÇS")
        self.combo_unidade.grid(row=1, column=3, padx=6)

        # Frame opções
        self.frame_opcoes = ctk.CTkFrame(self.coluna_esquerda)
        self.frame_opcoes.pack(fill="x", pady=10)

        self.check_abrir_var = tk.BooleanVar(value=True)
        self.check_abrir = ctk.CTkCheckBox(self.frame_opcoes, text="👁️ Abrir PDF criado", variable=self.check_abrir_var)
        self.check_abrir.pack(side="left", padx=10)

        self.check_imprimir_var = tk.BooleanVar(value=False)
        self.check_imprimir = ctk.CTkCheckBox(self.frame_opcoes, text="🖨️ Mandar direto para impressora", variable=self.check_imprimir_var)
        self.check_imprimir.pack(side="left", padx=10)

        self.btn_gerar = ctk.CTkButton(self.coluna_esquerda, text="🏷️ GERAR ETIQUETA", command=self.executar_geracao, font=("Segoe UI", 14, "bold"), height=45)
        self.btn_gerar.pack(fill="x", pady=15)

        self.grupo_preview = ctk.CTkFrame(self.coluna_direita)
        self.grupo_preview.pack(fill="both", expand=True)

        lbl12 = ctk.CTkLabel(self.grupo_preview, text="Miniatura de Impressão em Tempo Real", font=("Segoe UI", 12, "bold"))
        lbl12.pack(anchor="w", pady=5)
        self.labels.append(lbl12)

        self.preview = tk.Canvas(self.grupo_preview, width=480, height=320, bg="#edf2f7", bd=0, highlightthickness=0)
        self.preview.pack(fill="both", expand=True, padx=5, pady=5)

    def _obter_coluna_descricao(self):
        for coluna in self.banco.columns:
            if str(coluna).strip().lower() in ["descricao", "descrição"]: return coluna
        return None

    def _validar_dados(self, deposito):
        if not deposito: return "Selecione o depósito."
        return None

    def _atualizar_status(self, texto):
        if self.label_status: self.label_status.configure(text=texto)

def construir_aba_etiquetas(parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status=None):
    return AbaEtiquetas(parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status)
