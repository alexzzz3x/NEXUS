import json
import os
import shutil
import subprocess
import threading
import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

import pythoncom
import win32com.client
from utils_os import obter_diretorio_executavel

DEFAULT_PALETAS = {
    "Nexus": {"bg": "#090d16", "container": "#111827", "texto": "#f8fafc", "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"},
    "Azul Delray": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"},
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"}
}

_sap_ocupado = False
_login_em_andamento = False
_processo_sap_ativo = None
_abortar_solicitado = False

def is_sap_ocupado():
    return _sap_ocupado

def set_sap_ocupado(valor: bool):
    global _sap_ocupado, _abortar_solicitado
    _sap_ocupado = valor
    if valor:
        _abortar_solicitado = False

def is_abortar_solicitado():
    return _abortar_solicitado

def registrar_processo_sap(proc):
    global _processo_sap_ativo
    _processo_sap_ativo = proc

def limpar_processo_sap():
    global _processo_sap_ativo
    _processo_sap_ativo = None

def abortar_automacao_sap():
    """Interrompe qualquer automação SAP em andamento, matando processos filhos e liberando o estado."""
    global _abortar_solicitado, _processo_sap_ativo, _sap_ocupado
    _abortar_solicitado = True
    if _processo_sap_ativo is not None:
        try:
            _processo_sap_ativo.terminate()
        except Exception:
            pass
        _processo_sap_ativo = None

    try:
        creationflags = 0x08000000 if os.name == "nt" else 0
        subprocess.run(["taskkill", "/F", "/IM", "cscript.exe"], capture_output=True, creationflags=creationflags)
    except Exception:
        pass

    _sap_ocupado = False


def obter_caminho_saplogon():
    candidatos = [
        r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    no_path = shutil.which("saplogon.exe")
    if no_path:
        return no_path
    return None

def obter_sessao_ativa_sap():
    """Busca a sessão ativa do SAP GUI de forma resiliente."""
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not SapGuiAuto:
            return None
        app = getattr(SapGuiAuto, "GetScriptingEngine", None)
        if callable(app):
            app = app()
        if not app:
            return None

        # 1. Tenta ActiveSession direto do Application
        try:
            act = getattr(app, "ActiveSession", None)
            if act:
                return act
        except Exception:
            pass

        # 2. Varre todas as conexões e sessões
        try:
            conns_count = getattr(app.Children, "Count", 0)
            for i in range(conns_count):
                conn = app.Children(i)
                if not conn:
                    continue
                # Tenta ActiveSession da conexão
                try:
                    conn_act = getattr(conn, "ActiveSession", None)
                    if conn_act:
                        return conn_act
                except Exception:
                    pass

                # Varre filhos da conexão
                sess_count = getattr(conn.Children, "Count", 0)
                for j in range(sess_count):
                    sess = conn.Children(j)
                    if sess:
                        return sess
        except Exception:
            pass

        return None
    except Exception:
        return None

def verificar_sap_aberto():
    """Verifica se o SAP GUI está aberto e acessível (VBScript 32/64 bits + COM + Processo)."""
    # 1. Checagem via cscript nativo (conecta direto na ROT do SAP GUI, testando SysWOW64 e System32)
    try:
        import tempfile
        vbs = 'On Error Resume Next: Set sap = GetObject("SAPGUI"): If Err.Number <> 0 Then WScript.Quit 1 Else WScript.Quit 0'
        fd, path = tempfile.mkstemp(prefix="chk_sap_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="cp1252") as f:
            f.write(vbs)
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        
        # Testa cscript (padrão do sistema primeiro, depois System32 e SysWOW64)
        cscripts = ["cscript.exe"]
        cscript_sys32 = r"C:\Windows\System32\cscript.exe"
        cscript_32 = r"C:\Windows\SysWOW64\cscript.exe"
        if os.path.exists(cscript_sys32) and cscript_sys32 not in cscripts:
            cscripts.append(cscript_sys32)
        if os.path.exists(cscript_32) and cscript_32 not in cscripts:
            cscripts.append(cscript_32)

        for cs in cscripts:
            try:
                res = subprocess.run([cs, "//nologo", path], timeout=3, creationflags=creationflags)
                if res.returncode == 0:
                    try: os.remove(path)
                    except OSError: pass
                    return True
            except Exception:
                continue

        try: os.remove(path)
        except OSError: pass
    except Exception:
        pass

    # 2. Tenta via win32com direto
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if SapGuiAuto:
            return True
    except Exception:
        pass

    # 3. Fallback via verificação de processos ativos do SAP (sapgui.exe ou saplogon.exe)
    try:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        out = subprocess.check_output(
            'tasklist /FI "IMAGENAME eq sap*"',
            shell=True,
            text=True,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags
        )
        out_lower = out.lower()
        if "sapgui.exe" in out_lower or "saplogon.exe" in out_lower:
            return True
    except Exception:
        pass

    return False

ARQUIVO_LOGINS = os.path.join(obter_diretorio_executavel(), "data", "logins_sap.json")
logins_salvos = {}

def carregar_logins():
    global logins_salvos
    LOGINS_PADRAO = {
        "Adriana": {"usuario": "BR0425314228", "senha": "Cnt*11092027"},
        "Junior": {"usuario": "BR0287830298", "senha": "Cnt*11092027"},
        "Pedro": {"usuario": "BR0498361748", "senha": "Cnt*11092027"},
    }
    if os.path.exists(ARQUIVO_LOGINS):
        try:
            with open(ARQUIVO_LOGINS, "r", encoding="utf-8") as f:
                logins_salvos = json.load(f)
        except Exception:
            logins_salvos = {}
        if not logins_salvos:
            logins_salvos = dict(LOGINS_PADRAO)
            salvar_logins()
    else:
        logins_salvos = dict(LOGINS_PADRAO)
        salvar_logins()

def salvar_logins():
    try:
        os.makedirs(os.path.dirname(ARQUIVO_LOGINS), exist_ok=True)
        with open(ARQUIVO_LOGINS, "w", encoding="utf-8") as f:
            json.dump(logins_salvos, f, indent=4, ensure_ascii=False)
    except Exception as e:
        pass

def obter_info_sap_atual():
    if _sap_ocupado:
        return "SAP: ocupado..."
    
    # 1. Tenta via COM Python primeiro
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if SapGuiAuto:
            app = getattr(SapGuiAuto, "GetScriptingEngine", None)
            if callable(app):
                app = app()
            if app:
                session = obter_sessao_ativa_sap()
                if session:
                    info = session.Info
                    sistema = getattr(info, "SystemName", "") or "SAP"
                    cliente = getattr(info, "Client", "") or "-"
                    usuario = getattr(info, "User", "") or "-"
                    transacao = getattr(info, "Transaction", "") or "-"
                    return f"SAP: {sistema} | Cliente {cliente} | Usuário {usuario} | Tx {transacao}"
                else:
                    cnt = getattr(app.Children, "Count", 0)
                    if cnt > 0:
                        return f"SAP: conectado ({cnt} conexões - aguardando login/sessão)"
                    return "SAP: aberto (sem sessão ativa)"
    except Exception:
        pass

    # 2. Fallback via VBScript nativo (conecta com a ROT do SAP GUI)
    try:
        import tempfile
        vbs = '''On Error Resume Next
Set sap = GetObject("SAPGUI")
If Err.Number <> 0 Or sap Is Nothing Then WScript.Quit 1
Set app = sap.GetScriptingEngine
If Err.Number <> 0 Or app Is Nothing Then WScript.Quit 2
If app.Children.Count = 0 Then WScript.Quit 3
Set conn = app.Children(0)
If conn.Children.Count = 0 Then WScript.Quit 4
Set sess = conn.Children(0)
WScript.Echo sess.Info.SystemName & "|" & sess.Info.Client & "|" & sess.Info.User & "|" & sess.Info.Transaction
'''
        fd, path = tempfile.mkstemp(prefix="info_sap_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="cp1252") as f:
            f.write(vbs)
        
        cscripts = ["cscript.exe"]
        cscript_sys32 = r"C:\Windows\System32\cscript.exe"
        cscript_32 = r"C:\Windows\SysWOW64\cscript.exe"
        if os.path.exists(cscript_sys32) and cscript_sys32 not in cscripts:
            cscripts.append(cscript_sys32)
        if os.path.exists(cscript_32) and cscript_32 not in cscripts:
            cscripts.append(cscript_32)

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        saida_vbs = ""
        for cs in cscripts:
            try:
                res = subprocess.run([cs, "//nologo", path], capture_output=True, text=True, timeout=3, creationflags=creationflags)
                if res.returncode == 0 and res.stdout.strip():
                    saida_vbs = res.stdout.strip()
                    break
            except Exception:
                continue

        try: os.remove(path)
        except OSError: pass

        if saida_vbs and "|" in saida_vbs:
            parts = saida_vbs.split("|")
            if len(parts) >= 4:
                return f"SAP: {parts[0] or 'SAP'} | Cliente {parts[1] or '-'} | Usuário {parts[2] or '-'} | Tx {parts[3] or '-'}"
    except Exception:
        pass

    # 3. Se não tem sessão ativa, verifica se o processo saplogon ou sapgui está aberto
    try:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        out = subprocess.check_output(
            'tasklist /FI "IMAGENAME eq sap*"',
            shell=True,
            text=True,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags
        )
        out_lower = out.lower()
        if "sapgui.exe" in out_lower:
            return "SAP: aberto (sessão em carregamento)"
        elif "saplogon.exe" in out_lower:
            return "SAP: Logon aberto (aguardando conexão/login)"
    except Exception:
        pass

    return "SAP: desconectado"

def consultar_info_sap_async(callback, widget_ref=None):
    """Executa a consulta do SAP em uma thread em background e retorna o resultado para a thread principal de forma segura."""
    if _sap_ocupado:
        return

    def _worker():
        try:
            resultado = obter_info_sap_atual()
        except Exception:
            resultado = "SAP: desconectado"

        if widget_ref:
            try:
                widget_ref.after(0, lambda: callback(resultado))
            except Exception:
                pass
        else:
            try:
                callback(resultado)
            except Exception:
                pass

    threading.Thread(target=_worker, daemon=True).start()

def _atualizar_ui_status(widget, texto):
    if widget:
        try:
            widget.after(0, lambda: widget.configure(text=texto))
        except Exception:
            pass

def _mostrar_msgbox_async(widget, tipo, titulo, mensagem):
    def _exibir():
        try:
            if tipo == "erro":
                messagebox.showerror(titulo, mensagem)
            elif tipo == "aviso":
                messagebox.showwarning(titulo, mensagem)
            else:
                messagebox.showinfo(titulo, mensagem)
        except Exception:
            pass
    if widget:
        try:
            widget.after(0, _exibir)
            return
        except Exception:
            pass
    _exibir()

def logar_sap(usuario, senha, nome, label_status=None, on_concluido=None):
    global _sap_ocupado, _login_em_andamento
    if _login_em_andamento:
        _mostrar_msgbox_async(label_status, "aviso", "Aviso", "Já existe um processo de login SAP em andamento. Aguarde.")
        return

    _login_em_andamento = True
    _sap_ocupado = True

    try:
        pythoncom.CoInitialize()
        _atualizar_ui_status(label_status, f" Verificando SAP GUI para {nome}...")

        SapGuiAuto = None
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            pass

        if not SapGuiAuto:
            caminho_sap = obter_caminho_saplogon()
            if not caminho_sap:
                _mostrar_msgbox_async(
                    label_status, "erro", "Erro",
                    "Executável do SAP Logon (saplogon.exe) não foi encontrado nos caminhos padrão."
                )
                _atualizar_ui_status(label_status, " Erro: SAP Logon não encontrado.")
                return

            _atualizar_ui_status(label_status, " Abrindo SAP Logon...")
            subprocess.Popen(caminho_sap)
            tentativas = 0
            while tentativas < 12 and not SapGuiAuto:
                time.sleep(1)
                try:
                    pythoncom.CoInitialize()
                    SapGuiAuto = win32com.client.GetObject("SAPGUI")
                except Exception:
                    tentativas += 1

            if not SapGuiAuto:
                _mostrar_msgbox_async(label_status, "erro", "Erro", "O SAP demorou para abrir. Tente novamente.")
                _atualizar_ui_status(label_status, " Erro: Tempo limite para abrir o SAP expirou.")
                return

        _atualizar_ui_status(label_status, " Conectando ao ambiente SAP...")
        app = SapGuiAuto.GetScriptingEngine
        connection = None
        try:
            connection = app.OpenConnection("Enel_SP_RP1_CCS_Produtivo [VPN Palo Alto]", True)
        except Exception:
            try:
                if app.Children.Count > 0:
                    connection = app.Children(0)
            except Exception:
                pass

        if not connection:
            _mostrar_msgbox_async(label_status, "erro", "Erro", "Não foi possível abrir a conexão com o SAP.")
            _atualizar_ui_status(label_status, " Erro ao abrir conexão SAP.")
            return

        time.sleep(1)
        if connection.Children.Count == 0:
            _mostrar_msgbox_async(label_status, "erro", "Erro", "Nenhuma sessão disponível na conexão SAP.")
            _atualizar_ui_status(label_status, " Erro: Sessão SAP indisponível.")
            return

        session = connection.Children(0)
        _atualizar_ui_status(label_status, f" Autenticando usuário {nome}...")

        try:
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = ""
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = ""
        except Exception:
            pass

        try:
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha
            session.findById("wnd[0]").sendVKey(0)
        except Exception:
            try:
                session.findById("wnd[0]/usr/txtRSYST-BNAME[0,0]").Text = usuario
                session.findById("wnd[0]/usr/pwdRSYST-BCODE[0,0]").Text = senha
                session.findById("wnd[0]").sendVKey(0)
            except Exception as e:
                _mostrar_msgbox_async(
                    label_status, "erro", "Erro",
                    f"Erro ao preencher dados de login. O SAP está na tela inicial?\n\nDetalhe: {str(e)}"
                )
                _atualizar_ui_status(label_status, " Erro ao preencher credenciais SAP.")
                return

        _atualizar_ui_status(label_status, f" Usuário SAP logado: {nome} | {usuario}")
    except Exception as e:
        _mostrar_msgbox_async(label_status, "erro", "Erro Inesperado", f"Não foi possível integrar com o SAP:\n{str(e)}")
        _atualizar_ui_status(label_status, " Falha no Login SAP.")
    finally:
        _sap_ocupado = False
        _login_em_andamento = False
        if on_concluido:
            try:
                if label_status:
                    label_status.after(0, on_concluido)
                else:
                    on_concluido()
            except Exception:
                pass

def logout_sap(label_status=None):
    global _sap_ocupado
    _sap_ocupado = True
    def _worker():
        global _sap_ocupado
        try:
            pythoncom.CoInitialize()
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            connection = app.Children(0)
            session = connection.Children(0)
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
            session.findById("wnd[0]").sendVKey(0)
            _atualizar_ui_status(label_status, " SAP desconectado.")
        except Exception:
            _mostrar_msgbox_async(label_status, "info", "Aviso", "O SAP já está fechado ou inativo.")
        finally:
            _sap_ocupado = False

    threading.Thread(target=_worker, daemon=True).start()

class AbaSap:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.botoes_usuario = []
        self._conectando = False

        carregar_logins()
        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        self.entry_busca.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)
        self.btn_gerenciar.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_logout.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        
        self.frame_topo.configure(fg_color="transparent")
        self.lbl_busca.configure(text_color=c["texto"])

        self.atualizar_botoes(self.entry_busca.get())

    def atualizar_botoes(self, busca=""):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.botoes_usuario = []
        c = self._cores()

        for nome in sorted(logins_salvos.keys()):
            if busca.lower() not in nome.lower():
                continue
            dados = logins_salvos[nome]
            btn = ctk.CTkButton(
                self.scrollable_frame, text=f"👤 Conectar como {nome}", font=("Segoe UI", 13, "bold"),
                fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"], height=45, corner_radius=6,
                command=lambda n=nome, u=dados["usuario"], s=dados["senha"]: self.iniciar_login(u, s, n),
            )
            btn.pack(fill="x", pady=5)
            self.botoes_usuario.append(btn)

    def iniciar_login(self, u, s, n):
        if self._conectando or _login_em_andamento:
            messagebox.showwarning("Aviso", "Já existe um login em andamento. Aguarde.", parent=self.parent_frame.winfo_toplevel())
            return

        self._conectando = True
        for btn in self.botoes_usuario:
            btn.configure(state="disabled")

        def on_concluido():
            self._conectando = False
            for btn in self.botoes_usuario:
                try:
                    btn.configure(state="normal")
                except Exception:
                    pass

        threading.Thread(
            target=logar_sap,
            args=(u, s, n, self.label_status_global, on_concluido),
            daemon=True
        ).start()

    def abrir_gerenciador(self):
        c = self._cores()
        pop = ctk.CTkToplevel()
        pop.title("Gerenciar Usuários SAP")
        pop.geometry("380x480")
        pop.resizable(False, False)
        pop.grab_set()
        pop.configure(fg_color=c.get("bg"))

        ctk.CTkLabel(pop, text="Usuários cadastrados:", font=("Segoe UI", 12, "bold"), text_color=c["texto"]).pack(anchor="w", padx=20, pady=(15, 5))

        listbox_pop = tk.Listbox(pop, height=6, font=("Segoe UI", 11), bd=0, relief="flat", bg=c["input"], fg=c["texto"], selectbackground=c["accent"], selectforeground="white", highlightthickness=1, highlightbackground=c["container"])
        listbox_pop.pack(fill="x", padx=20)

        frame_form = ctk.CTkFrame(pop, fg_color="transparent")
        frame_form.pack(fill="both", expand=True, padx=20, pady=10)

        entry_nome = self._campo_popup(frame_form, "Nome:", c)
        entry_user = self._campo_popup(frame_form, "Usuário SAP:", c)
        entry_senha = self._campo_popup(frame_form, "Senha SAP:", c, show="*")

        def atualizar_lista_pop():
            listbox_pop.delete(0, tk.END)
            for n in sorted(logins_salvos.keys()):
                listbox_pop.insert(tk.END, n)

        def selecionar_pop(event):
            if not listbox_pop.curselection():
                return
            nome_sel = listbox_pop.get(listbox_pop.curselection())
            dados = logins_salvos.get(nome_sel)
            if dados:
                entry_nome.delete(0, tk.END)
                entry_nome.insert(0, nome_sel)
                entry_user.delete(0, tk.END)
                entry_user.insert(0, dados["usuario"])
                entry_senha.delete(0, tk.END)
                entry_senha.insert(0, dados["senha"])

        def btn_salvar():
            n = entry_nome.get().strip()
            u = entry_user.get().strip()
            s = entry_senha.get().strip()
            if not n or not u or not s:
                messagebox.showwarning("Aviso", "Preencha Nome, Usuário e Senha.", parent=pop)
                return
            logins_salvos[n] = {"usuario": u, "senha": s}
            salvar_logins()
            atualizar_lista_pop()
            self.atualizar_botoes(self.entry_busca.get())
            messagebox.showinfo("Sucesso", f"Usuário {n} salvo com sucesso!", parent=pop)

        def btn_excluir():
            n = entry_nome.get().strip()
            if n in logins_salvos:
                del logins_salvos[n]
                salvar_logins()
                atualizar_lista_pop()
                self.atualizar_botoes(self.entry_busca.get())
                for entry in [entry_nome, entry_user, entry_senha]:
                    entry.delete(0, tk.END)
                messagebox.showinfo("Sucesso", "Usuário excluído.", parent=pop)

        listbox_pop.bind("<<ListboxSelect>>", selecionar_pop)
        atualizar_lista_pop()

        btn_salvar_ui = ctk.CTkButton(frame_form, text="Salvar / Adicionar", command=btn_salvar, font=("Segoe UI", 12, "bold"), fg_color=c["accent"], hover_color=c["hover"], text_color="white", height=35)
        btn_salvar_ui.pack(fill="x", pady=(12, 4))
        btn_excluir_ui = ctk.CTkButton(frame_form, text="Excluir selecionado", command=btn_excluir, font=("Segoe UI", 12, "bold"), fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"], height=35)
        btn_excluir_ui.pack(fill="x", pady=4)

    def _construir_interface(self):
        self.frame_topo = ctk.CTkFrame(self.parent_frame)
        self.frame_topo.pack(fill="x", padx=40, pady=(30, 10))

        self.lbl_busca = ctk.CTkLabel(self.frame_topo, text="Buscar:", font=("Segoe UI", 12, "bold"))
        self.lbl_busca.pack(side="left")

        self.entry_busca = ctk.CTkEntry(self.frame_topo, font=("Segoe UI", 12))
        self.entry_busca.pack(side="left", fill="x", expand=True, padx=10)

        self.btn_gerenciar = ctk.CTkButton(self.frame_topo, text="⚙️ Gerenciar Usuários", command=self.abrir_gerenciador, font=("Segoe UI", 11, "bold"), width=150, height=30)
        self.btn_gerenciar.pack(side="right")

        self.scrollable_frame = ctk.CTkScrollableFrame(self.parent_frame, corner_radius=8)
        self.scrollable_frame.pack(fill="both", expand=True, padx=40, pady=10)

        def on_key_release(*args):
            self.atualizar_botoes(self.entry_busca.get())
        self.entry_busca.bind("<KeyRelease>", on_key_release)

        self.btn_logout = ctk.CTkButton(self.parent_frame, text="⏻ Encerrar Sessão (Logout)", command=lambda: logout_sap(self.label_status_global), font=("Segoe UI", 13, "bold"), height=45)
        self.btn_logout.pack(fill="x", padx=40, pady=(10, 30))

    def _campo_popup(self, parent, texto, cores, show=""):
        ctk.CTkLabel(parent, text=texto, text_color=cores["texto"], font=("Segoe UI", 11, "bold")).pack(anchor="w")
        entry = ctk.CTkEntry(parent, font=("Segoe UI", 12), show=show, fg_color=cores["input"], text_color=cores["texto"], border_color=cores["container"])
        entry.pack(fill="x", pady=(0, 8))
        return entry

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

def construir_aba_sap(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaSap(parent_frame, label_status_global, paletas, obter_tema_atual)
