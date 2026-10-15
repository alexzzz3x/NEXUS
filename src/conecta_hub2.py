import os
import sys
import threading

from utils_os import garantir_instancia_unica, obter_caminho_recurso, obter_diretorio_executavel

# Garante que apenas uma instância do Conecta Hub rode por vez quando executado como script principal
if __name__ == "__main__":
    _ok_instancia, _mutex_app = garantir_instancia_unica("Nexus_SAP_Hub_SingleInstance_Mutex")
    if not _ok_instancia:
        sys.exit(0)
    # Limpa processos cscript.exe órfãos deixados por travamentos ou encerramentos forçados anteriores
    try:
        import subprocess
        subprocess.run(["taskkill", "/F", "/IM", "cscript.exe"], capture_output=True)
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

from utils_ui import show_toast, play_sound, run_async
from utils_logger import log_error, log_info
from sidebar_navegacao import SidebarNavegacao
from utils_updater import iniciar_checagem_atualizacao, checar_atualizacao_manual, VERSAO_ATUAL

def global_exception_handler(exc_type, exc_value, exc_traceback):
    log_error(f"Uncaught exception: {exc_value}", exc_info=(exc_type, exc_value, exc_traceback))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

sys.excepthook = global_exception_handler

def tk_callback_exception_handler(self, exc_type, exc_value, exc_traceback):
    log_error(f"Uncaught Tk callback exception: {exc_value}", exc_info=(exc_type, exc_value, exc_traceback))

tk.Tk.report_callback_exception = tk_callback_exception_handler

ctk.set_default_color_theme("blue")  # Tema padrao do CTK

CONFIG_FILE = os.path.join(obter_diretorio_executavel(), "data", "config_tema_hub.txt")

def obter_arquivo_banco():
    caminho_exe = os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx")
    if os.path.exists(caminho_exe):
        return caminho_exe
    return obter_caminho_recurso("data", "banco_de_dados.xlsx")

ARQUIVO_BANCO = obter_arquivo_banco()

PALETAS = {
    "Nexus": {"bg": "#090d16", "container": "#111827", "texto": "#f8fafc", "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436", "scrollbar": "#1e293b", "scrollbar_hover": "#00d2ff", "scrollbar_trough": "#090d16"},
    "Azul Delray": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60", "scrollbar": "#5f7fa8", "scrollbar_hover": "#7fa2cf", "scrollbar_trough": "#10233d"},
    "Dark Mode": {"bg": "#1a202c", "container": "#2d3748", "texto": "#ffffff", "accent": "#319795", "hover": "#2c7a7b", "input": "#3a4454", "scrollbar": "#64748b", "scrollbar_hover": "#94a3b8", "scrollbar_trough": "#202938"},
    "Light Clean": {"bg": "#ffffff", "container": "#f8f9fa", "texto": "#2d3748", "accent": "#319795", "hover": "#2c7a7b", "input": "#edf2f7", "scrollbar": "#94a3b8", "scrollbar_hover": "#64748b", "scrollbar_trough": "#e2e8f0"},
    "Corporate Slate": {"bg": "#111827", "container": "#1f2937", "texto": "#f9fafb", "accent": "#2563eb", "hover": "#1d4ed8", "input": "#374151", "scrollbar": "#6b7280", "scrollbar_hover": "#9ca3af", "scrollbar_trough": "#172033"},
    "Office Light": {"bg": "#f3f6fb", "container": "#ffffff", "texto": "#1f2937", "accent": "#2563eb", "hover": "#1d4ed8", "input": "#e8eef7", "scrollbar": "#a7b4c8", "scrollbar_hover": "#6b7f99", "scrollbar_trough": "#dce5f2"},
    "Emerald Pro": {"bg": "#0f1f1b", "container": "#18352e", "texto": "#f8fafc", "accent": "#10b981", "hover": "#059669", "input": "#21483e", "scrollbar": "#6ba99a", "scrollbar_hover": "#8fd4c3", "scrollbar_trough": "#132b25"},
    "Graphite Orange": {"bg": "#171717", "container": "#262626", "texto": "#fafafa", "accent": "#f97316", "hover": "#ea580c", "input": "#3f3f46", "scrollbar": "#737373", "scrollbar_hover": "#a3a3a3", "scrollbar_trough": "#202020"},
    "Royal Indigo": {"bg": "#17152b", "container": "#262447", "texto": "#ffffff", "accent": "#8b5cf6", "hover": "#7c3aed", "input": "#34315f", "scrollbar": "#8179bd", "scrollbar_hover": "#a69cf0", "scrollbar_trough": "#201d3b"},
}
# Compatibilidade transparente para módulos legados
PALETAS["Conecta Blue"] = PALETAS["Azul Delray"]

etiquetas_view = None
sap_view = None
bandeirada_view = None
extrator_view = None
osme_view = None
conferencia_view = None
zm02_view = None
transferenciaosme_view = None
mudanca_recurso_view = None
dashboard_view = None
fichas_view = None
plataforma_view = None
consultar_migo_view = None
saida_osme_view = None
contagem_carro_view = None
atualizar_osme_view = None
contagem_deposito_view = None
sidebar_navegacao = None

def carregar_tema_salvo():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                tema = f.read().strip()
                if tema == "Conecta Blue":
                    return "Azul Delray"
                if tema in PALETAS:
                    return tema
        except Exception:
            pass
    return "Nexus"

def salvar_tema_atual(tema_nome):
    try:
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(tema_nome)
    except Exception:
        pass

def obter_tema_atual():
    return combo_temas.get()

_icon_handles = []

def aplicar_icone_janela(janela):
    global _icon_handles
    try:
        caminho_ico = obter_caminho_recurso("assets", "Nexus.ico")
        if not os.path.exists(caminho_ico):
            caminho_ico = obter_caminho_recurso("assets", "eclipse.ico")
        
        if not os.path.exists(caminho_ico):
            return

        janela.iconbitmap(caminho_ico)
        
        # Carrega resoluções via iconphoto do Tkinter
        try:
            from PIL import Image, ImageTk
            photo_imgs = []
            for s in [256, 64, 48, 32, 24, 16]:
                caminho_frame = obter_caminho_recurso("assets", f"sharp_frame_{s}.png")
                if os.path.exists(caminho_frame):
                    photo_imgs.append(ImageTk.PhotoImage(Image.open(caminho_frame)))
            if photo_imgs:
                janela.iconphoto(True, *photo_imgs)
                janela._icon_photos_ref = photo_imgs
        except Exception:
            pass

        # Garante nitidez máxima na barra de tarefas e título do Windows via WM_SETICON nativo
        if os.name == "nt":
            import ctypes
            hwnd = janela.winfo_id()
            LR_LOADFROMFILE = 0x0010
            IMAGE_ICON = 1
            WM_SETICON = 0x80
            ICON_SMALL = 0
            ICON_BIG = 1
            SM_CXICON = 11
            SM_CYICON = 12
            SM_CXSMICON = 49
            SM_CYSMICON = 50

            cx_big = ctypes.windll.user32.GetSystemMetrics(SM_CXICON)
            cy_big = ctypes.windll.user32.GetSystemMetrics(SM_CYICON)
            cx_sm = ctypes.windll.user32.GetSystemMetrics(SM_CXSMICON)
            cy_sm = ctypes.windll.user32.GetSystemMetrics(SM_CYSMICON)

            h_big = ctypes.windll.user32.LoadImageW(None, os.path.abspath(caminho_ico), IMAGE_ICON, cx_big, cy_big, LR_LOADFROMFILE)
            h_sm = ctypes.windll.user32.LoadImageW(None, os.path.abspath(caminho_ico), IMAGE_ICON, cx_sm, cy_sm, LR_LOADFROMFILE)

            hwnd_root = ctypes.windll.user32.GetAncestor(hwnd, 2) or hwnd
            hwnd_parent = ctypes.windll.user32.GetParent(hwnd) or hwnd
            targets = set([hwnd, hwnd_root, hwnd_parent])

            if h_big:
                for target in targets:
                    ctypes.windll.user32.SendMessageW(target, WM_SETICON, ICON_BIG, h_big)
                _icon_handles.append(h_big)
            if h_sm:
                for target in targets:
                    ctypes.windll.user32.SendMessageW(target, WM_SETICON, ICON_SMALL, h_sm)
                _icon_handles.append(h_sm)
    except Exception:
        pass

def aplicar_tema(tema_nome):
    c = PALETAS.get(tema_nome, PALETAS["Nexus"])
    
    # Modern CTK mode
    if "Light" in tema_nome:
        ctk.set_appearance_mode("light")
    else:
        ctk.set_appearance_mode("dark")
        
    root.configure(fg_color=c["bg"])
    salvar_tema_atual(tema_nome)

    # Legacy styling for modules that haven't been refactored yet
    scrollbar = c.get("scrollbar", c["input"])
    scrollbar_hover = c.get("scrollbar_hover", c["accent"])
    scrollbar_trough = c.get("scrollbar_trough", c["bg"])
    
    style.configure("TFrame", background=c["bg"])
    style.configure("TLabel", background=c["bg"], foreground=c["texto"])
    style.configure("TLabelframe", background=c["bg"], foreground=c["accent"])
    style.configure("TLabelframe.Label", background=c["bg"], foreground=c["accent"], font=("Segoe UI", 10, "bold"))
    style.configure("TButton", background=c["container"], foreground=c["texto"], borderwidth=0, focusthickness=0, focuscolor=c["container"], padding=(10, 6), font=("Segoe UI", 9, "bold"))
    style.map("TButton", background=[("disabled", c["input"]), ("pressed", c["accent"]), ("active", c["hover"])], foreground=[("disabled", c["texto"]), ("active", "white")])
    style.configure("TEntry", fieldbackground=c["input"], foreground=c["texto"], bordercolor=c["container"], lightcolor=c["container"], darkcolor=c["container"], padding=(6, 4))
    style.map("TEntry", bordercolor=[("focus", c["accent"])], lightcolor=[("focus", c["accent"])], darkcolor=[("focus", c["accent"])])
    style.configure("TRadiobutton", background=c["bg"], foreground=c["texto"], indicatorcolor=c["input"], focuscolor=c["bg"], font=("Segoe UI", 9))
    style.map("TRadiobutton", background=[("active", c["bg"])], foreground=[("disabled", c["scrollbar"]), ("active", c["texto"])])
    style.configure("Horizontal.TProgressbar", background=c["accent"], troughcolor=c["input"], bordercolor=c["input"], lightcolor=c["accent"], darkcolor=c["accent"], thickness=10)
    style.configure("TSeparator", background=c["container"])

    try:
        for scrollbar_style in ["Vertical.TScrollbar", "Horizontal.TScrollbar", "TScrollbar"]:
            style.configure(scrollbar_style, gripcount=0, background=scrollbar, darkcolor=scrollbar, lightcolor=scrollbar, troughcolor=scrollbar_trough, bordercolor=scrollbar_trough, arrowcolor=c["texto"], relief="flat", borderwidth=0, width=12, arrowsize=10)
            style.map(scrollbar_style, background=[("pressed", c["accent"]), ("active", scrollbar_hover)], troughcolor=[("active", scrollbar_trough)], arrowcolor=[("active", "white")])
    except Exception:
        pass

    for widget_type in ["TCombobox", "TSpinbox"]:
        style.configure(widget_type, fieldbackground=c["input"], background=c["input"], foreground=c["texto"], bordercolor=c["container"], arrowcolor=c["accent"])
        style.map(widget_type, fieldbackground=[("readonly", c["input"])], foreground=[("readonly", c["texto"])], background=[("readonly", c["input"])])

    try:
        root.option_add("*TCombobox*Listbox.background", c["input"])
        root.option_add("*TCombobox*Listbox.foreground", c["texto"])
        root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])
    except:
        pass

    frame_topo.configure(fg_color=c["bg"])
    if 'bar_status' in globals() and bar_status:
        bar_status.configure(fg_color=c["bg"])
    lbl_status.configure(text_color=c["texto"])
    lbl_sap_info.configure(text_color=c["texto"])
    lbl_creditos.configure(text_color=c["texto"])
    try:
        lbl_sub_header.configure(text_color=c.get("scrollbar_hover", "#38bdf8"))
    except Exception:
        pass

    if 'btn_chat' in globals() and btn_chat:
        btn_chat.configure(
            fg_color=c["accent"],
            hover_color=c["hover"],
            text_color="#090d16" if "Light" not in tema_nome else "#ffffff"
        )

    if etiquetas_view: etiquetas_view.aplicar_tema(tema_nome)
    if sap_view: sap_view.aplicar_tema(tema_nome)
    if bandeirada_view: bandeirada_view.aplicar_tema(tema_nome)
    if extrator_view: extrator_view.aplicar_tema(tema_nome)
    if osme_view: osme_view.aplicar_tema(tema_nome)
    if conferencia_view: conferencia_view.aplicar_tema(tema_nome)
    if transferenciaosme_view: transferenciaosme_view.aplicar_tema(tema_nome)
    if mudanca_recurso_view: mudanca_recurso_view.aplicar_tema(tema_nome)
    if dashboard_view: dashboard_view.aplicar_tema(tema_nome)
    if fichas_view: fichas_view.aplicar_tema(tema_nome)
    if plataforma_view: plataforma_view.aplicar_tema(tema_nome)
    if consultar_migo_view: consultar_migo_view.aplicar_tema(tema_nome)
    if saida_osme_view: saida_osme_view.aplicar_tema(tema_nome)
    if contagem_carro_view: contagem_carro_view.aplicar_tema(tema_nome)
    if atualizar_osme_view: atualizar_osme_view.aplicar_tema(tema_nome)
    if contagem_deposito_view: contagem_deposito_view.aplicar_tema(tema_nome)
    if sidebar_navegacao: sidebar_navegacao.aplicar_tema(tema_nome)

def atualizar_info_sap():
    """Consulta o status do SAP de forma 100% assíncrona, sem travar a interface gráfica."""
    def _aplicar(resultado):
        try:
            lbl_sap_info.configure(text=resultado)
        except Exception:
            pass
        finally:
            root.after(10000, atualizar_info_sap)

    consultar_info_sap_async(_aplicar, widget_ref=root)


# ================== MAIN UI ==================
try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("nexus.sap.nucleo.v22")
except Exception:
    pass

root = ctk.CTk()
root.title(f"N.E.X.U.S v{VERSAO_ATUAL}")
root.geometry("1160x700")
root.withdraw() # Hide root initially for splash screen
aplicar_icone_janela(root)

# Splash Screen
splash = ctk.CTkToplevel(root)
splash.overrideredirect(True)
splash_width = 500
splash_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width / 2) - (splash_width / 2)
y = (screen_height / 2) - (splash_height / 2)
splash.geometry(f'{splash_width}x{splash_height}+{int(x)}+{int(y)}')
splash.configure(fg_color="#0a192f")
splash.attributes("-topmost", True)
splash.attributes("-alpha", 0.0)

def fade_in(alpha=0.0):
    if alpha < 1.0:
        alpha += 0.08
        try:
            splash.attributes("-alpha", alpha)
            root.after(15, lambda: fade_in(alpha))
        except Exception:
            pass

if __name__ == "__main__":
    fade_in()

lbl_splash_title = ctk.CTkLabel(splash, text="N.E.X.U.S", font=("Segoe UI", 32, "bold"), text_color="#00d4ff")
lbl_splash_title.pack(expand=True, pady=(50, 0))

lbl_splash_sub = ctk.CTkLabel(splash, text="Núcleo de Estoque e Extração Unificada com o SAP", font=("Segoe UI", 11), text_color="#90cdf4")
lbl_splash_sub.pack(pady=(0, 15))

lbl_splash_loading = ctk.CTkLabel(splash, text="Carregando Módulos...", font=("Segoe UI", 12), text_color="white")
lbl_splash_loading.pack(pady=5)

splash_progress = ctk.CTkProgressBar(splash, width=300, progress_color="#f57c00")
splash_progress.pack(pady=20)
splash_progress.set(0)

style = ttk.Style()
style.theme_use("clam")
tema_inicial = carregar_tema_salvo()

# TOP FRAME
frame_topo = ctk.CTkFrame(root, height=60, corner_radius=0)
frame_topo.pack(fill=tk.X, pady=(0, 10))

caminho_img = obter_caminho_recurso("assets", "Nexus_logo.png")
if not os.path.exists(caminho_img):
    caminho_img = obter_caminho_recurso("assets", "Nexus.jpg")
if not os.path.exists(caminho_img):
    caminho_img = obter_caminho_recurso("assets", "logo_cntc.png")

try:
    from PIL import Image
    pil_img = Image.open(caminho_img)
    img_logo = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(48, 48))
    lbl_logo = ctk.CTkLabel(frame_topo, image=img_logo, text="")
except Exception:
    lbl_logo = ctk.CTkLabel(frame_topo, text="⚡", font=("Segoe UI", 20, "bold"))
lbl_logo.pack(side=tk.LEFT, padx=(15, 8), pady=6)

lbl_title = ctk.CTkLabel(frame_topo, text="N.E.X.U.S", font=("Segoe UI", 20, "bold"))
lbl_title.pack(side=tk.LEFT, padx=(0, 6))

lbl_sub_header = ctk.CTkLabel(frame_topo, text="• Núcleo de Estoque e Extração Unificada SAP", font=("Segoe UI", 12), text_color="#7fa2cf")
lbl_sub_header.pack(side=tk.LEFT, padx=(0, 10), pady=(4, 0))

combo_temas = ctk.CTkComboBox(frame_topo, values=[k for k in PALETAS.keys() if k != "Conecta Blue"], command=aplicar_tema, width=160, font=("Segoe UI", 12))
combo_temas.set(tema_inicial)
combo_temas.pack(side=tk.RIGHT, padx=20, pady=15)

btn_ver_atualizacao = ctk.CTkButton(
    frame_topo,
    text=f"🔄 v{VERSAO_ATUAL}",
    font=("Segoe UI", 11, "bold"),
    width=75,
    height=28,
    corner_radius=6,
    fg_color="#1e293b",
    hover_color="#334155",
    text_color="#38bdf8",
    command=lambda: checar_atualizacao_manual(root, PALETAS, obter_tema_atual)
)
btn_ver_atualizacao.pack(side=tk.RIGHT, padx=(0, 6), pady=15)

# STATUS BAR
bar_status = ctk.CTkFrame(root, height=30, corner_radius=0)
bar_status.pack(side="bottom", fill="x")

lbl_status = ctk.CTkLabel(bar_status, text=" Pronto | N.E.X.U.S", font=("Segoe UI", 11, "bold"))
lbl_status.pack(side="left", padx=10, pady=2)

lbl_sap_info = ctk.CTkLabel(bar_status, text="SAP: verificando...", font=("Segoe UI", 11, "bold"))
lbl_sap_info.pack(side="left", padx=20, pady=2)

lbl_creditos = ctk.CTkLabel(bar_status, text="2026 Desenvolvido por 🟢 Anderson Vieira™", font=("Segoe UI", 11, "italic"))
lbl_creditos.pack(side="right", padx=10, pady=2)

# CORPO PRINCIPAL (Sidebar à esquerda + Área de Conteúdo à direita)
frame_corpo = ctk.CTkFrame(root, corner_radius=0, fg_color="transparent")
frame_corpo.pack(side="top", fill=tk.BOTH, expand=True, padx=0, pady=0)

tab_name_home = "🏠 Início"
tab_name_gerador = "🎫 Etiquetas"
tab_name_extrator = "🔢 Extrator Seriais"
tab_name_sap = "🔑 Login SAP"
tab_name_bandeirada = "🚩 BANDEIRADA"
tab_name_osme = "📄 OSME"
tab_name_transfosme = "📦 Transf. OSME"
tab_name_mudanca = "🔄 Mudança Recurso"
tab_name_zm02 = "⚙️ Entrada ZM02"
tab_name_conferencia = "📦 Conferencia"
tab_name_fichas = "📋 Fichas"
tab_name_plataforma = "📈 Relatório Plataforma"
tab_name_consultar_migo = "🔍 Consultar MIGO"
tab_name_saida_osme = "📤 Saída OSME"
tab_name_contagem_carro = "🚗 Contagem de Carro"
tab_name_atualizar_osme = "🔄 Atualizar OSME"
tab_name_contagem_deposito = "🏢 Contagem Depósito"

CATEGORIAS_SIDEBAR = [
    {
        "nome": "PRINCIPAL",
        "itens": [
            ("🏠", "Início", tab_name_home),
        ]
    },
    {
        "nome": "OPERAÇÕES SAP",
        "itens": [
            ("🔑", "Login SAP", tab_name_sap),
            ("🚩", "BANDEIRADA", tab_name_bandeirada),
            ("📄", "OSME", tab_name_osme),
            ("🔄", "Atualizar OSME", tab_name_atualizar_osme),
            ("📤", "Saída OSME", tab_name_saida_osme),
            ("📦", "Transf. OSME", tab_name_transfosme),
            ("🔍", "Consultar MIGO", tab_name_consultar_migo),
            ("⚙️", "Entrada ZM02", tab_name_zm02),
            ("🔄", "Mudança Recurso", tab_name_mudanca),
        ]
    },
    {
        "nome": "ALMOXARIFADO & EQUIPES",
        "itens": [
            ("🏢", "Contagem Depósito", tab_name_contagem_deposito),
            ("🚗", "Contagem de Carro", tab_name_contagem_carro),
            ("📦", "Conferência", tab_name_conferencia),
            ("📋", "Fichas WorkForce", tab_name_fichas),
        ]
    },
    {
        "nome": "FERRAMENTAS & ETIQUETAS",
        "itens": [
            ("🎫", "Etiquetas", tab_name_gerador),
            ("🔢", "Extrator Seriais", tab_name_extrator),
            ("📈", "Relatório Plataforma", tab_name_plataforma),
        ]
    },
]

# 1. SIDEBAR LATERAL (Empacotada primeiro no lado esquerdo do frame_corpo)
sidebar_navegacao = SidebarNavegacao(
    frame_corpo,
    notebook=None,
    categorias=CATEGORIAS_SIDEBAR,
    paletas=PALETAS,
    obter_tema_atual=obter_tema_atual
)

# 2. CONTAINER DO CONTEÚDO (Empacotado à esquerda, expandindo e ocupando todo o espaço restante)
frame_conteudo = ctk.CTkFrame(frame_corpo, corner_radius=0, fg_color="transparent")
frame_conteudo.pack(side="left", fill=tk.BOTH, expand=True, padx=(4, 10), pady=(0, 4))

# 3. NOTEBOOK (TABS) inserido no frame_conteudo, com segmented_button 100% suprimido
notebook = ctk.CTkTabview(frame_conteudo, corner_radius=8)
notebook._set_grid_segmented_button = lambda: None
notebook.grid_rowconfigure(0, weight=0, minsize=0)
notebook.grid_rowconfigure(1, weight=0, minsize=0)
notebook.grid_rowconfigure(2, weight=0, minsize=0)
notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

tab_home = notebook.add(tab_name_home)
tab_gerador = notebook.add(tab_name_gerador)
tab_extrator = notebook.add(tab_name_extrator)
tab_sap = notebook.add(tab_name_sap)
tab_bandeirada = notebook.add(tab_name_bandeirada)
tab_osme = notebook.add(tab_name_osme)
tab_transfosme = notebook.add(tab_name_transfosme)
tab_mudanca = notebook.add(tab_name_mudanca)
tab_conferencia = notebook.add(tab_name_conferencia)
tab_zm02 = notebook.add(tab_name_zm02)
tab_fichas = notebook.add(tab_name_fichas)
tab_plataforma = notebook.add(tab_name_plataforma)
tab_consultar_migo = notebook.add(tab_name_consultar_migo)
tab_saida_osme = notebook.add(tab_name_saida_osme)
tab_contagem_carro = notebook.add(tab_name_contagem_carro)
tab_atualizar_osme = notebook.add(tab_name_atualizar_osme)
tab_contagem_deposito = notebook.add(tab_name_contagem_deposito)

notebook._segmented_button.grid_forget()

# Vincula o notebook preenchido à sidebar e seleciona a tela inicial
sidebar_navegacao.vincular_notebook(notebook)
sidebar_navegacao.selecionar_aba(tab_name_home, trocar_notebook=True)

# ================== INICIALIZAÇAO ASSÍNCRONA E NÃO-BLOQUEANTE DOS MÓDULOS ==================
def inicializar_modulos():
    global dashboard_view, etiquetas_view, extrator_view, sap_view, bandeirada_view, osme_view, transferenciaosme_view, conferencia_view, zm02_view, fichas_view, consultar_migo_view, saida_osme_view, contagem_carro_view, atualizar_osme_view, contagem_deposito_view

    # Importações pesadas só começam depois que o splash já foi desenhado.
    global construir_aba_sap, obter_info_sap_atual, consultar_info_sap_async
    global construir_aba_bandeirada, construir_aba_conferencia
    global construir_aba_etiquetas, carregar_caminho_planilha_lotes
    global construir_aba_extrair_osme, construir_aba_extrator, JanelaChat
    global construir_aba_zmcse003, construir_aba_transferencia_osme
    global construir_aba_mudanca_recurso, construir_aba_dashboard, construir_aba_fichas, construir_aba_plataforma
    global construir_aba_consultar_migo, construir_aba_saida_osme
    global construir_aba_contagem_carro, construir_aba_atualizar_osme, construir_aba_contagem_deposito
    from modulo_Sap import construir_aba_sap, obter_info_sap_atual, consultar_info_sap_async
    from modulo_bandeirada import construir_aba_bandeirada
    from modulo_conferencia import construir_aba_conferencia
    from modulo_etiquetas import construir_aba_etiquetas, carregar_caminho_planilha_lotes
    from modulo_extrair_OSME import construir_aba_extrair_osme
    from modulo_extrator import construir_aba_extrator
    from modulo_chat import JanelaChat
    from Modulo_ZMCSE003 import construir_aba_zmcse003
    from modulo_transferenciaosme import construir_aba_transferencia_osme
    from modulo_mudanca_recurso import construir_aba_mudanca_recurso
    from modulo_dashboard import construir_aba_dashboard
    from modulo_fichas import construir_aba_fichas
    from modulo_plataforma import construir_aba_plataforma
    from modulo_consultar_migo import construir_aba_consultar_migo
    from modulo_saida_osme import construir_aba_saida_osme
    from modulo_contagem_carro import construir_aba_contagem_carro
    from modulo_atualizar_osme import construir_aba_atualizar_osme
    from modulo_contagem_deposito import construir_aba_contagem_deposito

    def etapa_1():
        global dashboard_view
        splash_progress.set(0.15)
        lbl_splash_loading.configure(text="Carregando Painel Principal...")
        try:
            dashboard_view = construir_aba_dashboard(tab_home, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Dashboard: {e}")
        root.after(10, etapa_2)

    def etapa_2():
        global etiquetas_view
        splash_progress.set(0.30)
        lbl_splash_loading.configure(text="Carregando Gerador de Etiquetas...")
        try:
            etiquetas_view = construir_aba_etiquetas(tab_gerador, root, ARQUIVO_BANCO, PALETAS, obter_tema_atual, lbl_status)
            etiquetas_view.carregar_banco_dados_async()
        except Exception as e:
            log_error(f"Erro ao carregar Etiquetas: {e}")
        root.after(10, etapa_3)

    def etapa_3():
        global extrator_view
        splash_progress.set(0.48)
        lbl_splash_loading.configure(text="Carregando Extrator de Seriais...")
        try:
            extrator_view = construir_aba_extrator(tab_extrator, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Extrator: {e}")
        root.after(10, etapa_4)

    def etapa_4():
        global sap_view
        splash_progress.set(0.65)
        lbl_splash_loading.configure(text="Carregando Módulo SAP...")
        try:
            sap_view = construir_aba_sap(tab_sap, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar SAP: {e}")
        root.after(10, etapa_5)

    def etapa_5():
        global bandeirada_view, osme_view
        splash_progress.set(0.80)
        lbl_splash_loading.configure(text="Carregando Bandeirada e OSME...")
        try:
            bandeirada_view = construir_aba_bandeirada(tab_bandeirada, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Bandeirada: {e}")

        try:
            osme_view = construir_aba_extrair_osme(tab_osme, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar OSME: {e}")

        root.after(10, etapa_6)

    def etapa_6():
        global transferenciaosme_view, conferencia_view, zm02_view, mudanca_recurso_view, consultar_migo_view, saida_osme_view, contagem_carro_view, atualizar_osme_view
        splash_progress.set(0.95)
        lbl_splash_loading.configure(text="Carregando Módulos Finais e Atualização OSME...")

        try:
            transferenciaosme_view = construir_aba_transferencia_osme(tab_transfosme, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Transf. OSME: {e}")

        try:
            mudanca_recurso_view = construir_aba_mudanca_recurso(tab_mudanca, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Mudança de Recurso: {e}")

        try:
            conferencia_view = construir_aba_conferencia(tab_conferencia, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Conferência: {e}")

        try:
            zm02_view = construir_aba_zmcse003(tab_zm02, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar ZM02: {e}")

        try:
            consultar_migo_view = construir_aba_consultar_migo(tab_consultar_migo, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Consultar MIGO: {e}")

        try:
            saida_osme_view = construir_aba_saida_osme(tab_saida_osme, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Saída OSME: {e}")

        try:
            contagem_carro_view = construir_aba_contagem_carro(tab_contagem_carro, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Contagem de Carro: {e}")

        try:
            atualizar_osme_view = construir_aba_atualizar_osme(tab_atualizar_osme, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Atualizar OSME: {e}")

        try:
            contagem_deposito_view = construir_aba_contagem_deposito(tab_contagem_deposito, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Contagem de Depósito: {e}")

        root.after(10, etapa_final)

    def etapa_final():
        splash_progress.set(1.0)
        lbl_splash_loading.configure(text="Pronto!")

        def _finalizar_splash():
            try:
                root.deiconify()
                root.lift()
                root.focus_force()
            except Exception as e:
                log_error(f"Erro ao exibir janela principal: {e}")
            try:
                splash.destroy()
            except Exception:
                pass
            try:
                aplicar_icone_janela(root)
            except Exception:
                pass
            try:
                aplicar_tema(obter_tema_atual())
            except Exception:
                pass
            try:
                show_toast(root, "N.E.X.U.S pronto!", duration=2000, type="success")
                play_sound("done")
            except Exception:
                pass
            try:
                root.after(400, atualizar_info_sap)
            except Exception:
                pass
            try:
                root.after(2500, lambda: iniciar_checagem_atualizacao(root, PALETAS, obter_tema_atual))
            except Exception:
                pass

        root.after(60, _finalizar_splash)

    # O formulário é local e não depende do SAP; inicializa junto ao restante da UI.
    try:
        fichas_view = construir_aba_fichas(tab_fichas, PALETAS, obter_tema_atual)
    except Exception as e:
        log_error(f"Erro ao carregar Fichas: {e}")
        fichas_view = None

    try:
        plataforma_view = construir_aba_plataforma(tab_plataforma, lbl_status, PALETAS, obter_tema_atual)
    except Exception as e:
        log_error(f"Erro ao carregar Plataforma: {e}")
        plataforma_view = None

    # Dá ao sistema uma oportunidade de pintar o splash antes dos imports pesados.
    root.after(120, etapa_1)


# ================== INTEGRAÇÃO CHATBOT INTELIGENTE ==================
chat_janela_ref = [None]

def _trocar_aba(tab):
    def _set():
        try:
            notebook.set(tab)
            if sidebar_navegacao:
                sidebar_navegacao.selecionar_aba(tab, trocar_notebook=False)
        except Exception:
            pass
    root.after(0, _set)

def _buscar_lote_chat(lote):
    import pandas as pd
    lote = str(lote).strip().upper()
    if not lote:
        return {"erro": "Informe o número do lote."}
    try:
        caminho_planilha = carregar_caminho_planilha_lotes()
        if not os.path.exists(caminho_planilha):
            return {"erro": f"Planilha de lotes não encontrada em: {caminho_planilha}"}

        df = pd.read_excel(caminho_planilha, sheet_name="ESTOQUE", engine="openpyxl")
        df.columns = df.columns.str.strip()
        if "LOTE" not in df.columns:
            return {"erro": "Coluna 'LOTE' não encontrada na planilha de estoque."}

        df["LOTE"] = df["LOTE"].astype(str).str.strip().str.upper()
        filtrado = df[df["LOTE"] == lote]
        if filtrado.empty:
            return {"erro": f"Lote '{lote}' não encontrado na planilha de estoque."}

        seriais = filtrado["SERIE"].dropna().astype(str).str.strip().tolist() if "SERIE" in filtrado.columns else []
        material = str(filtrado.iloc[0]["MATERIAL"]).strip() if "MATERIAL" in filtrado.columns and pd.notna(filtrado.iloc[0]["MATERIAL"]) else ""
        deposito = str(filtrado.iloc[0]["DEPOSITO"]).strip() if "DEPOSITO" in filtrado.columns and pd.notna(filtrado.iloc[0]["DEPOSITO"]) else ""

        def _atualizar_ui():
            try:
                _trocar_aba(tab_name_gerador)
                if etiquetas_view:
                    etiquetas_view.campo_lote.delete(0, tk.END)
                    etiquetas_view.campo_lote.insert(0, lote)
                    etiquetas_view.campo_qtd.delete(0, tk.END)
                    etiquetas_view.campo_qtd.insert(0, str(len(seriais)))
                    etiquetas_view.seriais_lote = seriais
                    if material:
                        etiquetas_view.campo_material.delete(0, tk.END)
                        etiquetas_view.campo_material.insert(0, material)
                        etiquetas_view.buscar_material()
                    if deposito in ["CL03", "CL04", "CS06"]:
                        etiquetas_view.deposito_var.set(deposito)
                    etiquetas_view.atualizar_preview()
            except Exception:
                pass

        root.after(0, _atualizar_ui)

        desc = ""
        if etiquetas_view and not etiquetas_view.banco.empty and material:
            res_banco = etiquetas_view.banco[etiquetas_view.banco["Material"] == material]
            col_desc = etiquetas_view._obter_coluna_descricao() if hasattr(etiquetas_view, "_obter_coluna_descricao") else None
            if not res_banco.empty and col_desc:
                desc = str(res_banco.iloc[0][col_desc])

        return {
            "lote": lote,
            "material": material,
            "descricao": desc,
            "deposito": deposito,
            "quantidade": len(seriais),
            "seriais": seriais[:30],
            "total_seriais": len(seriais)
        }
    except Exception as e:
        return {"erro": str(e)}

def _buscar_multiplos_lotes_chat(lotes):
    if isinstance(lotes, str):
        import re
        lista_lotes = [l.strip().upper() for l in re.split(r"[,;\s\n]+", lotes) if l.strip()]
    elif isinstance(lotes, list):
        lista_lotes = [str(l).strip().upper() for l in lotes if str(l).strip()]
    else:
        return {"erro": "Formato de lotes inválido."}

    resultados = {}
    for l in lista_lotes[:15]:
        resultados[l] = _buscar_lote_chat(l)
    return {"total_lotes": len(resultados), "resultados": resultados}

def _gerar_doc_chat(lote):
    dados_lote = _buscar_lote_chat(lote)
    if "erro" in dados_lote:
        return {"erro": dados_lote["erro"]}

    seriais = dados_lote.get("seriais", [])
    if not seriais:
        return {"erro": f"Nenhum serial encontrado para o lote {lote}."}

    from modulo_etiquetas import carregar_google_script_url
    google_script_url = carregar_google_script_url()
    if not google_script_url:
        return {"erro": "URL do Google Apps Script não configurada nas Etiquetas."}

    payload = {
        "lote": dados_lote.get("lote", lote),
        "material": dados_lote.get("material", ""),
        "descricao": dados_lote.get("descricao", ""),
        "deposito": dados_lote.get("deposito", "CL03"),
        "seriais": seriais,
        "quantidade": len(seriais)
    }

    try:
        import urllib.request
        dados_json = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            google_script_url, data=dados_json,
            headers={"Content-Type": "application/json"}, method="POST"
        )
        with urllib.request.urlopen(req, timeout=35) as resp:
            resp_data = json.loads(resp.read().decode("utf-8"))
            link = resp_data.get("link", "")
            if link:
                def _set_qr():
                    try:
                        _trocar_aba(tab_name_gerador)
                        if etiquetas_view:
                            etiquetas_view.campo_qr_topo.delete(0, tk.END)
                            etiquetas_view.campo_qr_topo.insert(0, link)
                            etiquetas_view.atualizar_preview()
                    except Exception:
                        pass
                root.after(0, _set_qr)
                return {
                    "lote": lote,
                    "link": link,
                    "status": "Documento Google Docs criado com sucesso!",
                    "quantidade": len(seriais)
                }
            return {"erro": "O Google Apps Script não retornou um link válido."}
    except Exception as e:
        return {"erro": f"Falha ao conectar ao Google Apps Script: {str(e)}"}

def _gerar_etiqueta_chat(deposito, material="", descricao="", qtd="", unidade="PCS", lote=""):
    _trocar_aba(tab_name_gerador)
    from gerador import gerar_pdf_placa_tradicional
    try:
        caminho = gerar_pdf_placa_tradicional(deposito, material, descricao, qtd, unidade, lote=lote, tamanho="Grande")
        return {"status": "PDF gerado com sucesso", "arquivo": caminho, "deposito": deposito, "material": material}
    except Exception as e:
        return {"erro": str(e)}

def _consultar_material_chat(codigo):
    cod = str(codigo).strip()
    if not cod:
        return {"erro": "Código do material não informado."}

    import pandas as pd
    try:
        df = pd.DataFrame()
        if etiquetas_view and not etiquetas_view.banco.empty:
            df = etiquetas_view.banco
        elif os.path.exists(ARQUIVO_BANCO):
            df = pd.read_excel(ARQUIVO_BANCO)
            df.columns = df.columns.str.strip()
            if "Material" in df.columns:
                df["Material"] = df["Material"].astype(str).str.strip()

        if df.empty or "Material" not in df.columns:
            return {"erro": "Banco de materiais não carregado."}

        res = df[df["Material"] == cod]
        col_desc = None
        for col in ["Texto breve material", "Descricao", "DESCRICAO", "Texto Breve", "Descrição"]:
            if col in df.columns:
                col_desc = col
                break

        if not res.empty and col_desc:
            desc = str(res.iloc[0][col_desc]).strip()
            def _atualizar_ui():
                try:
                    _trocar_aba(tab_name_gerador)
                    if etiquetas_view:
                        etiquetas_view.campo_material.delete(0, tk.END)
                        etiquetas_view.campo_material.insert(0, cod)
                        etiquetas_view.campo_descricao.delete(0, tk.END)
                        etiquetas_view.campo_descricao.insert(0, desc)
                        etiquetas_view.atualizar_preview()
                except Exception:
                    pass
            root.after(0, _atualizar_ui)
            return {"codigo": cod, "descricao": desc}
        return {"codigo": cod, "descricao": "Material não encontrado no banco de dados."}
    except Exception as e:
        return {"erro": str(e)}

def _consultar_multiplos_materiais_chat(codigos):
    if isinstance(codigos, str):
        import re
        lista = [c.strip() for c in re.split(r"[,;\s\n]+", codigos) if c.strip()]
    elif isinstance(codigos, list):
        lista = [str(c).strip() for c in codigos if str(c).strip()]
    else:
        return {"erro": "Formato de códigos inválido."}

    resultados = {}
    for c in lista[:20]:
        resultados[c] = _consultar_material_chat(c)
    return {"total_consultados": len(resultados), "materiais": resultados}

def _listar_lotes_chat():
    import pandas as pd
    try:
        caminho = carregar_caminho_planilha_lotes()
        df = pd.read_excel(caminho, sheet_name="ESTOQUE", engine="openpyxl")
        df.columns = df.columns.str.strip()
        if "LOTE" in df.columns:
            lotes = df["LOTE"].dropna().unique().tolist()
            return {"total": len(lotes), "lotes": sorted([str(l).strip() for l in lotes])}
        return {"erro": "Coluna LOTE não encontrada na planilha de estoque."}
    except Exception as e:
        return {"erro": str(e)}

def _extrair_seriais_chat(texto):
    import re
    if not texto:
        return {"erro": "Nenhum texto informado para extração."}

    padrao_intervalo = re.compile(r"(\d+)\s*(?:a|à|-|ate|até)\s*(\d+)", re.IGNORECASE)
    linhas = texto.splitlines()
    seriais_extraidos = []

    for linha in linhas:
        match = padrao_intervalo.search(linha)
        if match:
            try:
                ini, fim = int(match.group(1)), int(match.group(2))
                if ini > fim:
                    ini, fim = fim, ini
                tamanho_padrao = max(len(match.group(1)), len(match.group(2)))
                if (fim - ini) <= 2000:
                    for num in range(ini, fim + 1):
                        seriais_extraidos.append(str(num).zfill(tamanho_padrao))
                continue
            except Exception:
                pass
        achados = re.findall(r"\b\d{6,18}\b", linha)
        for s in achados:
            seriais_extraidos.append(s)

    vistos = set()
    unicos = []
    for s in seriais_extraidos:
        if s not in vistos:
            vistos.add(s)
            unicos.append(s)

    def _atualizar_ui():
        try:
            _trocar_aba(tab_name_extrator)
            if extrator_view:
                extrator_view.txt_entrada.delete("1.0", tk.END)
                extrator_view.txt_entrada.insert("1.0", texto)
                extrator_view.txt_sem_ponto_virgula.delete("1.0", tk.END)
                extrator_view.txt_sem_ponto_virgula.insert("1.0", "\n".join(unicos))
                extrator_view.txt_com_ponto_virgula.delete("1.0", tk.END)
                extrator_view.txt_com_ponto_virgula.insert("1.0", ";\n".join(unicos) + (";" if unicos else ""))
                extrator_view.lbl_sem.configure(text=f"Total: {len(unicos)} seriais")
                extrator_view.lbl_com.configure(text=f"Total: {len(unicos)} seriais")
        except Exception:
            pass

    root.after(0, _atualizar_ui)
    return {"total": len(unicos), "seriais": unicos[:50], "total_extraidos": len(unicos)}

def _login_sap_chat(usuario):
    _trocar_aba(tab_name_sap)
    try:
        from modulo_Sap import logar_sap, logins_salvos, carregar_logins
        carregar_logins()
        nome_lower = usuario.lower().strip()
        encontrado = next((dados for nome, dados in logins_salvos.items() if nome.lower() == nome_lower or nome_lower in nome.lower()), None)
        if not encontrado:
            return {"erro": f"Usuário '{usuario}' não encontrado no cadastro."}
        import threading
        threading.Thread(target=logar_sap, args=(encontrado["usuario"], encontrado["senha"], usuario, lbl_status), daemon=True).start()
        return {"status": f"Iniciando login SAP para {usuario} em segundo plano..."}
    except Exception as e:
        return {"erro": str(e)}

def _info_sap_chat():
    return {"info": obter_info_sap_atual()}

def _copiar_seriais_chat():
    def _copiar():
        try:
            if extrator_view:
                _trocar_aba(tab_name_extrator)
                extrator_view.copiar_sem_ponto_virgula()
        except Exception:
            pass
    root.after(0, _copiar)
    return {"status": "Seriais copiados para a área de transferência com sucesso!"}

def _executar_bandeirada_chat(recursos):
    if isinstance(recursos, str):
        import re
        lista = [r.strip().upper() for r in re.split(r"[,;\s\n]+", recursos) if r.strip()]
    elif isinstance(recursos, list):
        lista = [str(r).strip().upper() for r in recursos if str(r).strip()]
    else:
        return {"erro": "Formato de recursos/equipes inválido."}

    if not lista:
        return {"erro": "Informe ao menos um recurso/equipe para bandeirada."}

    from modulo_Sap import verificar_sap_aberto
    if not verificar_sap_aberto():
        return {"erro": "O SAP precisa estar aberto para executar a Bandeirada. Abra o SAP e tente novamente."}

    def _iniciar():
        try:
            _trocar_aba(tab_name_bandeirada)
            if bandeirada_view:
                bandeirada_view.txt_recursos.delete("1.0", tk.END)
                bandeirada_view.txt_recursos.insert("1.0", "\n".join(lista))
                bandeirada_view.iniciar_bandeirada()
        except Exception:
            pass

    root.after(0, _iniciar)
    return {
        "status": f"Processo de Bandeirada iniciado no SAP para: {', '.join(lista)}",
        "equipes": lista
    }

def _extrair_osme_chat(recursos):
    if isinstance(recursos, str):
        import re
        lista = [r.strip().upper() for r in re.split(r"[,;\s\n]+", recursos) if r.strip()]
    elif isinstance(recursos, list):
        lista = [str(r).strip().upper() for r in recursos if str(r).strip()]
    else:
        return {"erro": "Formato de recursos/equipes inválido."}

    if not lista:
        return {"erro": "Informe ao menos um recurso/equipe para extração de OSME."}

    from modulo_Sap import verificar_sap_aberto
    if not verificar_sap_aberto():
        return {"erro": "O SAP precisa estar aberto para extrair OSME. Abra o SAP e tente novamente."}

    def _iniciar():
        try:
            _trocar_aba(tab_name_osme)
            if osme_view:
                osme_view.txt_recursos.delete("1.0", tk.END)
                osme_view.txt_recursos.insert("1.0", "\n".join(lista))
                osme_view.iniciar_extracao()
        except Exception:
            pass

    root.after(0, _iniciar)
    return {
        "status": f"Extração de OSME iniciada no SAP para: {', '.join(lista)}",
        "equipes": lista
    }

def _consultar_migo_chat(equipe, ordem=""):
    eq = str(equipe).strip().upper()
    if not eq:
        return {"erro": "Informe o código da equipe (ex: RCL001, RCL002)."}

    from modulo_Sap import verificar_sap_aberto
    if not verificar_sap_aberto():
        return {"erro": "O SAP precisa estar aberto para consultar a MIGO. Abra o SAP e tente novamente."}

    def _consultar():
        try:
            _trocar_aba(tab_name_consultar_migo)
            if consultar_migo_view:
                consultar_migo_view._selecionar_equipe(eq)
                if ordem:
                    consultar_migo_view.entry_ordem.delete(0, tk.END)
                    consultar_migo_view.entry_ordem.insert(0, str(ordem).strip())
                consultar_migo_view.iniciar_fluxo_completo()
        except Exception:
            pass

    root.after(0, _consultar)
    detalhe_ordem = f" (Ordem: {ordem})" if ordem else ""
    return {
        "status": f"Consulta IW39 ➡️ MIGO disparada no SAP para a equipe {eq}{detalhe_ordem}",
        "equipe": eq
    }

def _transferir_osme_chat(equipes, data_lancamento=None):
    if isinstance(equipes, str):
        import re
        lista = [r.strip().upper() for r in re.split(r"[,;\s\n]+", equipes) if r.strip()]
    elif isinstance(equipes, list):
        lista = [str(r).strip().upper() for r in equipes if str(r).strip()]
    else:
        return {"erro": "Formato de equipes inválido."}

    if not lista:
        return {"erro": "Informe as equipes para transferência de OSME."}

    def _executar():
        try:
            _trocar_aba(tab_name_transfosme)
            if transferenciaosme_view:
                if data_lancamento and hasattr(transferenciaosme_view, "entry_data_lancamento"):
                    transferenciaosme_view.entry_data_lancamento.delete(0, tk.END)
                    transferenciaosme_view.entry_data_lancamento.insert(0, str(data_lancamento).strip())
                transferenciaosme_view.txt_equipes.delete("1.0", tk.END)
                transferenciaosme_view.txt_equipes.insert("1.0", "\n".join(lista))
                transferenciaosme_view.iniciar_transferencia()
        except Exception:
            pass

    root.after(0, _executar)
    return {
        "status": f"Transferência de OSME disparada para {len(lista)} equipe(s): {', '.join(lista)}.",
        "equipes": lista
    }

def _mudar_recurso_chat(origem, destino="", acao="consultar"):
    origem = str(origem).strip().upper()
    destino = str(destino).strip().upper() if destino else ""
    if not origem:
        return {"erro": "Informe ao menos a equipe de origem."}

    def _aplicar():
        try:
            _trocar_aba(tab_name_mudanca)
            if mudanca_recurso_view:
                mudanca_recurso_view.ent_origem.delete(0, tk.END)
                mudanca_recurso_view.ent_origem.insert(0, origem)
                if destino:
                    mudanca_recurso_view.ent_destino.delete(0, tk.END)
                    mudanca_recurso_view.ent_destino.insert(0, destino)
                if acao == "consultar":
                    mudanca_recurso_view.acao_consultar_origem()
                elif acao == "executar" and destino:
                    mudanca_recurso_view.acao_executar_mudanca()
        except Exception:
            pass

    root.after(0, _aplicar)
    msg_acao = f"Consulta de saldo da equipe {origem} iniciada" if acao == "consultar" else f"Mudança de recurso {origem} ➡️ {destino} preparada"
    return {"status": msg_acao, "origem": origem, "destino": destino}

def _abrir_conferencia_chat(base="", fisico=""):
    def _abrir():
        try:
            _trocar_aba(tab_name_conferencia)
            if conferencia_view:
                if base:
                    conferencia_view.txt_base.delete("1.0", tk.END)
                    conferencia_view.txt_base.insert("1.0", str(base))
                if fisico:
                    conferencia_view.txt_fisico.delete("1.0", tk.END)
                    conferencia_view.txt_fisico.insert("1.0", str(fisico))
                if base and fisico:
                    conferencia_view.comparar()
        except Exception:
            pass
    root.after(0, _abrir)
    return {"status": "Módulo de Conferência aberto e pronto para verificação."}

def _abrir_fichas_chat(re="", recurso=""):
    def _abrir():
        try:
            _trocar_aba(tab_name_fichas)
            if fichas_view:
                if re and hasattr(fichas_view, "campo_re_eletricista_1"):
                    fichas_view.campo_re_eletricista_1.delete(0, tk.END)
                    fichas_view.campo_re_eletricista_1.insert(0, str(re).strip())
                if recurso and hasattr(fichas_view, "campo_recurso"):
                    fichas_view.campo_recurso.delete(0, tk.END)
                    fichas_view.campo_recurso.insert(0, str(recurso).strip().upper())
        except Exception:
            pass
    root.after(0, _abrir)
    detalhes = []
    if re: detalhes.append(f"RE: {re}")
    if recurso: detalhes.append(f"Recurso: {recurso}")
    det_txt = f" ({', '.join(detalhes)})" if detalhes else ""
    return {"status": f"Módulo Fichas aberto{det_txt}."}

def _abrir_modulo_chat(modulo):
    m = str(modulo).lower().strip()
    mapa_abas = {
        "bandeirada": tab_name_bandeirada,
        "bandeira": tab_name_bandeirada,
        "osme": tab_name_osme,
        "extrair osme": tab_name_osme,
        "migo": tab_name_consultar_migo,
        "consultar migo": tab_name_consultar_migo,
        "clandestina": tab_name_consultar_migo,
        "clandestinas": tab_name_consultar_migo,
        "etiqueta": tab_name_gerador,
        "etiquetas": tab_name_gerador,
        "gerador": tab_name_gerador,
        "extrator": tab_name_extrator,
        "seriais": tab_name_extrator,
        "sap": tab_name_sap,
        "login sap": tab_name_sap,
        "login": tab_name_sap,
        "transfosme": tab_name_transfosme,
        "transf osme": tab_name_transfosme,
        "transferencia osme": tab_name_transfosme,
        "transferência osme": tab_name_transfosme,
        "mudanca": tab_name_mudanca,
        "mudança": tab_name_mudanca,
        "mudanca recurso": tab_name_mudanca,
        "mudança recurso": tab_name_mudanca,
        "conferencia": tab_name_conferencia,
        "conferência": tab_name_conferencia,
        "zm02": tab_name_zm02,
        "entrada zm02": tab_name_zm02,
        "zmcse003": tab_name_zm02,
        "ficha": tab_name_fichas,
        "fichas": tab_name_fichas,
        "plataforma": tab_name_plataforma,
        "relatorio plataforma": tab_name_plataforma,
        "saida osme": tab_name_saida_osme,
        "saída osme": tab_name_saida_osme,
        "saida para osme": tab_name_saida_osme,
        "saída para osme": tab_name_saida_osme,
        "contagem carro": tab_name_contagem_carro,
        "contagem de carro": tab_name_contagem_carro,
        "contagem veiculo": tab_name_contagem_carro,
        "contagem de veiculo": tab_name_contagem_carro,
        "balanco carro": tab_name_contagem_carro,
        "balanço carro": tab_name_contagem_carro,
        "carro": tab_name_contagem_carro,
        "atualizar osme": tab_name_atualizar_osme,
        "atualizacao osme": tab_name_atualizar_osme,
        "atualização osme": tab_name_atualizar_osme,
        "atualizar vencimento": tab_name_atualizar_osme,
        "contagem deposito": tab_name_contagem_deposito,
        "contagem de deposito": tab_name_contagem_deposito,
        "contagem depósito": tab_name_contagem_deposito,
        "contagem de depósito": tab_name_contagem_deposito,
        "rotativo": tab_name_contagem_deposito,
        "lote 0": tab_name_contagem_deposito,
        "home": tab_name_home,
        "inicio": tab_name_home,
        "início": tab_name_home,
        "dashboard": tab_name_home,
    }
    aba_alvo = None
    for k, v in mapa_abas.items():
        if k in m or m in k:
            aba_alvo = v
            break
    if not aba_alvo:
        aba_alvo = tab_name_home

    _trocar_aba(aba_alvo)
    return {"status": f"Aba alterada para: {aba_alvo}"}

def _consultar_iq09_chat(seriais=None):
    from modulo_extrator import extrair_lista_seriais, consultar_medidor_iq09_sap, JanelaDetalheMedidor, JanelaMultiplosMedidores
    from modulo_Sap import verificar_sap_aberto

    seriais_extraidos = []
    if seriais:
        seriais_extraidos = extrair_lista_seriais(seriais)

    # Se não passou parâmetro direto mas há seriais já extraídos na aba Extrator, utiliza-os
    if not seriais_extraidos and extrator_view:
        try:
            lista_aba = extrator_view._obter_lista_sem_ponto_virgula()
            if lista_aba:
                seriais_extraidos = lista_aba
        except Exception:
            pass

    if not seriais_extraidos:
        return {"erro": "Informe o número do medidor, serial ou intervalo (ex: '18400025 a 18400028') para pesquisar na IQ09."}

    if not verificar_sap_aberto():
        return {"erro": "O SAP precisa estar aberto para pesquisar na IQ09. Abra o SAP e tente novamente."}

    # Executa a consulta no SAP diretamente (já estamos na thread de background do chatbot)
    resultado = consultar_medidor_iq09_sap(seriais_extraidos)
    if "erro" in resultado:
        return resultado

    # Abre o Pop-up na interface principal (thread do Tkinter)
    lista = resultado.get("lista_medidores", [])
    dados = resultado.get("dados_medidor")

    def _abrir_popup():
        try:
            if len(lista) > 1:
                JanelaMultiplosMedidores(root, lista, PALETAS, obter_tema_atual)
            elif dados:
                JanelaDetalheMedidor(root, dados, PALETAS, obter_tema_atual)
        except Exception as e:
            print(f"[CHAT] Erro ao abrir pop-up da IQ09: {e}")

    root.after(0, _abrir_popup)
    return resultado

def _saida_osme_chat(equipe=None, seriais=None, material=None, quantidade=None, material_ns=None, quantidade_ns=None):
    if not saida_osme_view:
        return {"erro": "Módulo Saída OSME ainda não foi inicializado."}

    mat_final = material_ns if material_ns is not None else material
    qtd_final = quantidade_ns if quantidade_ns is not None else quantidade

    resultado = {}
    ev = threading.Event()

    def _exec():
        try:
            _trocar_aba(tab_name_saida_osme)
            res = saida_osme_view.preparar_saida_via_chat(
                equipe=equipe,
                seriais=seriais,
                material_ns=mat_final,
                quantidade_ns=qtd_final
            )
            resultado.update(res)
        except Exception as e:
            resultado["erro"] = str(e)
        finally:
            ev.set()

    root.after(0, _exec)
    ev.wait(timeout=5.0)

    if "erro" in resultado:
        return resultado

    mats_resumo = resultado.get("resumo_materiais", [])
    txt_mats = "\n".join([f"  ▸ {m}" for m in mats_resumo]) if mats_resumo else "  (Nenhum material adicionado)"
    alerta_bloqueio = "\n⚠️ **ATENÇÃO: Existem itens BLOQUEADOS (já constam como fornecidos)!**" if resultado.get("tem_bloqueados") else ""

    msg = (
        f"📋 **Dados preparados na aba Saída OSME:**\n\n"
        f"• **Equipe**: `{resultado.get('equipe', '-')}`\n"
        f"• **OSME**: `{resultado.get('osme', '-')}`\n"
        f"• **Depósito**: `{resultado.get('deposito', '-')}`\n\n"
        f"**Materiais a processar:**\n{txt_mats}\n"
        f"{alerta_bloqueio}\n"
        f"❓ **Deseja processar a saída no SAP agora?**\n"
        f"👉 Responda **'confirmo'** ou **'sim'** aqui mesmo no chat para executar diretamente, ou acesse a aba para conferir."
    )
    resultado["mensagem"] = msg
    resultado["status"] = "Dados preparados na aba Saída OSME."
    return resultado

def _executar_saida_osme_chat():
    if not saida_osme_view:
        return {"erro": "Módulo Saída OSME ainda não foi inicializado."}
    resultado = {}
    ev = threading.Event()
    def _exec():
        try:
            _trocar_aba(tab_name_saida_osme)
            res = saida_osme_view.executar_saida_via_chat()
            resultado.update(res)
        except Exception as e:
            resultado["erro"] = str(e)
        finally:
            ev.set()
    root.after(0, _exec)
    ev.wait(timeout=5.0)
    return resultado

def _contagem_carro_chat(equipe=None, seriais=None, acao="abrir"):
    def _exec():
        try:
            _trocar_aba(tab_name_contagem_carro)
            if contagem_carro_view:
                if equipe:
                    contagem_carro_view.entry_equipe.delete(0, tk.END)
                    contagem_carro_view.entry_equipe.insert(0, str(equipe).strip().upper())
                if seriais:
                    from modulo_extrator import extrair_lista_seriais
                    lista = extrair_lista_seriais(str(seriais))
                    if lista:
                        contagem_carro_view.txt_seriais.delete("1.0", tk.END)
                        contagem_carro_view.txt_seriais.insert("1.0", "\n".join(lista))
                        contagem_carro_view._ao_alterar_seriais()
                if acao == "executar" and equipe and seriais:
                    contagem_carro_view.executar_ciclo_completo_acao()
                elif acao == "osme" and equipe:
                    contagem_carro_view.puxar_medidores_osme_acao()
                elif acao == "auditar" and seriais:
                    contagem_carro_view.consultar_zmcse01_acao()
        except Exception:
            pass
    root.after(0, _exec)
    eq_txt = f" para a equipe {equipe}" if equipe else ""
    return {"status": f"Módulo Contagem de Carro aberto{eq_txt}."}

funcoes_chat = {
    "buscar_lote": _buscar_lote_chat,
    "buscar_multiplos_lotes": _buscar_multiplos_lotes_chat,
    "gerar_doc_google": _gerar_doc_chat,
    "gerar_etiqueta": _gerar_etiqueta_chat,
    "consultar_material": _consultar_material_chat,
    "consultar_multiplos_materiais": _consultar_multiplos_materiais_chat,
    "listar_lotes": _listar_lotes_chat,
    "extrair_seriais": _extrair_seriais_chat,
    "login_sap": _login_sap_chat,
    "info_sap": _info_sap_chat,
    "copiar_seriais": _copiar_seriais_chat,
    "executar_bandeirada": _executar_bandeirada_chat,
    "extrair_osme": _extrair_osme_chat,
    "consultar_migo": _consultar_migo_chat,
    "transferir_osme": _transferir_osme_chat,
    "mudar_recurso": _mudar_recurso_chat,
    "abrir_conferencia": _abrir_conferencia_chat,
    "abrir_fichas": _abrir_fichas_chat,
    "abrir_modulo": _abrir_modulo_chat,
    "consultar_iq09": _consultar_iq09_chat,
    "saida_osme": _saida_osme_chat,
    "executar_saida_osme": _executar_saida_osme_chat,
    "contagem_carro": _contagem_carro_chat,
}

def abrir_chat():
    if chat_janela_ref[0] and chat_janela_ref[0].winfo_exists():
        chat_janela_ref[0].lift()
        chat_janela_ref[0].focus_force()
        return
    chat_popup = JanelaChat(root, PALETAS, obter_tema_atual, funcoes_chat)
    chat_janela_ref[0] = chat_popup.janela

btn_chat = ctk.CTkButton(
    root, text="💬", font=("Segoe UI", 16, "bold"),
    fg_color="#00d2ff", hover_color="#0284c7", text_color="#090d16",
    width=38, height=38, corner_radius=19,
    command=abrir_chat, cursor="hand2"
)
btn_chat.place(relx=1.0, rely=0.82, x=-16, anchor="e")

# Suporte suave para arrastar verticalmente o balão pela borda direita caso desejado
def _on_chat_drag_start(event):
    btn_chat._drag_y = event.y_root

def _on_chat_drag(event):
    dy = event.y_root - getattr(btn_chat, "_drag_y", event.y_root)
    btn_chat._drag_y = event.y_root
    cur_y = btn_chat.winfo_y() + dy
    min_y = 70
    max_y = root.winfo_height() - 70
    cur_y = max(min_y, min(max_y, cur_y))
    btn_chat.place(relx=1.0, rely=0.0, y=cur_y, x=-16, anchor="ne")

btn_chat.bind("<Button-1>", _on_chat_drag_start, add="+")
btn_chat.bind("<B1-Motion>", _on_chat_drag, add="+")

aplicar_tema(tema_inicial)

# Gerenciador da Bandeja do Sistema e Fechamento Seguro (segundo plano para fallback de pendências)
tray_manager = None
try:
    from utils_tray import NexusTrayManager
    tray_manager = NexusTrayManager(root)
except Exception as _e_tray:
    log_error(f"Aviso ao inicializar NexusTrayManager: {_e_tray}")

if __name__ == "__main__":
    # Inicia módulos em estágios assíncronos não-bloqueantes
    root.after(20, inicializar_modulos)
    root.mainloop()
