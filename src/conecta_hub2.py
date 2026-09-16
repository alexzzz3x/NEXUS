import os
import sys

from utils_os import garantir_instancia_unica, obter_caminho_recurso, obter_diretorio_executavel

# Garante que apenas uma instância do Conecta Hub rode por vez quando executado como script principal
if __name__ == "__main__":
    _ok_instancia, _mutex_app = garantir_instancia_unica("Nexus_SAP_Hub_SingleInstance_Mutex")
    if not _ok_instancia:
        sys.exit(0)

import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

from utils_ui import show_toast, play_sound, run_async
from utils_logger import log_error, log_info

def global_exception_handler(exc_type, exc_value, exc_traceback):
    log_error(f"Uncaught exception: {exc_value}", exc_info=(exc_type, exc_value, exc_traceback))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

sys.excepthook = global_exception_handler

ctk.set_default_color_theme("blue")  # Tema padrao do CTK

CONFIG_FILE = os.path.join(obter_diretorio_executavel(), "data", "config_tema_hub.txt")

def obter_arquivo_banco():
    caminho_exe = os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx")
    if os.path.exists(caminho_exe):
        return caminho_exe
    return obter_caminho_recurso("data", "banco_de_dados.xlsx")

ARQUIVO_BANCO = obter_arquivo_banco()

PALETAS = {
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60", "scrollbar": "#5f7fa8", "scrollbar_hover": "#7fa2cf", "scrollbar_trough": "#10233d"},
    "Dark Mode": {"bg": "#1a202c", "container": "#2d3748", "texto": "#ffffff", "accent": "#319795", "hover": "#2c7a7b", "input": "#3a4454", "scrollbar": "#64748b", "scrollbar_hover": "#94a3b8", "scrollbar_trough": "#202938"},
    "Light Clean": {"bg": "#ffffff", "container": "#f8f9fa", "texto": "#2d3748", "accent": "#319795", "hover": "#2c7a7b", "input": "#edf2f7", "scrollbar": "#94a3b8", "scrollbar_hover": "#64748b", "scrollbar_trough": "#e2e8f0"},
    "Corporate Slate": {"bg": "#111827", "container": "#1f2937", "texto": "#f9fafb", "accent": "#2563eb", "hover": "#1d4ed8", "input": "#374151", "scrollbar": "#6b7280", "scrollbar_hover": "#9ca3af", "scrollbar_trough": "#172033"},
    "Office Light": {"bg": "#f3f6fb", "container": "#ffffff", "texto": "#1f2937", "accent": "#2563eb", "hover": "#1d4ed8", "input": "#e8eef7", "scrollbar": "#a7b4c8", "scrollbar_hover": "#6b7f99", "scrollbar_trough": "#dce5f2"},
    "Emerald Pro": {"bg": "#0f1f1b", "container": "#18352e", "texto": "#f8fafc", "accent": "#10b981", "hover": "#059669", "input": "#21483e", "scrollbar": "#6ba99a", "scrollbar_hover": "#8fd4c3", "scrollbar_trough": "#132b25"},
    "Graphite Orange": {"bg": "#171717", "container": "#262626", "texto": "#fafafa", "accent": "#f97316", "hover": "#ea580c", "input": "#3f3f46", "scrollbar": "#737373", "scrollbar_hover": "#a3a3a3", "scrollbar_trough": "#202020"},
    "Royal Indigo": {"bg": "#17152b", "container": "#262447", "texto": "#ffffff", "accent": "#8b5cf6", "hover": "#7c3aed", "input": "#34315f", "scrollbar": "#8179bd", "scrollbar_hover": "#a69cf0", "scrollbar_trough": "#201d3b"},
}

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

def carregar_tema_salvo():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                tema = f.read().strip()
                if tema in PALETAS:
                    return tema
        except Exception:
            pass
    return "Conecta Blue"

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
    c = PALETAS[tema_nome]
    
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
    lbl_status.configure(text_color=c["texto"])
    lbl_sap_info.configure(text_color=c["texto"])
    lbl_creditos.configure(text_color=c["texto"])
    try:
        lbl_sub_header.configure(text_color=c.get("scrollbar_hover", "#7fa2cf"))
    except Exception:
        pass

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
root.title("N.E.X.U.S - Núcleo de estoque e extração unificada com o SAP")
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

lbl_splash_sub = ctk.CTkLabel(splash, text="Núcleo de estoque e extração unificada com o SAP", font=("Segoe UI", 11), text_color="#90cdf4")
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

combo_temas = ctk.CTkComboBox(frame_topo, values=list(PALETAS.keys()), command=aplicar_tema, width=160, font=("Segoe UI", 12))
combo_temas.set(tema_inicial)
combo_temas.pack(side=tk.RIGHT, padx=20, pady=15)

# NOTEBOOK (TABS)
notebook = ctk.CTkTabview(root, corner_radius=8)
notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

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

# STATUS BAR
bar_status = ctk.CTkFrame(root, height=30, corner_radius=0)
bar_status.pack(side="bottom", fill="x")

lbl_status = ctk.CTkLabel(bar_status, text=" Pronto | N.E.X.U.S", font=("Segoe UI", 11, "bold"))
lbl_status.pack(side="left", padx=10, pady=2)

lbl_sap_info = ctk.CTkLabel(bar_status, text="SAP: verificando...", font=("Segoe UI", 11, "bold"))
lbl_sap_info.pack(side="left", padx=20, pady=2)

lbl_creditos = ctk.CTkLabel(bar_status, text="2026 Desenvolvido por 🟢 Anderson Vieira™", font=("Segoe UI", 11, "italic"))
lbl_creditos.pack(side="right", padx=10, pady=2)

# ================== INICIALIZAÇAO ASSÍNCRONA E NÃO-BLOQUEANTE DOS MÓDULOS ==================
def inicializar_modulos():
    global dashboard_view, etiquetas_view, extrator_view, sap_view, bandeirada_view, osme_view, transferenciaosme_view, conferencia_view, zm02_view, fichas_view, consultar_migo_view

    # Importações pesadas só começam depois que o splash já foi desenhado.
    global construir_aba_sap, obter_info_sap_atual, consultar_info_sap_async
    global construir_aba_bandeirada, construir_aba_conferencia
    global construir_aba_etiquetas, carregar_caminho_planilha_lotes
    global construir_aba_extrair_osme, construir_aba_extrator, JanelaChat
    global construir_aba_zmcse003, construir_aba_transferencia_osme
    global construir_aba_mudanca_recurso, construir_aba_dashboard, construir_aba_fichas, construir_aba_plataforma
    global construir_aba_consultar_migo
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
            osme_view = construir_aba_extrair_osme(tab_osme, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar Bandeirada/OSME: {e}")
        root.after(10, etapa_6)

    def etapa_6():
        global transferenciaosme_view, conferencia_view, zm02_view, mudanca_recurso_view, consultar_migo_view
        splash_progress.set(0.95)
        lbl_splash_loading.configure(text="Carregando Mudança de Recurso, ZM02, Conferência e Consultar MIGO...")
        try:
            transferenciaosme_view = construir_aba_transferencia_osme(tab_transfosme, lbl_status, PALETAS, obter_tema_atual)
            mudanca_recurso_view = construir_aba_mudanca_recurso(tab_mudanca, lbl_status, PALETAS, obter_tema_atual)
            conferencia_view = construir_aba_conferencia(tab_conferencia, lbl_status, PALETAS, obter_tema_atual)
            zm02_view = construir_aba_zmcse003(tab_zm02, PALETAS, obter_tema_atual)
            consultar_migo_view = construir_aba_consultar_migo(tab_consultar_migo, lbl_status, PALETAS, obter_tema_atual)
        except Exception as e:
            log_error(f"Erro ao carregar módulos finais: {e}")
        root.after(10, etapa_final)

    def etapa_final():
        splash_progress.set(1.0)
        lbl_splash_loading.configure(text="Pronto!")

        def _finalizar_splash():
            try:
                splash.destroy()
            except Exception:
                pass
            root.deiconify()
            aplicar_icone_janela(root)
            root.lift()
            show_toast(root, "N.E.X.U.S pronto!", duration=2000, type="success")
            play_sound("done")
            root.after(400, atualizar_info_sap)

        root.after(60, _finalizar_splash)

    # O formulário é local e não depende do SAP; inicializa junto ao restante da UI.
    fichas_view = construir_aba_fichas(tab_fichas, PALETAS, obter_tema_atual)
    plataforma_view = construir_aba_plataforma(tab_plataforma, lbl_status, PALETAS, obter_tema_atual)

    # Dá ao sistema uma oportunidade de pintar o splash antes dos imports pesados.
    root.after(120, etapa_1)


# ================== INTEGRAÇÃO CHATBOT INTELIGENTE ==================
chat_janela_ref = [None]

def _trocar_aba(tab):
    def _set():
        try:
            notebook.set(tab)
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
}

def abrir_chat():
    if chat_janela_ref[0] and chat_janela_ref[0].winfo_exists():
        chat_janela_ref[0].lift()
        chat_janela_ref[0].focus_force()
        return
    chat_popup = JanelaChat(root, PALETAS, obter_tema_atual, funcoes_chat)
    chat_janela_ref[0] = chat_popup.janela

btn_chat = ctk.CTkButton(
    root, text="💬", font=("Segoe UI", 24),
    fg_color="#f57c00", hover_color="#e65100", width=60, height=60, corner_radius=30,
    command=abrir_chat, cursor="hand2"
)
btn_chat.place(relx=1.0, rely=1.0, x=-90, y=-90)

aplicar_tema(tema_inicial)

if __name__ == "__main__":
    # Inicia módulos em estágios assíncronos não-bloqueantes
    root.after(20, inicializar_modulos)
    root.mainloop()
