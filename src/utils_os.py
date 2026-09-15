import os
import sys
import ctypes
from ctypes import wintypes

ERROR_ALREADY_EXISTS = 183
SW_RESTORE = 9

def focar_janela_existente(titulo_parcial="N.E.X.U.S"):
    """Tenta encontrar a janela existente pelo título e colocá-la em primeiro plano."""
    try:
        user32 = ctypes.windll.user32
        
        def enum_windows_proc(hwnd, extra):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    val = buff.value.lower()
                    if titulo_parcial.lower() in val or "nexus" in val or "conecta hub" in val:
                        user32.ShowWindow(hwnd, SW_RESTORE)
                        user32.SetForegroundWindow(hwnd)
                        return False
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        user32.EnumWindows(WNDENUMPROC(enum_windows_proc), 0)
    except Exception:
        pass

def garantir_instancia_unica(nome_mutex="Nexus_SAP_Singleton_Mutex"):
    """
    Garante que apenas uma instância do aplicativo esteja rodando.
    Retorna (True, handle) se for a primeira instância.
    Retorna (False, None) se já houver outra instância aberta.
    """
    if os.name != "nt":
        return True, None

    try:
        kernel32 = ctypes.windll.kernel32
        mutex = kernel32.CreateMutexW(None, True, nome_mutex)
        last_error = kernel32.GetLastError()
        
        if last_error == ERROR_ALREADY_EXISTS:
            if mutex:
                kernel32.CloseHandle(mutex)
            focar_janela_existente("Conecta Hub")
            return False, None
            
        return True, mutex
    except Exception:
        return True, None

def obter_diretorio_executavel():
    """
    Retorna o diretório onde o executável principal está localizado.
    Se estiver rodando como script Python, retorna o diretório raiz do projeto.
    Usado para arquivos persistentes (Banco de dados, configurações, logins).
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # Se estiver dentro de src/, retorna a pasta raiz do projeto
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def obter_caminho_recurso(*caminhos):
    """
    Retorna o caminho de um recurso interno.
    Se compilado com PyInstaller (--onefile), os recursos estão em sys._MEIPASS.
    Usado para arquivos imutáveis lidos pelo app (ícones, scripts VBS internos).
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, *caminhos)
def obter_diretorio_dados():
    """Retorna um diretório seguro e persistente no AppData do usuário para salvar configs."""
    appdata = os.environ.get("APPDATA")
    if appdata:
        caminho = os.path.join(appdata, "ConectaHub")
    else:
        caminho = os.path.join(os.path.expanduser("~"), ".conectahub")
    os.makedirs(caminho, exist_ok=True)
    return caminho

def obter_caminho_config(nome_arquivo):
    """
    Localiza um arquivo de configuração dando prioridade ao AppData,
    depois à pasta do executável e por fim aos recursos embutidos.
    """
    # 1. Pasta persistente do usuário
    caminho_appdata = os.path.join(obter_diretorio_dados(), nome_arquivo)
    if os.path.exists(caminho_appdata):
        return caminho_appdata
        
    # 2. Pasta local do executável/projeto
    caminho_exe = os.path.join(obter_diretorio_executavel(), "data", nome_arquivo)
    if os.path.exists(caminho_exe):
        return caminho_exe

    # 3. Recurso embutido no PyInstaller
    caminho_rec = obter_caminho_recurso("data", nome_arquivo)
    if os.path.exists(caminho_rec):
        return caminho_rec
        
    return caminho_appdata

def obter_caminho_script(nome_script):
    """
    Localiza um script (.vbs, etc.), priorizando a pasta externa 'scripts/'
    no diretório do executável (para customizações externas) e caindo para os recursos
    embutidos no PyInstaller (sys._MEIPASS) ou pasta 'scripts/' do projeto.
    """
    # 1. Pasta 'scripts' junto ao executável
    caminho_exe = os.path.join(obter_diretorio_executavel(), "scripts", nome_script)
    if os.path.exists(caminho_exe):
        return caminho_exe

    # 2. Diretório raiz do executável
    caminho_exe_raiz = os.path.join(obter_diretorio_executavel(), nome_script)
    if os.path.exists(caminho_exe_raiz):
        return caminho_exe_raiz

    # 3. Recurso embutido no PyInstaller (sys._MEIPASS) ou pasta raiz scripts/ em desenvolvimento
    caminho_rec = obter_caminho_recurso("scripts", nome_script)
    if os.path.exists(caminho_rec):
        return caminho_rec

    # 4. Recurso diretamente na raiz do _MEIPASS
    caminho_rec_raiz = obter_caminho_recurso(nome_script)
    if os.path.exists(caminho_rec_raiz):
        return caminho_rec_raiz

    return caminho_rec


