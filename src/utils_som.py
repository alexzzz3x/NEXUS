"""
utils_som.py — Notificações sonoras e limpeza leve de recursos auxiliares.
"""

def emitir_alerta_sonoro(sucesso=True):
    """
    Toca um sinal sonoro nativo do Windows ao finalizar processos no SAP
    (Bandeirada, OSME, Contagem de Carro, etc.) para que o operador saiba
    que a automação foi concluída mesmo se a janela estiver minimizada.
    """
    try:
        import winsound
        if sucesso:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        else:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    except Exception:
        pass
