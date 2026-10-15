# -*- coding: utf-8 -*-
"""
Módulo de Gerenciamento da Bandeja do Sistema (Windows System Tray) e Fechamento Seguro.

Funcionalidades:
1. Ao fechar a janela principal (botão 'X'):
   - Se houver pendências de gravação de planilha na fila, alerta o usuário e recomenda minimizar para a bandeja.
   - Permite que o N.E.X.U.S continue executando em segundo plano na bandeja do Windows,
     garantindo que os timers de auto-sincronização gravem as pendências assim que o Excel for liberado.
2. Ícone na bandeja do sistema com:
   - Duplo clique para restaurar a janela.
   - Menu de contexto: 'Abrir N.E.X.U.S', 'Sincronizar Planilha Agora', 'Encerrar N.E.X.U.S'.
3. Resiliente: se pystray não estiver disponível, realiza o fallback transparente para janelas Tkinter normais.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from PIL import Image

from utils_os import obter_caminho_recurso
from utils_pendencias_planilha import (
    obter_quantidade_pendencias,
    processar_todas_pendencias
)


class ModalConfirmarFechamento(ctk.CTkToplevel):
    """
    Modal de confirmação customizado para quando o usuário clica no 'X'.
    Oferece opções claras:
    - Minimizar para a Bandeja (Segundo Plano)
    - Encerrar Definitivamente
    - Cancelar
    """
    def __init__(self, parent, qtd_pendencias=0, callback_minimizar=None, callback_encerrar=None):
        super().__init__(parent)
        self.callback_minimizar = callback_minimizar
        self.callback_encerrar = callback_encerrar

        self.title("N.E.X.U.S — Fechar ou Minimizar?")
        self.geometry("480x280")
        self.resizable(False, False)
        self.configure(fg_color="#0f172a")

        self.transient(parent)
        self.grab_set()

        # Centraliza na tela
        self.update_idletasks()
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 240
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 140
            self.geometry(f"+{max(10, x)}+{max(10, y)}")
        except Exception:
            pass

        self._construir_ui(qtd_pendencias)

    def _construir_ui(self, qtd_pendencias):
        f_top = ctk.CTkFrame(self, fg_color="transparent")
        f_top.pack(fill="x", padx=20, pady=(20, 10))

        if qtd_pendencias > 0:
            icone = "⚠️"
            cor_titulo = "#f59e0b"
            titulo = f"Atenção: {qtd_pendencias} Operação(ões) na Fila de Pendências!"
            subtitulo = (
                f"Existem {qtd_pendencias} operação(ões) aguardando a liberação da planilha por outro usuário.\n\n"
                "Se você encerrar o programa agora, a gravação automática será pausada.\n"
                "Recomendamos minimizar para a bandeja para que o N.E.X.U.S conclua a gravação sozinho em segundo plano."
            )
        else:
            icone = "🛡️"
            cor_titulo = "#38bdf8"
            titulo = "Como deseja proceder com o N.E.X.U.S?"
            subtitulo = (
                "Você pode minimizar o aplicativo para a bandeja do Windows (ao lado do relógio)\n"
                "para manter ativas as integrações e sincronizações em segundo plano,\n"
                "ou encerrar completamente o sistema."
            )

        lbl_tit = ctk.CTkLabel(
            f_top,
            text=f"{icone} {titulo}",
            font=("Segoe UI", 13, "bold"),
            text_color=cor_titulo,
            wraplength=440,
            justify="left"
        )
        lbl_tit.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(
            self,
            text=subtitulo,
            font=("Segoe UI", 11),
            text_color="#cbd5e1",
            wraplength=440,
            justify="left"
        )
        lbl_sub.pack(padx=20, pady=(5, 20), anchor="w")

        # Botões de Ação
        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", side="bottom", padx=20, pady=(0, 20))

        btn_min = ctk.CTkButton(
            f_btns,
            text="📥 Minimizar p/ Bandeja (Recomendado)",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=36,
            command=self._acao_minimizar
        )
        btn_min.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_exit = ctk.CTkButton(
            f_btns,
            text="❌ Encerrar",
            font=("Segoe UI", 11),
            fg_color="#dc2626",
            hover_color="#b91c1c",
            width=100,
            height=36,
            command=self._acao_encerrar
        )
        btn_exit.pack(side="left", padx=(0, 6))

        btn_canc = ctk.CTkButton(
            f_btns,
            text="Cancelar",
            font=("Segoe UI", 11),
            fg_color="#334155",
            hover_color="#475569",
            width=80,
            height=36,
            command=self.destroy
        )
        btn_canc.pack(side="left")

    def _acao_minimizar(self):
        self.destroy()
        if self.callback_minimizar:
            self.callback_minimizar()

    def _acao_encerrar(self):
        self.destroy()
        if self.callback_encerrar:
            self.callback_encerrar()


class NexusTrayManager:
    """Gerencia a criação do ícone de bandeja e integração com o ciclo de vida da janela do Tkinter."""
    def __init__(self, root, icone_path=None, callback_on_sync=None, callback_on_exit=None):
        self.root = root
        self.icone_path = icone_path or obter_caminho_recurso("assets", "Nexus.ico")
        self.callback_on_sync = callback_on_sync
        self.callback_on_exit = callback_on_exit
        self.tray_icon = None
        self.minimizado_na_bandeja = False
        self._lock = threading.Lock()

        # Vincula o evento WM_DELETE_WINDOW
        self.root.protocol("WM_DELETE_WINDOW", self.solicitar_fechamento_janela)

    def _obter_imagem_icone(self):
        try:
            if self.icone_path and os.path.exists(self.icone_path):
                return Image.open(self.icone_path)
        except Exception:
            pass
        return Image.new("RGB", (64, 64), color=(0, 210, 255))

    def solicitar_fechamento_janela(self):
        """Disparado quando o usuário clica no 'X' da janela."""
        qtd = 0
        try:
            qtd = obter_quantidade_pendencias()
        except Exception:
            pass

        # Abre o modal de escolha
        ModalConfirmarFechamento(
            parent=self.root,
            qtd_pendencias=qtd,
            callback_minimizar=self.minimizar_para_bandeja,
            callback_encerrar=self.encerrar_aplicacao
        )

    def minimizar_para_bandeja(self, notificar=True):
        """Oculta a janela principal do Tkinter e inicia o ícone de bandeja."""
        with self._lock:
            try:
                import pystray
                from pystray import MenuItem as item, Menu
            except ImportError:
                # Se pystray não estiver presente, apenas minimiza na barra de tarefas
                self.root.iconify()
                return

            if self.tray_icon is not None:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
                self.tray_icon = None

            self.root.withdraw()
            self.minimizado_na_bandeja = True

            img = self._obter_imagem_icone()

            menu = Menu(
                item("🖥️ Abrir N.E.X.U.S", self._on_menu_restaurar, default=True),
                item("🔄 Sincronizar Planilha Agora", self._on_menu_sincronizar),
                Menu.SEPARATOR,
                item("❌ Encerrar N.E.X.U.S", self._on_menu_encerrar)
            )

            qtd = obter_quantidade_pendencias()
            tooltip = f"N.E.X.U.S — Ativo em 2º plano ({qtd} pendências)" if qtd > 0 else "N.E.X.U.S — Automação SAP"

            self.tray_icon = pystray.Icon("nexus_hub", img, tooltip, menu)

            try:
                self.tray_icon.run_detached()
            except Exception as e:
                print(f"Erro ao iniciar tray icon: {e}")
                self.root.deiconify()
                self.minimizado_na_bandeja = False
                return

            if notificar:
                try:
                    msg = (
                        f"N.E.X.U.S continua em segundo plano com {qtd} operação(ões) aguardando gravação na planilha."
                        if qtd > 0 else
                        "N.E.X.U.S foi minimizado para a bandeja e continua ativo em segundo plano."
                    )
                    self.tray_icon.notify(
                        msg,
                        title="N.E.X.U.S Ativo em Segundo Plano"
                    )
                except Exception:
                    pass

    def restaurar_janela(self):
        """Restaura a janela principal e remove o ícone da bandeja."""
        with self._lock:
            if self.tray_icon is not None:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
                self.tray_icon = None

            self.minimizado_na_bandeja = False
            self.root.after(0, self._restaurar_tk)

    def _restaurar_tk(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def _on_menu_restaurar(self, icon, item):
        self.restaurar_janela()

    def _on_menu_sincronizar(self, icon, item):
        try:
            sucessos, restantes, msg = processar_todas_pendencias()
            if icon:
                if sucessos > 0 and restantes == 0:
                    icon.notify(f"🎉 Todas as {sucessos} pendências foram gravadas na planilha com sucesso!", title="Sincronização Concluída")
                elif sucessos > 0:
                    icon.notify(f"✓ {sucessos} gravadas. Restam {restantes} ainda bloqueadas.", title="Sincronização Parcial")
                else:
                    icon.notify("A planilha ainda continua aberta por outro usuário.", title="Planilha em Uso")
        except Exception as e:
            print(f"Erro ao sincronizar via tray: {e}")

    def _on_menu_encerrar(self, icon, item):
        qtd = obter_quantidade_pendencias()
        if qtd > 0:
            # Precisa de confirmação para não perder gravação
            def _confirmar():
                self.restaurar_janela()
                resp = messagebox.askyesno(
                    "Encerrar N.E.X.U.S?",
                    f"⚠️ ATENÇÃO: Há {qtd} operação(ões) na Fila de Pendências aguardando gravação na planilha!\n\n"
                    "Se você encerrar o programa agora, elas não serão gravadas automaticamente até que você reabra o N.E.X.U.S.\n\n"
                    "Deseja realmente encerrar agora?",
                    icon="warning"
                )
                if resp:
                    self.encerrar_aplicacao()
                else:
                    self.minimizar_para_bandeja(notificar=False)

            self.root.after(0, _confirmar)
        else:
            self.encerrar_aplicacao()

    def encerrar_aplicacao(self):
        """Encerra a aplicação de forma limpa."""
        with self._lock:
            if self.tray_icon is not None:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
                self.tray_icon = None

        if self.callback_on_exit:
            try:
                self.callback_on_exit()
            except Exception:
                pass

        try:
            self.root.destroy()
        except Exception:
            pass
        sys.exit(0)
