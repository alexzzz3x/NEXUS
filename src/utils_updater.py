# -*- coding: utf-8 -*-
"""
utils_updater.py — Sistema de Atualização Automática e Segura do N.E.X.U.S.
Totalmente em conformidade com as regras de segurança corporativa (Drive X aceita apenas .zip).

Fluxo:
1. Verifica na pasta da rede se existe 'versao.json' com versão superior à local.
2. Se houver, exibe pop-up moderno com notas da versão.
3. Ao clicar em 'Atualizar', copia o 'NEXUS.zip' da rede para o %TEMP% local.
4. Spawna um script silencioso (.bat) que aguarda o fechamento do N.E.X.U.S,
   descompacta o novo executável substituindo o local e reinicia o N.E.X.U.S atualizado.
"""

import os
import sys
import json
import re
import shutil
import tempfile
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

import urllib.request

VERSAO_ATUAL = "2.2.1"

# 1. Endereço online no GitHub (prioridade máxima — funciona de qualquer lugar)
URL_VERSAO_GITHUB = "https://raw.githubusercontent.com/Alexzzz1x/N.E.X.U.S/main/versao.json"

# 2. Endereço na rede corporativa (fallback offline/intranet)
PASTA_REDE = r"X:\obras\Elpa - Almox CR Mauá\ANDERSON"
ARQUIVO_VERSAO_REDE = os.path.join(PASTA_REDE, "versao.json")
ARQUIVO_ZIP_REDE = os.path.join(PASTA_REDE, "NEXUS.zip")


def parse_versao(v_str):
    """Converte strings como '2.2.0' ou 'v2.2.1' em tupla de inteiros (2, 2, 0)."""
    if not v_str:
        return (0, 0, 0)
    nums = re.findall(r"\d+", str(v_str))
    return tuple(map(int, nums)) if nums else (0, 0, 0)


def verificar_atualizacao_disponivel():
    """
    Verifica se há uma nova versão disponível.
    1º Tenta via URL online do GitHub (Internet).
    2º Se offline ou inacessível, tenta via pasta da rede X: (Intranet).
    Retorna (tem_nova_versao, dados_versao).
    """
    # 1. Tentativa Online via GitHub
    try:
        req = urllib.request.Request(
            URL_VERSAO_GITHUB,
            headers={"User-Agent": "NEXUS-Desktop-Updater"}
        )
        with urllib.request.urlopen(req, timeout=3.5) as resp:
            if resp.status == 200:
                dados = json.loads(resp.read().decode("utf-8"))
                versao_remota = dados.get("versao", "")
                if parse_versao(versao_remota) > parse_versao(VERSAO_ATUAL):
                    return True, dados
    except Exception:
        pass

    # 2. Fallback na Rede Local (Drive X:)
    try:
        if os.path.exists(ARQUIVO_VERSAO_REDE):
            with open(ARQUIVO_VERSAO_REDE, "r", encoding="utf-8") as f:
                dados = json.load(f)
            versao_remota = dados.get("versao", "")
            if parse_versao(versao_remota) > parse_versao(VERSAO_ATUAL):
                return True, dados
    except Exception:
        pass

    return False, None


class ModalAtualizacao(ctk.CTkToplevel):
    """Pop-up modal moderno anunciando a nova versão do N.E.X.U.S."""
    
    def __init__(self, parent, dados_versao, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.dados_versao = dados_versao
        self.paletas = paletas or {}
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        
        self.title("🚀 Atualização Disponível — N.E.X.U.S")
        self.geometry("540x510")
        self.minsize(500, 470)
        self.resizable(False, False)
        self.attributes("-topmost", True)
        
        # Centraliza na tela
        try:
            self.update_idletasks()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = int((sw - 540) / 2)
            y = int((sh - 510) / 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        self._construir_ui()

    def _construir_ui(self):
        tema = self.obter_tema_atual()
        paleta = self.paletas.get(tema, {
            "bg": "#090d16", "container": "#111827", "texto": "#f8fafc",
            "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"
        })
        
        self.configure(fg_color=paleta.get("bg", "#090d16"))

        # 1. Cabeçalho (Topo)
        f_top = ctk.CTkFrame(self, fg_color=paleta.get("container", "#111827"), height=65, corner_radius=0)
        f_top.pack(fill="x", side="top")
        f_top.pack_propagate(False)

        lbl_tit = ctk.CTkLabel(
            f_top,
            text="🚀 NOVA ATUALIZAÇÃO DISPONÍVEL!",
            font=("Segoe UI", 16, "bold"),
            text_color=paleta.get("accent", "#00d2ff")
        )
        lbl_tit.pack(side="left", padx=20, pady=18)

        # 2. Rodapé / Botões (Empacotado no BOTTOM antes do corpo para nunca sumir ou ser cortado)
        f_botoes = ctk.CTkFrame(self, fg_color="transparent")
        f_botoes.pack(fill="x", side="bottom", padx=20, pady=(0, 18))

        self.btn_atualizar = ctk.CTkButton(
            f_botoes,
            text="🚀 Atualizar Agora",
            font=("Segoe UI", 12, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=40,
            command=self._iniciar_atualizacao
        )
        self.btn_atualizar.pack(side="right", padx=(8, 0))

        self.btn_depois = ctk.CTkButton(
            f_botoes,
            text="Lembrar Mais Tarde",
            font=("Segoe UI", 11),
            fg_color=paleta.get("input", "#1a2436"),
            hover_color=paleta.get("container", "#111827"),
            text_color="#94a3b8",
            height=40,
            command=self.destroy
        )
        self.btn_depois.pack(side="right")

        # 3. Container Central (Preenche todo o espaço intermediário restante)
        f_corpo = ctk.CTkFrame(self, fg_color="transparent")
        f_corpo.pack(fill="both", expand=True, padx=20, pady=(12, 10))

        versao_remota = self.dados_versao.get("versao", "Nova")
        data_remota = self.dados_versao.get("data", "")
        notas = self.dados_versao.get("notas", "Melhorias gerais e correções de estabilidade.")

        f_chips = ctk.CTkFrame(f_corpo, fg_color=paleta.get("container", "#111827"), corner_radius=8)
        f_chips.pack(fill="x", pady=(0, 10))

        lbl_ver = ctk.CTkLabel(
            f_chips,
            text=f"Versão Atual: v{VERSAO_ATUAL}   ➔   Nova Versão: v{versao_remota}",
            font=("Segoe UI", 13, "bold"),
            text_color="#10b981"
        )
        lbl_ver.pack(padx=14, pady=(8, 2), anchor="w")

        if data_remota:
            lbl_dt = ctk.CTkLabel(
                f_chips,
                text=f"Publicada em: {data_remota}",
                font=("Segoe UI", 10),
                text_color="#94a3b8"
            )
            lbl_dt.pack(padx=14, pady=(0, 8), anchor="w")

        # Caixa com Notas da Versão
        lbl_not_tit = ctk.CTkLabel(f_corpo, text="📝 O que mudou nesta versão:", font=("Segoe UI", 11, "bold"), text_color=paleta.get("texto", "#f8fafc"))
        lbl_not_tit.pack(anchor="w", pady=(0, 4))

        txt_notas = ctk.CTkTextbox(f_corpo, fg_color=paleta.get("container", "#111827"), font=("Segoe UI", 11), height=105, corner_radius=8)
        txt_notas.pack(fill="both", expand=True, pady=(0, 10))
        txt_notas.insert("1.0", notas)
        txt_notas.configure(state="disabled")

        # Status / Barra de Progresso durante a instalação
        self.lbl_status = ctk.CTkLabel(
            f_corpo,
            text="Clique em 'Atualizar Agora' para baixar e aplicar automaticamente.",
            font=("Segoe UI", 10),
            text_color="#94a3b8"
        )
        self.lbl_status.pack(anchor="w", pady=(0, 3))

        self.progress_bar = ctk.CTkProgressBar(f_corpo, progress_color=paleta.get("accent", "#00d2ff"), height=8)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)

    def _iniciar_atualizacao(self):
        self.btn_atualizar.configure(state="disabled", text="⏳ Atualizando...")
        self.btn_depois.configure(state="disabled")
        self.progress_bar.set(0)
        self.lbl_status.configure(text="Iniciando download da atualização...", text_color="#38bdf8")

        def _thread_copia():
            try:
                if not getattr(sys, "frozen", False):
                    # Modo Desenvolvimento (.py)
                    self.after(0, lambda: messagebox.showinfo(
                        "Modo Desenvolvimento",
                        "Você está executando o código-fonte Python (.py).\nA atualização automática é aplicada apenas no executável compilado (.exe)."
                    ))
                    self.after(0, self.destroy)
                    return

                caminho_exe = os.path.abspath(sys.executable)
                dir_exe = os.path.dirname(caminho_exe)
                url_down = self.dados_versao.get("url_download", "")

                pasta_temp = tempfile.gettempdir()
                zip_local = os.path.join(pasta_temp, "NEXUS_UPDATE_TEMP.zip")

                # 1. Download Online (HTTP/HTTPS) ou Cópia de Rede
                if url_down.startswith("http://") or url_down.startswith("https://"):
                    self.after(0, lambda: self.lbl_status.configure(text="Conectando ao servidor...", text_color="#38bdf8"))
                    req = urllib.request.Request(url_down, headers={"User-Agent": "NEXUS-Desktop-Updater"})
                    with urllib.request.urlopen(req, timeout=90) as resp, open(zip_local, "wb") as out_f:
                        total_len = resp.headers.get("Content-Length")
                        total_size = int(total_len) if total_len and total_len.isdigit() else 0
                        dl_bytes = 0
                        chunk_size = 65536
                        while True:
                            chunk = resp.read(chunk_size)
                            if not chunk:
                                break
                            out_f.write(chunk)
                            dl_bytes += len(chunk)
                            if total_size > 0:
                                pct = min(1.0, dl_bytes / total_size)
                                self.after(0, lambda p=pct, b=dl_bytes, t=total_size: (
                                    self.progress_bar.set(p),
                                    self.lbl_status.configure(text=f"Baixando: {b // 1024} KB / {t // 1024} KB ({int(p*100)}%)")
                                ))
                else:
                    caminho_origem = url_down if (url_down and os.path.exists(url_down)) else ARQUIVO_ZIP_REDE
                    if not os.path.exists(caminho_origem):
                        raise FileNotFoundError(f"Arquivo de atualização não encontrado:\n{caminho_origem}")
                    self.after(0, lambda: (
                        self.progress_bar.start(),
                        self.lbl_status.configure(text="Copiando pacote da rede local... Por favor, aguarde.", text_color="#38bdf8")
                    ))
                    shutil.copy2(caminho_origem, zip_local)

                self.after(0, lambda: (
                    self.progress_bar.stop(),
                    self.progress_bar.set(1.0),
                    self.lbl_status.configure(text="Aplicando nova versão e reiniciando...", text_color="#10b981")
                ))

                # 2. Gera o script silencioso .bat que substitui o executável
                script_bat = os.path.join(pasta_temp, "aplicar_nexus_update.bat")
                conteudo_bat = f"""@echo off
title Atualizando N.E.X.U.S...
echo Aguardando fechamento do N.E.X.U.S...
timeout /t 2 /nobreak > nul
taskkill /F /IM "N.E.X.U.S.exe" > nul 2>&1
timeout /t 1 /nobreak > nul

echo Extraindo nova versao...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '{zip_local}' -DestinationPath '{dir_exe}' -Force"

echo Limpando arquivos temporarios...
del /F /Q "{zip_local}" > nul 2>&1

echo Reiniciando N.E.X.U.S...
start "" "{caminho_exe}"
del /F /Q "%~f0" > nul 2>&1
exit
"""
                with open(script_bat, "w", encoding="latin-1") as f:
                    f.write(conteudo_bat)

                # 3. Dispara o .bat desacoplado e encerra a instância atual
                subprocess.Popen(["cmd.exe", "/c", script_bat], shell=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
                
                # Encerra o processo do Python/NEXUS
                os._exit(0)

            except Exception as erro:
                self.after(0, lambda e=erro: self._falha_atualizacao(e))

        t = threading.Thread(target=_thread_copia, daemon=True)
        t.start()

    def _falha_atualizacao(self, erro):
        self.progress_bar.stop()
        self.btn_atualizar.configure(state="normal", text="Tentar Novamente")
        self.btn_depois.configure(state="normal")
        self.lbl_status.configure(text=f"❌ Erro ao atualizar: {erro}", text_color="#ef4444")
        messagebox.showerror("Erro na Atualização", f"Não foi possível concluir a atualização automática:\n{erro}", parent=self)


def iniciar_checagem_atualizacao(root, paletas=None, obter_tema_atual=None):
    """Inicia a verificação de versão em segundo plano silencioso."""
    def _verificar():
        tem_nova, dados = verificar_atualizacao_disponivel()
        if tem_nova and dados:
            root.after(0, lambda: ModalAtualizacao(root, dados, paletas, obter_tema_atual))
            
    threading.Thread(target=_verificar, daemon=True).start()


def checar_atualizacao_manual(root, paletas=None, obter_tema_atual=None):
    """Disparado quando o usuário clica em 'Verificar Atualizações' no menu/sidebar."""
    def _verificar():
        tem_nova, dados = verificar_atualizacao_disponivel()
        if tem_nova and dados:
            root.after(0, lambda: ModalAtualizacao(root, dados, paletas, obter_tema_atual))
        else:
            root.after(0, lambda: messagebox.showinfo(
                "N.E.X.U.S Atualizado",
                f"Você já está utilizando a versão mais recente do N.E.X.U.S (v{VERSAO_ATUAL})."
            ))
            
    threading.Thread(target=_verificar, daemon=True).start()
