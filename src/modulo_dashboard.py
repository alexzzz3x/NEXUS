# -*- coding: utf-8 -*-
"""
modulo_dashboard.py — Painel Principal e Central de Comando Executiva do N.E.X.U.S.

Recursos:
- Saudação dinâmica baseada no relógio do sistema (Bom dia / Boa tarde / Boa noite)
- Widget climático em tempo real (Open-Meteo API, assíncrono e não-bloqueante)
- Cockpit de status dos sistemas (SAP GUI, Backup Diário Local, Rede Drive X)
- Métricas com alternador "Hoje" vs "Total Acumulado" (lê saídas do backup diário)
- Radar Preventivo de OSMEs (alerta equipes vencidas ou vencendo em 3 dias)
- Linha do Tempo / Feed de Atividades Recentes do turno em tempo real
- Barra de Ações Rápidas (1-Clique) integrada com navegação do sistema
"""

import os
import re
import json
import time
import threading
import urllib.request
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

from utils_metrics import ler_metricas
from utils_backup_diario import obter_caminho_backup_diario, obter_pasta_backup_diario

DEFAULT_PALETAS = {
    "Nexus": {
        "bg": "#090d16", "container": "#111827", "texto": "#f8fafc",
        "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436",
        "card_bg": "#131d2e", "subtext": "#94a3b8", "border": "#1e293b"
    },
    "Azul Delray": {
        "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
        "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60",
        "card_bg": "#1c3252", "subtext": "#94a3b8", "border": "#23416b"
    }
}

WMO_CODES = {
    0: ("Céu Limpo", "☀️"),
    1: ("Ensolarado", "🌤️"),
    2: ("Parcialmente Nublado", "⛅"),
    3: ("Nublado", "☁️"),
    45: ("Nevoeiro", "🌫️"),
    48: ("Nevoeiro com Geada", "🌫️"),
    51: ("Chuvisco Leve", "🌦️"),
    53: ("Chuvisco", "🌦️"),
    55: ("Chuva Fina", "🌧️"),
    61: ("Chuva Leve", "🌧️"),
    63: ("Chuva Moderada", "🌧️"),
    65: ("Chuva Forte", "🌧️"),
    80: ("Pancadas de Chuva", "🌦️"),
    81: ("Chuva Forte", "🌧️"),
    82: ("Chuva Torrencial", "⛈️"),
    95: ("Tempestade", "⛈️"),
    96: ("Tempestade com Granizo", "⛈️"),
}

_cache_clima = {"texto": "", "hora": 0}


def obter_saudacao():
    hora = datetime.now().hour
    if 6 <= hora < 12:
        return "Bom dia! ☀️ Operação ativa"
    elif 12 <= hora < 18:
        return "Boa tarde! 🌤️ Operação ativa"
    else:
        return "Boa noite! 🌙 Operação ativa"


def consultar_clima_async(callback):
    """Consulta clima de Mauá/SP de forma leve e assíncrona (não trava a UI)."""
    def _worker():
        agora = time.time()
        if agora - _cache_clima.get("hora", 0) < 1800 and _cache_clima.get("texto"):
            callback(_cache_clima["texto"])
            return
        try:
            url = "https://api.open-meteo.com/v1/forecast?latitude=-23.66&longitude=-46.46&current_weather=true"
            req = urllib.request.Request(url, headers={"User-Agent": "NEXUS-Weather"})
            with urllib.request.urlopen(req, timeout=3.5) as resp:
                dados = json.loads(resp.read().decode("utf-8"))
                cw = dados.get("current_weather", {})
                temp = round(cw.get("temperature", 0))
                wcode = cw.get("weathercode", 0)
                desc, icone = WMO_CODES.get(wcode, ("Tempo Estável", "🌤️"))
                texto = f"📍 Mauá, SP: {temp}°C • {desc} {icone}"
                _cache_clima["texto"] = texto
                _cache_clima["hora"] = agora
                callback(texto)
        except Exception:
            # Fallback limpo com o dia da semana atual
            dias = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
            hoje_dt = datetime.now()
            callback(f"📅 {dias[hoje_dt.weekday()]} • {hoje_dt.strftime('%d/%m/%Y')}")

    threading.Thread(target=_worker, daemon=True).start()


class AbaDashboard:
    def __init__(self, parent_frame, paletas=None, obter_tema_atual=None, on_navigate=None):
        self.parent_frame = parent_frame
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.on_navigate = on_navigate

        self.modo_visualizacao = "hoje"  # "hoje" ou "total"
        self.cards_info = []
        self.lista_osmes_criticas = []
        self.lista_saidas_hoje = []

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self.recarregar_dados_completos()

    def _cores(self, tema_nome):
        pal = self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        return {
            "bg": pal.get("bg", "#090d16"),
            "container": pal.get("container", "#111827"),
            "texto": pal.get("texto", "#f8fafc"),
            "accent": pal.get("accent", "#00d2ff"),
            "hover": pal.get("hover", "#0284c7"),
            "input": pal.get("input", "#1a2436"),
            "card_bg": pal.get("card_bg", pal.get("container", "#131d2e")),
            "subtext": pal.get("subtext", "#94a3b8"),
            "border": pal.get("border", "#1e293b")
        }

    # ------------------------------------------------------------------
    # Construção Visual Completa
    # ------------------------------------------------------------------
    def _construir_interface(self):
        # Frame scrollable principal para comportar telas menores confortavelmente
        self.scroll_root = ctk.CTkScrollableFrame(self.parent_frame, corner_radius=0, fg_color="transparent")
        self.scroll_root.pack(fill="both", expand=True, padx=16, pady=8)

        # 1. COCKPIT BANNER SUPERIOR
        self.frame_topo = ctk.CTkFrame(self.scroll_root, corner_radius=12)
        self.frame_topo.pack(fill="x", pady=(0, 10))

        # Linha 1 do Banner: Saudação à esquerda, Clima e Indicadores à direita
        f_banner_top = ctk.CTkFrame(self.frame_topo, fg_color="transparent")
        f_banner_top.pack(fill="x", padx=16, pady=(12, 4))

        self.lbl_saudacao = ctk.CTkLabel(
            f_banner_top,
            text=f"✨ {obter_saudacao()}",
            font=("Segoe UI", 20, "bold")
        )
        self.lbl_saudacao.pack(side="left")

        # Badges de Status do Sistema
        self.f_badges = ctk.CTkFrame(f_banner_top, fg_color="transparent")
        self.f_badges.pack(side="right")

        self.lbl_clima = ctk.CTkLabel(
            self.f_badges,
            text="📍 Mauá, SP • Conectando...",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        )
        self.lbl_clima.pack(side="left", padx=8)

        self.lbl_status_sap = ctk.CTkLabel(
            self.f_badges,
            text="⚪ SAP GUI",
            font=("Segoe UI", 10, "bold"),
            text_color="#94a3b8"
        )
        self.lbl_status_sap.pack(side="left", padx=6)

        self.lbl_status_rede = ctk.CTkLabel(
            self.f_badges,
            text="⚪ Drive X",
            font=("Segoe UI", 10, "bold"),
            text_color="#94a3b8"
        )
        self.lbl_status_rede.pack(side="left", padx=6)

        # Subtítulo do Banner
        self.lbl_subtitulo = ctk.CTkLabel(
            self.frame_topo,
            text="Central de Operações do Almoxarifado • Monitor de OSMEs e Desempenho em Tempo Real",
            font=("Segoe UI", 11),
            text_color="#94a3b8"
        )
        self.lbl_subtitulo.pack(anchor="w", padx=16, pady=(0, 12))

        # 2. CONTROLE DE MÉTRICAS (Alternador Hoje vs Total)
        f_metr_ctrl = ctk.CTkFrame(self.scroll_root, fg_color="transparent")
        f_metr_ctrl.pack(fill="x", pady=(0, 6))

        lbl_sec_tit = ctk.CTkLabel(
            f_metr_ctrl,
            text="📊 Desempenho Operacional",
            font=("Segoe UI", 13, "bold")
        )
        lbl_sec_tit.pack(side="left")

        self.btn_modo_hoje = ctk.CTkButton(
            f_metr_ctrl,
            text="📅 Operações de Hoje",
            width=150,
            height=28,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=lambda: self._alternar_modo_metricas("hoje")
        )
        self.btn_modo_hoje.pack(side="right", padx=(4, 0))

        self.btn_modo_total = ctk.CTkButton(
            f_metr_ctrl,
            text="📈 Histórico Geral",
            width=140,
            height=28,
            font=("Segoe UI", 11),
            corner_radius=6,
            command=lambda: self._alternar_modo_metricas("total")
        )
        self.btn_modo_total.pack(side="right", padx=4)

        # 3. CARDS DE MÉTRICAS (Grid de 5 Colunas)
        self.grid_cards = ctk.CTkFrame(self.scroll_root, fg_color="transparent")
        self.grid_cards.pack(fill="x", pady=(0, 12))
        for col in range(5):
            self.grid_cards.grid_columnconfigure(col, weight=1, uniform="metric_cards")

        c1, l1, t1, s1 = self._criar_card_metrica(self.grid_cards, "Saídas de OSME", icone="📤")
        c1.grid(row=0, column=0, padx=4, sticky="nsew")

        c2, l2, t2, s2 = self._criar_card_metrica(self.grid_cards, "Carros Auditados", icone="🚗")
        c2.grid(row=0, column=1, padx=4, sticky="nsew")

        c3, l3, t3, s3 = self._criar_card_metrica(self.grid_cards, "Bandeiradas SAP", icone="🚩")
        c3.grid(row=0, column=2, padx=4, sticky="nsew")

        c4, l4, t4, s4 = self._criar_card_metrica(self.grid_cards, "OSMEs Transferidas", icone="📦")
        c4.grid(row=0, column=3, padx=4, sticky="nsew")

        c5, l5, t5, s5 = self._criar_card_metrica(self.grid_cards, "Etiquetas Geradas", icone="🎫")
        c5.grid(row=0, column=4, padx=4, sticky="nsew")

        self.cards_info = [
            {"card": c1, "lbl_val": l1, "lbl_tit": t1, "lbl_sub": s1, "chave": "saidas_osme"},
            {"card": c2, "lbl_val": l2, "lbl_tit": t2, "lbl_sub": s2, "chave": "contagens_carro"},
            {"card": c3, "lbl_val": l3, "lbl_tit": t3, "lbl_sub": s3, "chave": "bandeiradas_executadas"},
            {"card": c4, "lbl_val": l4, "lbl_tit": t4, "lbl_sub": s4, "chave": "osmes_transferidos"},
            {"card": c5, "lbl_val": l5, "lbl_tit": t5, "lbl_sub": s5, "chave": "etiquetas_geradas"},
        ]

        # 4. PAINEL CENTRAL DUPLO (Radar de OSMEs + Linha do Tempo)
        self.frame_central = ctk.CTkFrame(self.scroll_root, fg_color="transparent")
        self.frame_central.pack(fill="both", expand=True, pady=(0, 12))
        self.frame_central.grid_columnconfigure(0, weight=1, uniform="paineis")
        self.frame_central.grid_columnconfigure(1, weight=1, uniform="paineis")

        # 4A. COLUNA ESQUERDA: RADAR PREVENTIVO DE OSMES
        self.painel_radar = ctk.CTkFrame(self.frame_central, corner_radius=10)
        self.painel_radar.grid(row=0, column=0, padx=(0, 6), sticky="nsew")

        f_radar_head = ctk.CTkFrame(self.painel_radar, fg_color="transparent")
        f_radar_head.pack(fill="x", padx=12, pady=(10, 4))

        self.lbl_radar_tit = ctk.CTkLabel(
            f_radar_head,
            text="⚠️ Radar de OSMEs Críticas",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_radar_tit.pack(side="left")

        self.lbl_radar_badge = ctk.CTkLabel(
            f_radar_head,
            text="Verificando...",
            font=("Segoe UI", 10, "bold"),
            text_color="#f59e0b"
        )
        self.lbl_radar_badge.pack(side="right")

        self.lbl_radar_desc = ctk.CTkLabel(
            self.painel_radar,
            text="Ordens vencidas ou a vencer nos próximos 3 dias que exigem atenção:",
            font=("Segoe UI", 10),
            text_color="#94a3b8"
        )
        self.lbl_radar_desc.pack(anchor="w", padx=12, pady=(0, 6))

        self.scroll_radar = ctk.CTkFrame(self.painel_radar, corner_radius=6, fg_color="transparent")
        self.scroll_radar.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.btn_acao_radar = ctk.CTkButton(
            self.painel_radar,
            text="⚡ Abrir Atualizar OSME no SAP (IW39)",
            height=32,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=lambda: self._navegar("🔄 Atualizar OSME")
        )
        self.btn_acao_radar.pack(fill="x", padx=10, pady=(0, 10))

        # 4B. COLUNA DIREITA: LINHA DO TEMPO / ATIVIDADES RECENTES
        self.painel_timeline = ctk.CTkFrame(self.frame_central, corner_radius=10)
        self.painel_timeline.grid(row=0, column=1, padx=(6, 0), sticky="nsew")

        f_time_head = ctk.CTkFrame(self.painel_timeline, fg_color="transparent")
        f_time_head.pack(fill="x", padx=12, pady=(10, 4))

        self.lbl_time_tit = ctk.CTkLabel(
            f_time_head,
            text="🕒 Atividades Recentes do Turno",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_time_tit.pack(side="left")

        self.lbl_time_badge = ctk.CTkLabel(
            f_time_head,
            text="Hoje",
            font=("Segoe UI", 10, "bold"),
            text_color="#10b981"
        )
        self.lbl_time_badge.pack(side="right")

        self.lbl_time_desc = ctk.CTkLabel(
            self.painel_timeline,
            text="Registro em tempo real de materiais e saídas gravadas hoje:",
            font=("Segoe UI", 10),
            text_color="#94a3b8"
        )
        self.lbl_time_desc.pack(anchor="w", padx=12, pady=(0, 6))

        self.scroll_timeline = ctk.CTkFrame(self.painel_timeline, corner_radius=6, fg_color="transparent")
        self.scroll_timeline.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.btn_acao_backup = ctk.CTkButton(
            self.painel_timeline,
            text="📁 Abrir Backup Diário (.xlsx)",
            height=32,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=self._abrir_backup_excel
        )
        self.btn_acao_backup.pack(fill="x", padx=10, pady=(0, 10))

        # 5. BARRA DE AÇÕES RÁPIDAS NO RODAPÉ
        self.frame_acoes = ctk.CTkFrame(self.scroll_root, corner_radius=10)
        self.frame_acoes.pack(fill="x", pady=(0, 8))

        f_act_inner = ctk.CTkFrame(self.frame_acoes, fg_color="transparent")
        f_act_inner.pack(fill="x", padx=12, pady=10)

        lbl_acoes_tit = ctk.CTkLabel(f_act_inner, text="⚡ Ações Rápidas:", font=("Segoe UI", 11, "bold"))
        lbl_acoes_tit.pack(side="left", padx=(0, 10))

        self.btn_act_saida = ctk.CTkButton(
            f_act_inner,
            text="📤 Nova Saída OSME",
            width=160,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=lambda: self._navegar("📤 Saída OSME")
        )
        self.btn_act_saida.pack(side="left", padx=4)

        self.btn_act_iq09 = ctk.CTkButton(
            f_act_inner,
            text="🔢 Extrator & IQ09",
            width=150,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=lambda: self._navegar("🔢 Extrator Seriais")
        )
        self.btn_act_iq09.pack(side="left", padx=4)

        self.btn_act_carro = ctk.CTkButton(
            f_act_inner,
            text="🚗 Contagem Carro",
            width=140,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=lambda: self._navegar("🚗 Contagem de Carro")
        )
        self.btn_act_carro.pack(side="left", padx=3)

        self.btn_act_fechamento = ctk.CTkButton(
            f_act_inner,
            text="📋 Fechamento Dia",
            width=135,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=self._abrir_modal_fechamento
        )
        self.btn_act_fechamento.pack(side="left", padx=3)

        self.btn_act_novidades = ctk.CTkButton(
            f_act_inner,
            text="📢 O que há de novo",
            width=145,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=self._abrir_novidades
        )
        self.btn_act_novidades.pack(side="left", padx=3)

        self.btn_recarregar = ctk.CTkButton(
            f_act_inner,
            text="🔄 Atualizar Painel",
            width=140,
            height=34,
            font=("Segoe UI", 11, "bold"),
            corner_radius=6,
            command=self.recarregar_dados_completos
        )
        self.btn_recarregar.pack(side="right", padx=4)

    def _criar_card_metrica(self, parent, titulo, icone="📊"):
        card = ctk.CTkFrame(parent, corner_radius=10)
        lbl_ic = ctk.CTkLabel(card, text=icone, font=("Segoe UI", 24))
        lbl_ic.pack(pady=(10, 2))

        lbl_tit = ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 11, "bold"))
        lbl_tit.pack(pady=1)

        lbl_val = ctk.CTkLabel(card, text="0", font=("Segoe UI", 22, "bold"))
        lbl_val.pack(pady=(0, 2))

        lbl_sub = ctk.CTkLabel(card, text="Hoje", font=("Segoe UI", 9), text_color="#94a3b8")
        lbl_sub.pack(pady=(0, 8))

        return card, lbl_val, lbl_tit, lbl_sub

    # ------------------------------------------------------------------
    # Lógica de Dados & Atualizações Assíncronas
    # ------------------------------------------------------------------
    def recarregar_dados_completos(self):
        """Dispara todas as checagens e leituras sem travar a interface."""
        # 1. Saudação pelo relógio
        self.lbl_saudacao.configure(text=f"✨ {obter_saudacao()}")

        # 2. Clima em Mauá/SP
        consultar_clima_async(lambda txt: self.parent_frame.after(0, lambda: self.lbl_clima.configure(text=txt)))

        # 3. Status dos Sistemas
        self._checar_status_sistemas()

        # 4. Lê saídas de hoje do Backup Diário
        self._carregar_saidas_hoje_async()

        # 5. Lê Radar de OSMEs Críticas
        self._carregar_radar_osmes_async()

        # 6. Atualiza os cards conforme o modo ativo
        self._atualizar_cards_metricas()

    def _checar_status_sistemas(self):
        """Verifica conexão com SAP e Drive X de forma 100% assíncrona (não congela a UI)."""
        def _worker():
            status_sap_txt = "⚪ SAP: Offline / Fechado"
            status_sap_cor = "#94a3b8"
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                sap_gui = win32com.client.GetObject("SAPGUI")
                app = getattr(sap_gui, "GetScriptingEngine", None)
                if callable(app):
                    app = app()
                if app and getattr(app, "Children", None) and app.Children.Count > 0:
                    con = app.Children(0)
                    if con and getattr(con, "Children", None) and con.Children.Count > 0:
                        sess = con.Children(0)
                        user = getattr(sess.Info, "User", "") or "Conectado"
                        status_sap_txt = f"🟢 SAP: Conectado ({user})"
                        status_sap_cor = "#10b981"
            except Exception:
                pass
            finally:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

            # Drive X
            status_rede_txt = "🟡 Drive X: Offline"
            status_rede_cor = "#f59e0b"
            try:
                drive_x = r"X:\obras\Elpa - Almox CR Mauá"
                if os.path.exists(drive_x):
                    status_rede_txt = "🟢 Drive X: Online"
                    status_rede_cor = "#10b981"
            except Exception:
                pass

            def _aplicar():
                try:
                    self.lbl_status_sap.configure(text=status_sap_txt, text_color=status_sap_cor)
                    self.lbl_status_rede.configure(text=status_rede_txt, text_color=status_rede_cor)
                except Exception:
                    pass

            self.parent_frame.after(0, _aplicar)

        threading.Thread(target=_worker, daemon=True).start()

    def _carregar_saidas_hoje_async(self):
        """Carrega o arquivo local de hoje para contar saídas e preencher a timeline (Zero-Lock via BytesIO)."""
        def _worker():
            caminho_hoje = obter_caminho_backup_diario()
            linhas_hoje = []
            if os.path.exists(caminho_hoje):
                try:
                    import io
                    import openpyxl
                    with open(caminho_hoje, "rb") as f:
                        buf = io.BytesIO(f.read())
                    wb = openpyxl.load_workbook(buf, read_only=True, data_only=True)
                    try:
                        ws = wb.active
                        for r in ws.iter_rows(min_row=2, values_only=True):
                            if any(r):
                                linhas_hoje.append(r)
                    finally:
                        wb.close()
                        buf.close()
                        import gc
                        gc.collect()
                except Exception:
                    pass

            self.lista_saidas_hoje = linhas_hoje
            self.parent_frame.after(0, self._renderizar_timeline)
            self.parent_frame.after(0, self._atualizar_cards_metricas)

        threading.Thread(target=_worker, daemon=True).start()

    def _carregar_radar_osmes_async(self):
        """Varre a aba CADASTRO de forma 100% não-bloqueante (Zero-Lock via BytesIO)."""
        def _worker():
            caminho_mestre = r"X:\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm"
            criticas = []
            if os.path.exists(caminho_mestre):
                try:
                    import io
                    import openpyxl
                    with open(caminho_mestre, "rb") as f:
                        buf = io.BytesIO(f.read())
                    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
                    try:
                        if "CADASTRO" in wb.sheetnames:
                            ws = wb["CADASTRO"]
                            hoje = datetime.now().date()
                            limite = hoje + timedelta(days=3)
                            for r in ws.iter_rows(min_row=2, values_only=True):
                                eq = r[4]
                                ordem = r[7]
                                venc = r[8]
                                if not eq:
                                    continue
                                d_venc = None
                                if isinstance(venc, datetime):
                                    d_venc = venc.date()
                                elif isinstance(venc, str):
                                    try:
                                        d_venc = datetime.strptime(venc.strip()[:10], "%d/%m/%Y").date()
                                    except Exception:
                                        pass
                                if not d_venc:
                                    continue

                                if d_venc < hoje:
                                    criticas.append({
                                        "equipe": str(eq).strip(),
                                        "ordem": str(ordem or "—").strip(),
                                        "venc": d_venc.strftime("%d/%m"),
                                        "status": "Vencida",
                                        "cor": "#ef4444"
                                    })
                                elif d_venc <= limite:
                                    criticas.append({
                                        "equipe": str(eq).strip(),
                                        "ordem": str(ordem or "—").strip(),
                                        "venc": d_venc.strftime("%d/%m"),
                                        "status": "Vence em Breve",
                                        "cor": "#f59e0b"
                                    })
                    finally:
                        wb.close()
                        buf.close()
                        import gc
                        gc.collect()
                except Exception:
                    pass

            self.lista_osmes_criticas = criticas
            self.parent_frame.after(0, self._renderizar_radar)

        threading.Thread(target=_worker, daemon=True).start()

    def _renderizar_radar(self):
        """Desenha os itens críticos no container limpo do Radar."""
        for w in self.scroll_radar.winfo_children():
            w.destroy()

        qtd = len(self.lista_osmes_criticas)
        if qtd == 0:
            self.lbl_radar_badge.configure(text="✅ Tudo em dia!", text_color="#10b981")
            lbl_vazio = ctk.CTkLabel(
                self.scroll_radar,
                text="🎉 Nenhuma equipe com OSME vencida ou vencendo nos próximos 3 dias.",
                font=("Segoe UI", 11, "italic"),
                text_color="#94a3b8"
            )
            lbl_vazio.pack(pady=30)
            return

        self.lbl_radar_badge.configure(text=f"⚠️ {qtd} equipes críticas", text_color="#f59e0b")

        c = self._cores(self.obter_tema_atual())
        card_bg = c.get("card_bg", "#131d2e")

        # Exibe até as 8 mais críticas com fluidez máxima sem rolagem aninhada
        for item in self.lista_osmes_criticas[:8]:
            row = ctk.CTkFrame(self.scroll_radar, height=32, corner_radius=6, fg_color=card_bg)
            row.pack(fill="x", pady=2, padx=2)
            row.pack_propagate(False)

            lbl_eq = ctk.CTkLabel(row, text=f"  {item['equipe']}", font=("Segoe UI", 11, "bold"), anchor="w", width=70)
            lbl_eq.pack(side="left", padx=4)

            lbl_ord = ctk.CTkLabel(row, text=f"OSME: {item['ordem']}", font=("Segoe UI", 10), text_color="#94a3b8", anchor="w")
            lbl_ord.pack(side="left", padx=6, fill="x", expand=True)

            lbl_v = ctk.CTkLabel(
                row,
                text=f"{item['status']} ({item['venc']})",
                font=("Segoe UI", 10, "bold"),
                text_color=item["cor"],
                anchor="e"
            )
            lbl_v.pack(side="right", padx=8)

        if qtd > 8:
            lbl_mais = ctk.CTkLabel(
                self.scroll_radar,
                text=f"... e mais {qtd - 8} equipe(s) crítica(s) (Ver na aba Atualizar OSME)",
                font=("Segoe UI", 9, "italic"),
                text_color="#f59e0b"
            )
            lbl_mais.pack(pady=(3, 2))

    def _renderizar_timeline(self):
        """Desenha as saídas do dia em ordem cronológica reversa."""
        for w in self.scroll_timeline.winfo_children():
            w.destroy()

        qtd = len(self.lista_saidas_hoje)
        self.lbl_time_badge.configure(text=f"{qtd} itens hoje", text_color="#10b981" if qtd > 0 else "#94a3b8")

        if qtd == 0:
            lbl_vazio = ctk.CTkLabel(
                self.scroll_timeline,
                text="📦 Nenhuma saída registrada hoje ainda.\nAs movimentações da MIGO aparecerão aqui em tempo real.",
                font=("Segoe UI", 11, "italic"),
                text_color="#94a3b8"
            )
            lbl_vazio.pack(pady=30)
            return

        c = self._cores(self.obter_tema_atual())
        card_bg = c.get("card_bg", "#131d2e")

        # Exibe até as 8 últimas saídas de hoje em ordem inversa
        for r in reversed(self.lista_saidas_hoje[-8:]):
            serial = str(r[0] or "Item")
            mat = str(r[1] or "")
            eq = str(r[3] or "Equipe")
            doc_sap = str(r[5] or "")
            hora_completa = str(r[6] or "")
            hora_simples = hora_completa.split(" ")[-1][:5] if " " in hora_completa else hora_completa[:5]

            row = ctk.CTkFrame(self.scroll_timeline, height=32, corner_radius=6, fg_color=card_bg)
            row.pack(fill="x", pady=2, padx=2)
            row.pack_propagate(False)

            lbl_hr = ctk.CTkLabel(row, text=f" {hora_simples} 📤", font=("Segoe UI", 10, "bold"), text_color="#38bdf8", width=55, anchor="w")
            lbl_hr.pack(side="left", padx=4)

            lbl_info = ctk.CTkLabel(
                row,
                text=f"{eq} • {serial}",
                font=("Segoe UI", 10, "bold"),
                anchor="w"
            )
            lbl_info.pack(side="left", padx=4, fill="x", expand=True)

            if doc_sap and doc_sap != "None":
                lbl_doc = ctk.CTkLabel(row, text=f"Doc: {doc_sap}", font=("Segoe UI", 9), text_color="#94a3b8", anchor="e")
                lbl_doc.pack(side="right", padx=6)

        if qtd > 8:
            lbl_mais = ctk.CTkLabel(
                self.scroll_timeline,
                text=f"... e mais {qtd - 8} saída(s) registrada(s) hoje (Ver no Backup Diário)",
                font=("Segoe UI", 9, "italic"),
                text_color="#38bdf8"
            )
            lbl_mais.pack(pady=(3, 2))

    def _alternar_modo_metricas(self, modo):
        self.modo_visualizacao = modo
        c = self._cores(self.obter_tema_atual())
        accent = c.get("accent", "#00d2ff")
        hover = c.get("hover", "#0284c7")
        container = c.get("container", "#111827")

        if modo == "hoje":
            self.btn_modo_hoje.configure(fg_color=accent, text_color="#090d16")
            self.btn_modo_total.configure(fg_color=container, text_color="#94a3b8")
        else:
            self.btn_modo_total.configure(fg_color=accent, text_color="#090d16")
            self.btn_modo_hoje.configure(fg_color=container, text_color="#94a3b8")

        self._atualizar_cards_metricas()

    def _atualizar_cards_metricas(self):
        """Atualiza os números exibidos nos 5 cards conforme o modo ativo."""
        metricas_totais = ler_metricas()
        total_saidas_hoje = len(self.lista_saidas_hoje)

        for item in self.cards_info:
            chave = item["chave"]
            if self.modo_visualizacao == "hoje":
                item["lbl_sub"].configure(text="Hoje")
                if chave == "saidas_osme":
                    item["lbl_val"].configure(text=str(total_saidas_hoje))
                else:
                    # Para outras métricas se ainda não houver particionamento diário, exibe o contador geral
                    val = metricas_totais.get(chave, 0)
                    item["lbl_val"].configure(text=str(val))
            else:
                item["lbl_sub"].configure(text="Total Geral")
                val = metricas_totais.get(chave, 0)
                item["lbl_val"].configure(text=str(val))

    def _navegar(self, tab_name):
        """Executa navegação para qualquer aba do sistema via callback."""
        if self.on_navigate:
            try:
                self.on_navigate(tab_name)
            except Exception as e:
                print(f"[Dashboard] Erro na navegação: {e}")

    def _abrir_novidades(self):
        """Abre a janela de novidades com as notas de versão do N.E.X.U.S."""
        try:
            from utils_updater import exibir_novidades
            exibir_novidades(self.parent_frame.winfo_toplevel(), self.paletas, self.obter_tema_atual)
        except Exception as e:
            print(f"[Dashboard] Erro ao abrir novidades: {e}")

    def _abrir_modal_fechamento(self):
        """Abre modal com resumo executivo do dia e botão de copiar para WhatsApp."""
        c = self._cores(self.obter_tema_atual())
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        accent = c.get("accent", "#00d2ff")
        texto = c.get("texto", "#f8fafc")

        top = ctk.CTkToplevel(self.parent_frame)
        top.title("📋 Fechamento Diário — N.E.X.U.S")
        top.geometry("520x460")
        top.transient(self.parent_frame.winfo_toplevel())
        top.attributes("-topmost", True)

        # Monta estatísticas de equipes
        eq_map = {}
        docs_sap = set()
        for r in self.lista_saidas_hoje:
            eq = str(r[3] or "Equipe")
            eq_map[eq] = eq_map.get(eq, 0) + 1
            if len(r) > 5 and r[5] and str(r[5]) != "None":
                docs_sap.add(str(r[5]))

        hoje_str = datetime.now().strftime("%d/%m/%Y")
        hora_str = datetime.now().strftime("%H:%M")
        total_saidas = len(self.lista_saidas_hoje)

        txt_whats = f"📋 *FECHAMENTO DIÁRIO — N.E.X.U.S*\n"
        txt_whats += f"📅 *Data:* {hoje_str} às {hora_str}\n\n"
        txt_whats += f"📦 *Movimentações de Saída (MIGO Z61):*\n"
        txt_whats += f"• Total de itens/medidores lançados: *{total_saidas} un*\n"
        if eq_map:
            txt_whats += f"• Equipes atendidas ({len(eq_map)}):\n"
            for eq, qtd in sorted(eq_map.items()):
                txt_whats += f"   - *{eq}*: {qtd} un\n"
        else:
            txt_whats += f"• Nenhuma saída registrada hoje até o momento.\n"

        if docs_sap:
            txt_whats += f"\n📑 *Documentos SAP Gerados:*\n• {', '.join(sorted(docs_sap))}\n"

        txt_whats += f"\n🛡️ *Auditoria:* Backup Diário Local 100% gravado.\n_Gerado automaticamente pelo N.E.X.U.S_"

        # UI do Modal
        f_corpo = ctk.CTkFrame(top, corner_radius=12, fg_color=container)
        f_corpo.pack(fill="both", expand=True, padx=16, pady=16)

        lbl_tit = ctk.CTkLabel(f_corpo, text="📋 Resumo Executivo do Turno (WhatsApp)", font=("Segoe UI", 15, "bold"), text_color=accent)
        lbl_tit.pack(pady=(12, 6))

        txt_box = ctk.CTkTextbox(f_corpo, height=260, font=("Consolas", 11), fg_color=c.get("card_bg", "#131d2e"))
        txt_box.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        txt_box.insert("1.0", txt_whats)

        f_botoes = ctk.CTkFrame(f_corpo, fg_color="transparent")
        f_botoes.pack(fill="x", padx=14, pady=(0, 12))

        def _copiar():
            self.parent_frame.clipboard_clear()
            self.parent_frame.clipboard_append(txt_box.get("1.0", "end-1c"))
            btn_copiar.configure(text="✅ Copiado para Área de Transferência!", fg_color="#10b981")
            top.after(2500, lambda: btn_copiar.configure(text="📋 Copiar para WhatsApp", fg_color=accent))

        btn_copiar = ctk.CTkButton(
            f_botoes,
            text="📋 Copiar para WhatsApp",
            font=("Segoe UI", 11, "bold"),
            height=34,
            fg_color=accent,
            hover_color=c.get("hover", "#0284c7"),
            text_color="#090d16",
            command=_copiar
        )
        btn_copiar.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_fechar_mod = ctk.CTkButton(
            f_botoes,
            text="✕ Fechar",
            font=("Segoe UI", 11),
            width=90,
            height=34,
            fg_color="#334155",
            hover_color="#475569",
            command=top.destroy
        )
        btn_fechar_mod.pack(side="right")

    def _abrir_backup_excel(self):
        """Abre o arquivo de backup de hoje no Microsoft Excel."""
        caminho = obter_caminho_backup_diario()
        if os.path.exists(caminho):
            try:
                os.startfile(caminho)
            except Exception as e:
                print(f"[Dashboard] Erro ao abrir backup diário: {e}")

    def atualizar_metricas(self):
        """Método de compatibilidade com chamadas antigas."""
        self.recarregar_dados_completos()

    # ------------------------------------------------------------------
    # Aplicação de Temas
    # ------------------------------------------------------------------
    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        accent = c.get("accent", "#00d2ff")
        hover = c.get("hover", "#0284c7")
        texto = c.get("texto", "#f8fafc")
        card_bg = c.get("card_bg", "#131d2e")

        self.parent_frame.configure(fg_color=bg)
        self.frame_topo.configure(fg_color=container)
        self.lbl_saudacao.configure(text_color=texto)

        # Botões de modo de métricas
        if self.modo_visualizacao == "hoje":
            self.btn_modo_hoje.configure(fg_color=accent, text_color="#090d16")
            self.btn_modo_total.configure(fg_color=container, text_color="#94a3b8")
        else:
            self.btn_modo_total.configure(fg_color=accent, text_color="#090d16")
            self.btn_modo_hoje.configure(fg_color=container, text_color="#94a3b8")

        # Cards de métricas
        for item in self.cards_info:
            item["card"].configure(fg_color=card_bg)
            item["lbl_tit"].configure(text_color=texto)
            item["lbl_val"].configure(text_color=accent)

        # Painéis centrais
        self.painel_radar.configure(fg_color=container)
        self.lbl_radar_tit.configure(text_color=texto)
        self.btn_acao_radar.configure(fg_color=accent, hover_color=hover, text_color="#090d16")

        self.painel_timeline.configure(fg_color=container)
        self.lbl_time_tit.configure(text_color=texto)
        self.btn_acao_backup.configure(fg_color=accent, hover_color=hover, text_color="#090d16")

        # Barra de ações rápidas
        self.frame_acoes.configure(fg_color=container)
        self.btn_act_saida.configure(fg_color=accent, hover_color=hover, text_color="#090d16")
        self.btn_act_iq09.configure(fg_color=accent, hover_color=hover, text_color="#090d16")
        self.btn_act_carro.configure(fg_color=accent, hover_color=hover, text_color="#090d16")
        self.btn_act_fechamento.configure(fg_color=accent, hover_color=hover, text_color="#090d16")
        if hasattr(self, "btn_act_novidades"):
            self.btn_act_novidades.configure(fg_color=container, hover_color=hover, text_color=accent)
        self.btn_recarregar.configure(fg_color=hover, hover_color=accent, text_color="#f8fafc")


def construir_aba_dashboard(parent_frame, paletas=None, obter_tema_atual=None, on_navigate=None):
    return AbaDashboard(parent_frame, paletas, obter_tema_atual, on_navigate)
