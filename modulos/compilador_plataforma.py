# -*- coding: utf-8 -*-
"""
Módulo de compilação e consolidação de Fichas da Plataforma (Manhã x Tarde).
Processa dados brutos, agrupa fichas por segmento e turnos operacionais,
e gera o relatório Excel oficial formatado e estilizado idêntico ao modelo original.
"""

import os
import re
import unicodedata
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def _normalizar_texto(texto):
    if not isinstance(texto, str):
        texto = str(texto) if pd.notna(texto) else ""
    nfkd = unicodedata.normalize("NFKD", texto)
    s = "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()
    return re.sub(r"[\s_\-]+", "_", s)


def _parse_datetime(valor):
    if isinstance(valor, datetime):
        return valor
    if pd.isna(valor) or valor is None:
        return None
    s = str(valor).strip()
    if not s or s.lower() in ("none", "nan", "nat", "-"):
        return None

    formatos = [
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    try:
        dt = pd.to_datetime(s, dayfirst=True)
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        pass
    return None


def carregar_dados_brutos(caminho_arquivo):
    """Lê arquivo bruto em formato Excel (.xlsx, .xlsm, .xls) ou CSV (.csv)."""
    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    ext = os.path.splitext(caminho_arquivo)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True, read_only=True)
        sheet_target = "DADOS_BRUTOS" if "DADOS_BRUTOS" in wb.sheetnames else ("DADOS BRUTOS" if "DADOS BRUTOS" in wb.sheetnames else wb.sheetnames[0])
        wb.close()
        df = pd.read_excel(caminho_arquivo, sheet_name=sheet_target)
    elif ext == ".csv":
        try:
            df = pd.read_csv(caminho_arquivo, sep=";", encoding="utf-8")
        except Exception:
            try:
                df = pd.read_csv(caminho_arquivo, sep=",", encoding="utf-8")
            except Exception:
                df = pd.read_csv(caminho_arquivo, sep=";", encoding="latin1")
    else:
        raise ValueError("Formato de arquivo não suportado. Use .xlsx, .xlsm, .xls ou .csv")

    return df


def mapear_colunas(df):
    """Mapeia dinamicamente as colunas do arquivo bruto para os campos padrão."""
    mapa = {}
    colunas_originais = df.columns.tolist()

    regras = {
        "ficha": [r"^ficha", r"^nr_ficha", r"^num_ficha", r"^id_ficha", r"^ordem", r"^os"],
        "base": [r"^base", r"^cidade", r"^unidade"],
        "area": [r"^area", r"^polo", r"^regiao", r"^segmento"],
        "data": [r"^data$", r"^dt_abertura", r"^data_abertura"],
        "recurso": [r"^recurso", r"^equipe", r"^veiculo", r"^placa"],
        "re_1": [r"^re_1", r"^re1"],
        "nome_1": [r"^nome_1", r"^nome1"],
        "re_2": [r"^re_2", r"^re2"],
        "nome_2": [r"^nome_2", r"^nome2"],
        "equipamento": [r"^equipamento", r"^material", r"^item", r"^descricao"],
        "pedido": [r"^pedido", r"^qtd_ped", r"^quantidade_ped", r"^qtd"],
        "fornecido": [r"^fornecido", r"^qtd_forn"],
        "devolucao": [r"^devolucao", r"^devolvido"],
        "obs": [r"^obs", r"^observacao"],
        "baixa_sistema": [r"^baixa_sistema", r"^baixa"],
        "finalizado": [r"^finalizado", r"^status", r"^concluido"],
        "criado_por": [r"^criado_por", r"^aberto_por"],
        "criado_dh": [r"^criado_data_hora", r"^data_hora_criacao", r"^data_abertura_hora", r"^criado_em", r"^data_criacao"],
        "alterado_por": [r"^alterado_por", r"^encerrado_por", r"^finalizado_por"],
        "alterado_dh": [r"^alterado_data_hora", r"^data_hora_alteracao", r"^data_hora_encerramento", r"^alterado_em", r"^data_encerramento", r"^data_finalizacao"]
    }

    for col in colunas_originais:
        col_norm = _normalizar_texto(col)
        for padrao_nome, patterns in regras.items():
            if padrao_nome not in mapa:
                for pat in patterns:
                    if re.search(pat, col_norm):
                        mapa[padrao_nome] = col
                        break

    return mapa


def processar_fichas(df):
    """
    Processa e consolida as fichas brutas.
    Retorna dicionário com:
      - df_detalhada: DataFrame consolidado (1 linha por ficha)
      - df_resumo: Tabela resumo (LESTE, SUL, CLANDESTINA, TOTAL)
      - metricas: KPIs globais
      - df_bruto: DataFrame original limpo
    """
    col_map = mapear_colunas(df)
    if "ficha" not in col_map:
        raise KeyError("Coluna 'Ficha' não foi encontrada no arquivo de dados brutos.")

    col_ficha = col_map["ficha"]
    df_valido = df[df[col_ficha].notna() & (df[col_ficha].astype(str).str.strip() != "")].copy()
    df_valido[col_ficha] = df_valido[col_ficha].astype(str).str.strip()

    col_base = col_map.get("base")
    col_area = col_map.get("area")
    col_data = col_map.get("data")
    col_recurso = col_map.get("recurso")
    col_re1 = col_map.get("re_1")
    col_nome1 = col_map.get("nome_1")
    col_re2 = col_map.get("re_2")
    col_nome2 = col_map.get("nome_2")
    col_equip = col_map.get("equipamento")
    col_pedido = col_map.get("pedido")
    col_baixa = col_map.get("baixa_sistema")
    col_final = col_map.get("finalizado")
    col_criado_por = col_map.get("criado_por")
    col_criado_dh = col_map.get("criado_dh")
    col_alterado_por = col_map.get("alterado_por")
    col_alterado_dh = col_map.get("alterado_dh")

    fichas_dict = {}

    for _, row in df_valido.iterrows():
        f_id = str(row[col_ficha]).strip()
        if f_id not in fichas_dict:
            fichas_dict[f_id] = {
                "Ficha": f_id,
                "Data": row[col_data] if col_data else "",
                "Base": str(row[col_base] or "").strip() if col_base else "",
                "Area": str(row[col_area] or "").strip() if col_area else "",
                "Recurso": str(row[col_recurso] or "").strip().upper() if col_recurso else "",
                "RE 1": row[col_re1] if col_re1 else "",
                "Nome 1": str(row[col_nome1] or "").strip() if col_nome1 else "",
                "RE 2": row[col_re2] if col_re2 else "",
                "Nome 2": str(row[col_nome2] or "").strip() if col_nome2 else "",
                "Equipamentos": [],
                "Baixa Sistema": str(row[col_baixa] or "").strip() if col_baixa else "",
                "Finalizado": str(row[col_final] or "").strip() if col_final else "",
                "Criado Por": str(row[col_criado_por] or "").strip() if col_criado_por else "",
                "Criado Data Hora": _parse_datetime(row[col_criado_dh]) if col_criado_dh else None,
                "Alterado Por": str(row[col_alterado_por] or "").strip() if col_alterado_por else "",
                "Alterado Data Hora": _parse_datetime(row[col_alterado_dh]) if col_alterado_dh else None,
            }

        if col_equip:
            eq = str(row[col_equip] or "").strip()
            qtd = str(row[col_pedido] or "").strip() if col_pedido else ""
            if eq and eq.lower() not in ("nan", "none", ""):
                desc = f"{eq} ({qtd})" if qtd and qtd.lower() not in ("nan", "none", "") else eq
                if desc not in fichas_dict[f_id]["Equipamentos"]:
                    fichas_dict[f_id]["Equipamentos"].append(desc)

    linhas_detalhada = []
    totais_por_segmento = {
        "LESTE": {"total": 0, "manha": 0, "tarde": 0},
        "SUL": {"total": 0, "manha": 0, "tarde": 0},
        "CLANDESTINA": {"total": 0, "manha": 0, "tarde": 0},
    }
    outros_segmentos = {}

    total_fichas = len(fichas_dict)
    total_finalizadas = 0
    total_manha = 0
    total_tarde = 0
    total_rapidas = 0
    soma_tempo_rapidas = 0.0

    recursos_manha_dict = {}
    dias_manha_dict = {}

    for f_id, f in fichas_dict.items():
        equip_str = "; ".join(f["Equipamentos"])
        finalizado_str = f["Finalizado"]
        is_finalizado = str(finalizado_str).strip().lower() in ("sim", "true", "1", "s", "concluido", "finalizado")
        finalizado_norm = "Sim" if is_finalizado else "Não"

        dt_criado = f["Criado Data Hora"]
        dt_alterado = f["Alterado Data Hora"]

        periodo_criacao = ""
        if dt_criado:
            periodo_criacao = "Manha" if dt_criado.hour < 12 else "Tarde"

        periodo_finalizacao = "Nao finalizada"
        intervalo_min = 0.0
        mesma_manha = ""
        rapida = "Nao"

        if is_finalizado and dt_alterado:
            total_finalizadas += 1
            if dt_alterado.hour < 12:
                periodo_finalizacao = "Manha"
                total_manha += 1
            else:
                periodo_finalizacao = "Tarde"
                total_tarde += 1

            if dt_criado:
                diff_sec = (dt_alterado - dt_criado).total_seconds()
                intervalo_min = round(diff_sec / 60.0, 1)

            if periodo_criacao == "Manha" and periodo_finalizacao == "Manha":
                if dt_criado and dt_alterado and dt_criado.date() == dt_alterado.date():
                    mesma_manha = "Sim"
                    if 0 <= intervalo_min <= 120:
                        rapida = "Sim"
                        total_rapidas += 1
                        soma_tempo_rapidas += intervalo_min

            # Agregação para os gráficos auxiliares do Dashboard
            if periodo_finalizacao == "Manha":
                rec = f["Recurso"].upper()
                if rec not in recursos_manha_dict:
                    recursos_manha_dict[rec] = [0, 0]  # [total_manha, rapidas]
                recursos_manha_dict[rec][0] += 1
                if rapida == "Sim":
                    recursos_manha_dict[rec][1] += 1

                dia_str = dt_alterado.strftime("%d/%m")
                if dia_str not in dias_manha_dict:
                    dias_manha_dict[dia_str] = [0, 0]
                dias_manha_dict[dia_str][0] += 1
                if rapida == "Sim":
                    dias_manha_dict[dia_str][1] += 1

        recurso = f["Recurso"].upper()
        area = f["Area"].upper()

        if recurso.startswith("RCL"):
            segmento = "CLANDESTINA"
        elif "LESTE" in area:
            segmento = "LESTE"
        elif "SUL" in area:
            segmento = "SUL"
        else:
            segmento = area if (area and area != "NAN") else "OUTROS"

        if segmento in totais_por_segmento:
            totais_por_segmento[segmento]["total"] += 1
            if is_finalizado:
                if periodo_finalizacao == "Manha":
                    totais_por_segmento[segmento]["manha"] += 1
                elif periodo_finalizacao == "Tarde":
                    totais_por_segmento[segmento]["tarde"] += 1
        else:
            if segmento not in outros_segmentos:
                outros_segmentos[segmento] = {"total": 0, "manha": 0, "tarde": 0}
            outros_segmentos[segmento]["total"] += 1
            if is_finalizado:
                if periodo_finalizacao == "Manha":
                    outros_segmentos[segmento]["manha"] += 1
                elif periodo_finalizacao == "Tarde":
                    outros_segmentos[segmento]["tarde"] += 1

        linhas_detalhada.append({
            "Ficha": f["Ficha"],
            "Data": f["Data"],
            "Base": f["Base"],
            "Area": f["Area"],
            "Recurso": f["Recurso"],
            "Segmento": segmento,
            "RE 1": f["RE 1"],
            "Nome 1": f["Nome 1"],
            "RE 2": f["RE 2"],
            "Nome 2": f["Nome 2"],
            "Equipamentos": equip_str,
            "Baixa Sistema": f["Baixa Sistema"],
            "Finalizado": finalizado_norm,
            "Criado Por": f["Criado Por"],
            "Criado Data Hora": dt_criado,
            "Alterado Por": f["Alterado Por"],
            "Alterado Data Hora": dt_alterado,
            "Periodo Criacao": periodo_criacao,
            "Periodo Finalizacao": periodo_finalizacao,
            "Mesma Manha": mesma_manha,
            "Intervalo Minutos": intervalo_min if is_finalizado else None,
            "Rapida Manha 120min": rapida,
        })

    df_det = pd.DataFrame(linhas_detalhada)

    resumo_linhas = []
    ordem_padrao = ["LESTE", "SUL", "CLANDESTINA"] + sorted(outros_segmentos.keys())

    soma_total = 0
    soma_manha = 0
    soma_tarde = 0

    for seg in ordem_padrao:
        dados_seg = totais_por_segmento.get(seg) or outros_segmentos.get(seg, {"total": 0, "manha": 0, "tarde": 0})
        tot = dados_seg["total"]
        m = dados_seg["manha"]
        t = dados_seg["tarde"]
        pct_m = (m / tot) if tot > 0 else 0.0
        pct_t = (t / tot) if tot > 0 else 0.0

        soma_total += tot
        soma_manha += m
        soma_tarde += t

        resumo_linhas.append({
            "SEGMENTO": seg,
            "TOTAL FICHAS": tot,
            "ENCERRADAS MANHA": m,
            "% MANHA": pct_m,
            "ENCERRADAS A TARDE": t,
            "% TARDE": pct_t
        })

    pct_total_m = (soma_manha / soma_total) if soma_total > 0 else 0.0
    pct_total_t = (soma_tarde / soma_total) if soma_total > 0 else 0.0

    resumo_linhas.append({
        "SEGMENTO": "TOTAL",
        "TOTAL FICHAS": soma_total,
        "ENCERRADAS MANHA": soma_manha,
        "% MANHA": pct_total_m,
        "ENCERRADAS A TARDE": soma_tarde,
        "% TARDE": pct_total_t
    })

    df_resumo = pd.DataFrame(resumo_linhas)
    tempo_medio_rapidas = (soma_tempo_rapidas / total_rapidas) if total_rapidas > 0 else 0.0

    metricas = {
        "total_fichas": total_fichas,
        "total_finalizadas": total_finalizadas,
        "finalizadas_manha": total_manha,
        "finalizadas_tarde": total_tarde,
        "rapidas_manha": total_rapidas,
        "tempo_medio_rapidas": round(tempo_medio_rapidas, 1),
        "pct_manha_leste": (totais_por_segmento["LESTE"]["manha"] / totais_por_segmento["LESTE"]["total"]) if totais_por_segmento["LESTE"]["total"] > 0 else 0.0,
        "recursos_manha": recursos_manha_dict,
        "dias_manha": dias_manha_dict
    }

    return {
        "df_detalhada": df_det,
        "df_resumo": df_resumo,
        "metricas": metricas,
        "df_bruto": df_valido
    }


def obter_caminho_template():
    """Localiza o arquivo modelo oficial (exemplo.xlsm)."""
    candidatos = [
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "TEMPLATE_RELATORIO_PLATAFORMA.xlsm"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "TEMPLATE_RELATORIO_PLATAFORMA.xlsm"),
        os.path.join(os.path.expanduser("~"), "Downloads", "exemplo.xlsm"),
        r"C:\Users\anderson.vieira\Desktop\Planilha do pedro\PLANILHA_PLATAFORMA_AUTOMATICA.xlsm"
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return None


def exportar_relatorio_excel(resultado, caminho_saida):
    """
    Exporta o relatório preenchendo o modelo oficial completo (exemplo.xlsm).
    Preserva todas as 5 abas (Dashboard, DADOS_BRUTOS, Pivot Finalizadas, Pivot Segmentos, Detalhada),
    gráficos, fórmulas e macros, exatamente como solicitado.
    """
    template_path = obter_caminho_template()

    # Se encontrar o template oficial, injeta os dados nele preservando 100% da estrutura
    if template_path and os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path, keep_vba=True)

        df_bruto = resultado["df_bruto"]
        df_det = resultado["df_detalhada"]
        metricas = resultado["metricas"]

        # 1. ABA DADOS_BRUTOS
        ws_raw = wb["DADOS_BRUTOS"] if "DADOS_BRUTOS" in wb.sheetnames else wb["DADOS BRUTOS"]
        # Limpa linhas antigas
        for r in range(2, max(ws_raw.max_row + 1, len(df_bruto) + 10)):
            for c in range(1, 21):
                ws_raw.cell(r, c).value = None

        # Preenche linhas novas
        col_names = df_bruto.columns.tolist()
        for r_idx, row_val in enumerate(df_bruto.itertuples(index=False), 2):
            for c_idx, val in enumerate(row_val, 1):
                if c_idx <= 20:
                    c = ws_raw.cell(r_idx, c_idx)
                    if isinstance(val, datetime):
                        c.value = val
                        c.number_format = "dd/mm/yyyy hh:mm:ss"
                    elif pd.isna(val) or val is None or str(val).lower() == "nan":
                        c.value = ""
                    else:
                        c.value = val

        ws_raw["W2"].value = f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | {len(df_det)} fichas consolidadas"

        # 2. ABA DETALHADA
        ws_det = wb["Detalhada"]
        cols_det = [
            "Ficha", "Data", "Base", "Area", "Recurso", "RE 1", "Nome 1", "RE 2", "Nome 2",
            "Equipamentos", "Baixa Sistema", "Finalizado", "Criado Por", "Criado Data Hora",
            "Alterado Por", "Alterado Data Hora", "Periodo Criacao", "Periodo Finalizacao",
            "Mesma Manha", "Intervalo Minutos", "Rapida Manha 120min"
        ]

        # Limpa linhas antigas
        for r in range(2, max(ws_det.max_row + 1, len(df_det) + 10)):
            for c in range(1, 22):
                ws_det.cell(r, c).value = None

        for r_idx, row_dict in enumerate(df_det.to_dict(orient="records"), 2):
            for c_idx, col_k in enumerate(cols_det, 1):
                val = row_dict.get(col_k)
                cell = ws_det.cell(r_idx, c_idx)
                if isinstance(val, datetime):
                    cell.value = val
                    cell.number_format = "dd/mm/yyyy hh:mm:ss"
                elif col_k == "Intervalo Minutos" and val is not None:
                    cell.value = float(val)
                    cell.number_format = "0.0"
                elif pd.isna(val) or val is None or str(val).lower() == "nan":
                    cell.value = ""
                else:
                    cell.value = val

        # Atualiza o tamanho da Tabela tbBaseDetalhada
        if "tbBaseDetalhada" in ws_det.tables:
            ws_det.tables["tbBaseDetalhada"].ref = f"A1:U{len(df_det) + 1}"

        # 3. ABA DASHBOARD (Atualiza auxiliares de gráficos)
        ws_dash = wb["Dashboard"]
        # Limpa e atualiza AF3:AH100 (Top recursos manhã)
        for r in range(3, 100):
            for c in range(32, 35):
                ws_dash.cell(r, c).value = None

        recursos_manha = metricas.get("recursos_manha", {})
        recursos_ordenados = sorted(recursos_manha.items(), key=lambda x: (x[1][0], x[1][1]), reverse=True)[:10]
        for idx, (rec, vals) in enumerate(recursos_ordenados, 3):
            ws_dash.cell(idx, 32).value = rec
            ws_dash.cell(idx, 33).value = vals[0]
            ws_dash.cell(idx, 34).value = vals[1]

        # Limpa e atualiza AK3:AM100 (Dias manhã)
        for r in range(3, 100):
            for c in range(37, 40):
                ws_dash.cell(r, c).value = None

        dias_manha = metricas.get("dias_manha", {})
        for idx, (dia, vals) in enumerate(sorted(dias_manha.items()), 3):
            ws_dash.cell(idx, 37).value = dia
            ws_dash.cell(idx, 38).value = vals[0]
            ws_dash.cell(idx, 39).value = vals[1]

        # 4. ABA PIVOT SEGMENTOS (Garante que os valores calculados também estejam presentes)
        ws_seg = wb["Pivot Segmentos"]
        df_resumo = resultado["df_resumo"]
        for idx, r in df_resumo.iterrows():
            seg_nome = str(r["SEGMENTO"]).strip()
            # Mapeia linhas: 5=LESTE, 6=SUL, 7=CLANDESTINA, 8=TOTAL
            if seg_nome == "LESTE":
                r_num = 5
            elif seg_nome == "SUL":
                r_num = 6
            elif seg_nome == "CLANDESTINA":
                r_num = 7
            elif seg_nome == "TOTAL":
                r_num = 8
            else:
                continue

            ws_seg.cell(r_num, 2).value = int(r["TOTAL FICHAS"])
            ws_seg.cell(r_num, 3).value = int(r["ENCERRADAS MANHA"])
            ws_seg.cell(r_num, 4).value = float(r["% MANHA"])
            ws_seg.cell(r_num, 5).value = int(r["ENCERRADAS A TARDE"])
            ws_seg.cell(r_num, 6).value = float(r["% TARDE"])

        # Salva o arquivo preservando o formato .xlsm
        if not caminho_saida.lower().endswith(".xlsm") and not caminho_saida.lower().endswith(".xlsx"):
            caminho_saida += ".xlsm"

        os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
        wb.save(caminho_saida)
        return caminho_saida

    # Fallback caso o template não seja encontrado
    return _gerar_excel_standalone(resultado, caminho_saida)


def _gerar_excel_standalone(resultado, caminho_saida):
    """Fallback que desenha a planilha do zero caso nenhum template seja localizado."""
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    cor_navy = "0F3252"
    cor_blue_hdr = "0677D9"
    cor_zebra = "F8FAFC"
    cor_borda = "D0D7DE"

    fill_navy = PatternFill(start_color=cor_navy, end_color=cor_navy, fill_type="solid")
    fill_hdr = PatternFill(start_color=cor_blue_hdr, end_color=cor_blue_hdr, fill_type="solid")
    fill_zebra = PatternFill(start_color=cor_zebra, end_color=cor_zebra, fill_type="solid")

    font_title = Font(name="Aptos", size=15, bold=True, color="FFFFFF")
    font_header = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
    font_bold_navy = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
    font_row_bold = Font(name="Aptos", size=10, bold=True, color="0F172A")
    font_row_norm = Font(name="Aptos", size=10, bold=False, color="1E293B")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border_side = Side(border_style="thin", color=cor_borda)
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # ABA PIVOT SEGMENTOS
    ws_pivot = wb.create_sheet(title="Pivot Segmentos")
    ws_pivot.views.sheetView[0].showGridLines = True

    ws_pivot.merge_cells("A1:F2")
    cell_title = ws_pivot["A1"]
    cell_title.value = "FICHAS POR SEGMENTO | ENCERRAMENTO MANHA x TARDE"
    cell_title.font = font_title
    cell_title.fill = fill_navy
    cell_title.alignment = align_center

    headers_resumo = ["SEGMENTO", "TOTAL FICHAS", "ENCERRADAS MANHA", "%", "ENCERRADAS A TARDE", "%"]
    ws_pivot.row_dimensions[4].height = 24

    for col_idx, h_text in enumerate(headers_resumo, 1):
        c = ws_pivot.cell(row=4, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_hdr
        c.alignment = align_center
        c.border = border_cell

    df_resumo = resultado["df_resumo"]
    row_curr = 5

    for _, r in df_resumo.iterrows():
        seg_nome = str(r["SEGMENTO"]).strip()
        is_total = (seg_nome == "TOTAL")

        ws_pivot.row_dimensions[row_curr].height = 20

        c1 = ws_pivot.cell(row=row_curr, column=1, value=seg_nome)
        c2 = ws_pivot.cell(row=row_curr, column=2, value=int(r["TOTAL FICHAS"]))
        c3 = ws_pivot.cell(row=row_curr, column=3, value=int(r["ENCERRADAS MANHA"]))
        c4 = ws_pivot.cell(row=row_curr, column=4, value=float(r["% MANHA"]))
        c5 = ws_pivot.cell(row=row_curr, column=5, value=int(r["ENCERRADAS A TARDE"]))
        c6 = ws_pivot.cell(row=row_curr, column=6, value=float(r["% TARDE"]))

        c1.alignment = align_left
        c2.alignment = align_center
        c3.alignment = align_center
        c4.alignment = align_center
        c5.alignment = align_center
        c6.alignment = align_center

        c2.number_format = "#,##0"
        c3.number_format = "#,##0"
        c4.number_format = "0%"
        c5.number_format = "#,##0"
        c6.number_format = "0%"

        cells = [c1, c2, c3, c4, c5, c6]

        if is_total:
            for cell in cells:
                cell.font = font_bold_navy
                cell.fill = fill_navy
                cell.border = border_cell
        else:
            c1.font = font_row_bold
            for cell in cells[1:]:
                cell.font = font_row_norm
            for cell in cells:
                cell.border = border_cell
                if row_curr % 2 == 1:
                    cell.fill = fill_zebra

        row_curr += 1

    larguras_pivot = {1: 22, 2: 16, 3: 20, 4: 12, 5: 22, 6: 12}
    for col_idx, w in larguras_pivot.items():
        ws_pivot.column_dimensions[get_column_letter(col_idx)].width = w

    # ABA DETALHADA
    ws_det = wb.create_sheet(title="Detalhada")
    ws_det.views.sheetView[0].showGridLines = True

    df_det = resultado["df_detalhada"]
    cols_det = df_det.columns.tolist()

    ws_det.row_dimensions[1].height = 24
    for col_idx, h in enumerate(cols_det, 1):
        c = ws_det.cell(row=1, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_navy
        c.alignment = align_center
        c.border = border_cell

    for r_idx, row_val in enumerate(df_det.itertuples(index=False), 2):
        ws_det.row_dimensions[r_idx].height = 18
        for c_idx, val in enumerate(row_val, 1):
            c = ws_det.cell(row=r_idx, column=c_idx)
            if isinstance(val, datetime):
                c.value = val
                c.number_format = "dd/mm/yyyy hh:mm:ss"
                c.alignment = align_center
            elif isinstance(val, (int, float)) and pd.notna(val):
                c.value = val
                c.alignment = align_right
            else:
                c.value = str(val) if pd.notna(val) else ""
                c.alignment = align_left

            c.font = font_row_norm
            c.border = border_cell
            if r_idx % 2 == 1:
                c.fill = fill_zebra

    for col_idx in range(1, len(cols_det) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws_det.cell(row=r, column=col_idx).value or "")) for r in range(1, min(ws_det.max_row + 1, 50)))
        ws_det.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # ABA DADOS BRUTOS
    ws_bruto = wb.create_sheet(title="DADOS BRUTOS")
    ws_bruto.views.sheetView[0].showGridLines = True
    df_b = resultado["df_bruto"]

    ws_bruto.row_dimensions[1].height = 22
    for col_idx, col_name in enumerate(df_b.columns, 1):
        c = ws_bruto.cell(row=1, column=col_idx, value=str(col_name))
        c.font = font_header
        c.fill = fill_hdr
        c.alignment = align_center
        c.border = border_cell

    for r_idx, row_val in enumerate(df_b.itertuples(index=False), 2):
        for c_idx, val in enumerate(row_val, 1):
            c = ws_bruto.cell(row=r_idx, column=c_idx)
            if isinstance(val, datetime):
                c.value = val
                c.number_format = "dd/mm/yyyy hh:mm:ss"
            elif isinstance(val, (int, float)) and pd.notna(val):
                c.value = val
            else:
                c.value = str(val) if pd.notna(val) else ""
            c.font = font_row_norm
            c.border = border_cell

    os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
    wb.save(caminho_saida)
    return caminho_saida


def gerar_html_clipboard(resultado):
    """Gera tabela estilizada em HTML com CSS inline para colar com perfeição no Outlook / Word."""
    df_resumo = resultado["df_resumo"]
    html = [
        '<div style="font-family: Calibri, Aptos, Segoe UI, sans-serif; margin: 10px 0;">',
        '  <table style="border-collapse: collapse; width: 100%; max-width: 680px; font-size: 13px; border: 1px solid #0F3252; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">',
        '    <thead>',
        '      <tr style="background-color: #0F3252; color: #ffffff; text-align: center; font-weight: bold; height: 36px;">',
        '        <th colspan="6" style="padding: 10px; font-size: 14px; letter-spacing: 0.5px;">FICHAS POR SEGMENTO | ENCERRAMENTO MANHÃ x TARDE</th>',
        '      </tr>',
        '      <tr style="background-color: #0677D9; color: #ffffff; text-align: center; font-size: 12px; height: 30px;">',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252; text-align: left;">SEGMENTO</th>',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252;">TOTAL FICHAS</th>',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252;">ENCERRADAS MANHÃ</th>',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252;">%</th>',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252;">ENCERRADAS À TARDE</th>',
        '        <th style="padding: 8px 12px; border: 1px solid #0F3252;">%</th>',
        '      </tr>',
        '    </thead>',
        '    <tbody>'
    ]

    for idx, r in df_resumo.iterrows():
        seg = str(r["SEGMENTO"]).strip()
        is_total = (seg == "TOTAL")

        if is_total:
            bg = "#0F3252"
            color = "#ffffff"
            weight = "bold"
        else:
            bg = "#F8FAFC" if idx % 2 == 1 else "#ffffff"
            color = "#0F172A"
            weight = "bold" if seg in ("LESTE", "SUL", "CLANDESTINA") else "normal"

        pct_m = f"{r['% MANHA'] * 100:.0f}%"
        pct_t = f"{r['% TARDE'] * 100:.0f}%"

        html.append(f'      <tr style="background-color: {bg}; color: {color}; font-weight: {weight}; height: 28px; text-align: center;">')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1; text-align: left;">{seg}</td>')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1;">{int(r["TOTAL FICHAS"])}</td>')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1;">{int(r["ENCERRADAS MANHA"])}</td>')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1;">{pct_m}</td>')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1;">{int(r["ENCERRADAS A TARDE"])}</td>')
        html.append(f'        <td style="padding: 6px 12px; border: 1px solid #CBD5E1;">{pct_t}</td>')
        html.append('      </tr>')

    html.append('    </tbody>')
    html.append('  </table>')
    html.append('</div>')

    return "\n".join(html)


def gerar_texto_clipboard(resultado):
    """Gera tabela formatada em texto / Markdown para envio no Teams / WhatsApp."""
    df_resumo = resultado["df_resumo"]
    linhas = [
        "📊 *FICHAS POR SEGMENTO | ENCERRAMENTO MANHÃ x TARDE*",
        "",
        "SEGMENTO | TOTAL | MANHÃ | % | TARDE | %",
        "---|---|---|---|---|---"
    ]
    for _, r in df_resumo.iterrows():
        seg = r["SEGMENTO"]
        tot = r["TOTAL FICHAS"]
        m = r["ENCERRADAS MANHA"]
        pm = f"{r['% MANHA'] * 100:.0f}%"
        t = r["ENCERRADAS A TARDE"]
        pt = f"{r['% TARDE'] * 100:.0f}%"
        if seg == "TOTAL":
            linhas.append(f"*{seg}* | *{tot}* | *{m}* | *{pm}* | *{t}* | *{pt}*")
        else:
            linhas.append(f"{seg} | {tot} | {m} | {pm} | {t} | {pt}")

    return "\n".join(linhas)
