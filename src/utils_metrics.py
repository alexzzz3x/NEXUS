import os
import json
from utils_os import obter_diretorio_executavel

ARQUIVO_METRICAS = os.path.join(obter_diretorio_executavel(), "data", "metrics.json")

def _ler_metricas():
    if os.path.exists(ARQUIVO_METRICAS):
        try:
            with open(ARQUIVO_METRICAS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def ler_metricas():
    return _ler_metricas()

def incrementar_metrica(chave, valor=1):
    metricas = _ler_metricas()
    metricas[chave] = metricas.get(chave, 0) + valor
    try:
        os.makedirs(os.path.dirname(ARQUIVO_METRICAS), exist_ok=True)
        with open(ARQUIVO_METRICAS, "w", encoding="utf-8") as f:
            json.dump(metricas, f, ensure_ascii=False, indent=4)
    except Exception as e:
        try:
            from utils_logger import log_error
            log_error(f"Erro ao salvar métricas: {e}")
        except:
            pass
