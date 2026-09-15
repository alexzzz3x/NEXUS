import os
import logging
from datetime import datetime
from utils_os import obter_diretorio_executavel

LOG_FILE = os.path.join(obter_diretorio_executavel(), "data", "app_diagnostic.log")

def _setup_logger():
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    logger = logging.getLogger("ConectaHub")
    logger.setLevel(logging.DEBUG)
    
    if not logger.handlers:
        fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        
        logger.addHandler(fh)
    return logger

logger = _setup_logger()

def log_info(msg):
    logger.info(msg)

def log_warning(msg):
    logger.warning(msg)

def log_error(msg, exc_info=False):
    logger.error(msg, exc_info=exc_info)

def exportar_logs(pasta_destino):
    import shutil
    try:
        shutil.copy2(LOG_FILE, os.path.join(pasta_destino, f"ConectaHub_Diagnostic_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"))
        return True
    except Exception as e:
        log_error(f"Erro ao exportar logs: {e}")
        return False
