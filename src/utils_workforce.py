"""
utils_workforce.py — Serviços utilitários independentes de integração e automação WorkForce.
Camada desacoplada para evitar dependências diretas entre módulos de interface gráfica (AbaFichas, Contagem, Transferência).
"""

import os
import re
import json
import time
import webbrowser
from collections import OrderedDict

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
except ImportError:
    webdriver = None
    ChromeOptions = None
    ChromeService = None

# Lista de referências ativas para impedir que o Garbage Collector do Python feche o navegador Chrome
_ACTIVE_WORKFORCE_DRIVERS = []


def _obter_caminho_config():
    try:
        from utils_os import obter_diretorio_executavel
        base = obter_diretorio_executavel()
    except Exception:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    pasta_data = os.path.join(base, "data")
    os.makedirs(pasta_data, exist_ok=True)
    return os.path.join(pasta_data, "config_fichas.json")


def _carregar_banco_nomes():
    """Carrega mapa de RE -> Nome completo a partir do Banco de nomes.xlsx."""
    mapa = {}
    try:
        from utils_os import obter_diretorio_executavel, obter_caminho_recurso
        caminhos = [
            os.path.join(obter_diretorio_executavel(), "data", "Banco de nomes.xlsx"),
            obter_caminho_recurso("data", "Banco de nomes.xlsx"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "Banco de nomes.xlsx"),
            os.path.join(os.getcwd(), "data", "Banco de nomes.xlsx"),
        ]
        import pandas as pd
        for c in caminhos:
            if os.path.exists(c):
                df = pd.read_excel(c)
                col_re = next((col for col in df.columns if "re" in col.lower()), None)
                col_nome = next((col for col in df.columns if "nome" in col.lower()), None)
                if col_re and col_nome:
                    for _, row in df.iterrows():
                        re_val = str(row[col_re]).strip().split(".")[0]
                        nome_val = str(row[col_nome]).strip()
                        if re_val and nome_val and re_val != "nan" and nome_val != "nan":
                            mapa[re_val] = nome_val
                if mapa:
                    break
    except Exception:
        pass
    return mapa


def _carregar_dicionario_materiais():
    """Tenta carregar o catálogo de materiais para exibição amigável de nomes."""
    catalogo = {}
    try:
        from utils_os import obter_diretorio_executavel, obter_caminho_recurso
        caminhos = [
            os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx"),
            obter_caminho_recurso("data", "banco_de_dados.xlsx"),
            os.path.join(obter_diretorio_executavel(), "data", "Banco de nomes.xlsx"),
            obter_caminho_recurso("data", "Banco de nomes.xlsx"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "banco_de_dados.xlsx"),
            os.path.join(os.getcwd(), "data", "banco_de_dados.xlsx"),
        ]
        import pandas as pd
        for c in caminhos:
            if os.path.exists(c):
                df = pd.read_excel(c)
                col_mat = next((col for col in df.columns if "material" in col.lower() or "codigo" in col.lower()), None)
                col_desc = next((col for col in df.columns if "descri" in col.lower() or "nome" in col.lower()), None)
                if col_mat and col_desc:
                    for _, row in df.iterrows():
                        m = str(row[col_mat]).strip().split(".")[0]
                        d = str(row[col_desc]).strip()
                        if m and d and m != "nan" and d != "nan":
                            catalogo[m] = d
                if catalogo:
                    break
    except Exception:
        pass
    return catalogo


def carregar_credenciais_workforce():
    caminho = _obter_caminho_config()
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"usuario": "19595", "senha": "979648"}


def _selecionar_dropdown_elemento(driver, select_element, re_codigo, mapa_nomes=None):
    """Seleciona a opção do dropdown combinando o RE ou o nome do Banco de nomes."""
    re_str = str(re_codigo).strip()
    nome_alvo = (mapa_nomes.get(re_str, "") if mapa_nomes else "").lower().strip()
    partes_nome = [p for p in nome_alvo.split() if len(p) > 2]
    for _ in range(3):
        try:
            s = Select(select_element)
            for opt in s.options:
                t = opt.text.lower().strip()
                v = opt.get_attribute("value") or ""
                if re_str and (re_str in t or re_str == v):
                    s.select_by_visible_text(opt.text)
                    return True
                if nome_alvo and nome_alvo in t:
                    s.select_by_visible_text(opt.text)
                    return True
                if len(partes_nome) >= 2 and partes_nome[0] in t and partes_nome[-1] in t:
                    s.select_by_visible_text(opt.text)
                    return True
        except Exception:
            time.sleep(0.5)
    return False


def executar_fluxo_workforce(dados, callback_status=None, login_usuario=None, login_senha=None, mapa_nomes=None):
    """
    Automação robusta com Selenium para preencher Ficha de Solicitação & Materiais no WorkForce.
    Pode ser chamada de qualquer módulo (AbaFichas, Transferência entre Equipes, Contagem de Carro).
    Retorna (sucesso: bool, driver_ou_mensagem).
    """
    def _status(msg):
        if callback_status:
            try:
                callback_status(msg)
            except Exception:
                pass

    if not login_usuario or not login_senha:
        cfg = carregar_credenciais_workforce()
        login_usuario = login_usuario or cfg.get("usuario", "19595")
        login_senha = login_senha or cfg.get("senha", "979648")

    if mapa_nomes is None:
        mapa_nomes = _carregar_banco_nomes()

    url_nova_ficha = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFichaDetail.aspx?MedicaoFichaId=0"
    url_login = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFicha.aspx"

    if webdriver is None:
        webbrowser.open(url_nova_ficha, new=2)
        _status("⚠️ Selenium não instalado; formulário aberto no navegador padrão.")
        return False, "Selenium não instalado"

    driver = None
    try:
        if ChromeOptions is not None:
            options = ChromeOptions()
        else:
            options = webdriver.ChromeOptions()

        options.add_argument("--start-maximized")
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")

        _status("Iniciando Chrome...")
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 25)

        # 1. Tela de Login
        _status("Acessando autenticação do WorkForce...")
        driver.get(url_login)

        usuario_el = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id*='wtUserNameInput'], input[type='text']"))
        )
        senha_el = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id*='wtPasswordInput'], input[type='password']"))
        )

        usuario_el.clear()
        usuario_el.send_keys(str(login_usuario))
        senha_el.clear()
        senha_el.send_keys(str(login_senha))

        btn_login = driver.find_element(
            By.CSS_SELECTOR, "input[id*='wt16'], input[type='submit'][value*='Entrar'], input[type='submit']"
        )

        _status(f"Logando como {login_usuario}...")
        btn_login.click()

        wait.until(EC.staleness_of(btn_login))
        time.sleep(1.5)

        # 2. Navegar para a Nova Ficha
        _status("Carregando formulário de Nova Ficha...")
        driver.get(url_nova_ficha)

        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idBase')]")
            )
        )
        time.sleep(1.0)

        # 3. Base (deve ser selecionada primeiro pois o WorkForce recarrega campos via Ajax)
        base_alvo = dados.get("base", "Mauá").strip()
        if base_alvo:
            _status(f"Configurando Base: {base_alvo}...")
            try:
                sel_base_el = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idBase')]")
                s = Select(sel_base_el)
                alvo = base_alvo.lower()
                alterou_base = False
                if alvo not in s.first_selected_option.text.lower():
                    for opt in s.options:
                        opt_txt = opt.text.lower()
                        if alvo in opt_txt or ("santo andr" in opt_txt and "santo andre" in alvo):
                            s.select_by_visible_text(opt.text)
                            alterou_base = True
                            break
                if alterou_base:
                    # Aguarda o recarregamento Ajax dos funcionários da Base
                    time.sleep(3.0)
            except Exception:
                pass

        # 4. Preenchimento de Data e Responsáveis
        _status("Preenchendo cabeçalho e responsáveis...")

        # Data obrigatória (DD-MM-YYYY)
        data_val = dados.get("data", time.strftime("%d-%m-%Y"))
        if data_val:
            try:
                el = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_data')]")
                el.clear()
                el.send_keys(str(data_val).replace("/", "-"))
            except Exception:
                pass

        # Resp Almox (Input + Select)
        re_almox = dados.get("re_almox") or dados.get("re_almox1") or "19595"
        if re_almox:
            try:
                inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reRespAlmox') and not(contains(@id, 'reRespAlmox2'))]")
                inp.clear()
                inp.send_keys(str(re_almox).strip())
                sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idRespAlmox') and not(contains(@id, 'idRespAlmox2'))]")
                _selecionar_dropdown_elemento(driver, sel, re_almox, mapa_nomes)
            except Exception:
                pass

        # Resp Almox Entregue (Input + Select)
        re_almox2 = dados.get("re_almox_entregue") or dados.get("re_almox2")
        if re_almox2:
            try:
                inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reRespAlmox2')]")
                inp.clear()
                inp.send_keys(str(re_almox2).strip())
                sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idRespAlmox2')]")
                _selecionar_dropdown_elemento(driver, sel, re_almox2, mapa_nomes)
            except Exception:
                pass

        # Eletricista 1 (Input + Select)
        re_elet1 = dados.get("re_eletricista_1") or dados.get("re_elet1")
        if re_elet1:
            try:
                inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reEletricista1')]")
                inp.clear()
                inp.send_keys(str(re_elet1).strip())
                sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idEletricista1')]")
                _selecionar_dropdown_elemento(driver, sel, re_elet1, mapa_nomes)
            except Exception:
                pass

        # Eletricista 2 (Input + Select)
        re_elet2 = dados.get("re_eletricista_2") or dados.get("re_elet2")
        if re_elet2:
            try:
                inp = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_reEletricista2')]")
                inp.clear()
                inp.send_keys(str(re_elet2).strip())
                sel = driver.find_element(By.XPATH, "//select[contains(@id, 'almoxarifado_materialMedicaoFicha_idEletricista2')]")
                _selecionar_dropdown_elemento(driver, sel, re_elet2, mapa_nomes)
            except Exception:
                pass

        # Recurso
        recurso = dados.get("recurso")
        if recurso:
            try:
                el = driver.find_element(By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_recurso')]")
                el.clear()
                el.send_keys(str(recurso).strip())
            except Exception:
                pass

        # Observação
        observacao = dados.get("observacao")
        if observacao:
            try:
                obs_els = driver.find_elements(By.XPATH, "//textarea[not(contains(@id, 'ModalConsulta'))]")
                if obs_els:
                    obs_els[0].clear()
                    obs_els[0].send_keys(str(observacao))
            except Exception:
                pass

        # 5. Preenchimento dos Materiais na tabela (Pedido e Fornecido)
        materiais = dados.get("materiais", [])
        if materiais:
            _status(f"Preenchendo {len(materiais)} material(is) na tabela...")
            time.sleep(1.0)

            # Normaliza materiais para lista de [codigo, quantidade, codigo_limpo, aliases]
            def _obter_aliases_material(codigo_str):
                c_up = str(codigo_str).strip().upper()
                c_clean = c_up.lstrip("0") or "0"
                aliases = [c_up, c_clean]
                if "510388" in c_up or c_clean == "510388":
                    aliases.extend(["SMART METER BIFÁSICO", "SMART METER BIFASICO", "SMART METER", "SMART"])
                elif "510128" in c_up or c_clean == "510128":
                    aliases.extend(["DELTA SMART", "DELTA"])
                elif "510136" in c_up or c_clean == "510136" or "510137" in c_up:
                    aliases.extend(["SMART"])
                elif c_up.startswith("D23"):
                    aliases.extend(["ESTRELA LUZ BIFÁSICO (NOVO)", "ESTRELA LUZ", "ESTRELA LUZ TB", "BIFÁSICO", "BIFASICO"])
                elif c_up.startswith("D34"):
                    aliases.extend(["ESTRELA FORÇA TRIFÁSICO (NOVO)", "ESTRELA FORÇA", "ESTRELA FORCA", "ESTRELA FORÇA TB", "TRIFÁSICO", "TRIFASICO"])
                elif c_up.startswith("D12") or c_up.startswith("D13"):
                    aliases.extend(["MONOFÁSICO", "MONOFASICO", "BICONCÊNTRICO", "BICONCENTRICO"])
                elif c_up.startswith("C0006"):
                    aliases.extend(["CENTRALIZADO DIRETO", "CENTRALIZADO"])
                elif c_up.startswith("A34"):
                    aliases.extend(["TRIFASICO INDIRETO", "INDIRETO"])
                elif "330702" in c_up:
                    aliases.extend(["RAMAL 10 MM", "RAMAL 10 MM TRIPLEX"])
                elif "130097" in c_up:
                    aliases.extend(["CHAVE, AFERICAO", "CHAVE DE AFERICAO"])
                return list(OrderedDict.fromkeys(aliases)) if 'OrderedDict' in globals() else list(dict.fromkeys(aliases))

            materiais_norm = []
            for item in materiais:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    c = str(item[0]).strip()
                    q = str(item[1]).strip()
                elif isinstance(item, dict):
                    c = str(item.get("codigo", "")).strip()
                    q = str(item.get("quantidade", 1)).strip()
                else:
                    continue
                if not c:
                    continue
                c_limpo = c.lstrip("0") if c.isdigit() else c
                als = _obter_aliases_material(c)
                materiais_norm.append([c, q, c_limpo, als])

            js_preencher = """
                var materiais = arguments[0];
                var table = document.querySelector("table[id*='AlmoxarifadoMaterialMedicaoTable']");
                if (!table) return {sucesso: false, erro: "Tabela não encontrada"};

                var rows = table.querySelectorAll("tbody tr, tr");
                var preenchidos = [];
                var nao_encontrados = [];

                for (var m = 0; m < materiais.length; m++) {
                    var codAlvo = String(materiais[m][0]).trim().toUpperCase();
                    var qtdAlvo = String(materiais[m][1]).trim();
                    var codClean = String(materiais[m][2] || '').trim().toUpperCase();
                    var aliases = materiais[m][3] || [];
                    var achou = false;

                    for (var i = 0; i < rows.length; i++) {
                        var r = rows[i];
                        var cells = r.querySelectorAll("td");
                        if (cells.length >= 6) {
                            var txtCod = cells[1].textContent.trim().toUpperCase();
                            
                            var match = false;
                            if (txtCod === codAlvo || (codAlvo.length >= 3 && txtCod.indexOf(codAlvo) !== -1)) {
                                match = true;
                            } else if (codClean && (txtCod === codClean || (codClean.length >= 3 && txtCod.indexOf(codClean) !== -1))) {
                                match = true;
                            } else if (codAlvo && txtCod.length >= 3 && codAlvo.indexOf(txtCod) !== -1) {
                                match = true;
                            } else {
                                for (var a = 0; a < aliases.length; a++) {
                                    var al = String(aliases[a]).trim().toUpperCase();
                                    if (al && al.length >= 3 && (txtCod === al || txtCod.indexOf(al) !== -1)) {
                                        match = true;
                                        break;
                                    }
                                }
                            }

                            if (match) {
                                // Cell 4: Pedido, Cell 5: Fornecido
                                var inps = r.querySelectorAll("input[type='number']");
                                var inpPed = (cells.length > 4 ? cells[4].querySelector("input") : null) || inps[0];
                                var inpForn = (cells.length > 5 ? cells[5].querySelector("input") : null) || inps[1];

                                if (inpPed) {
                                    inpPed.focus();
                                    inpPed.value = qtdAlvo;
                                    inpPed.dispatchEvent(new Event('input', {bubbles: true}));
                                    inpPed.dispatchEvent(new Event('change', {bubbles: true}));
                                    inpPed.blur();
                                }
                                if (inpForn) {
                                    inpForn.focus();
                                    inpForn.value = qtdAlvo;
                                    inpForn.dispatchEvent(new Event('input', {bubbles: true}));
                                    inpForn.dispatchEvent(new Event('change', {bubbles: true}));
                                    inpForn.blur();
                                }

                                r.scrollIntoView({behavior: 'smooth', block: 'center'});
                                achou = true;
                                var rotulo = codClean || codAlvo;
                                preenchidos.push(rotulo + " (" + qtdAlvo + " un)");
                                break;
                            }
                        }
                    }
                    if (!achou) {
                        nao_encontrados.push(codClean || codAlvo);
                    }
                }
                return {sucesso: true, preenchidos: preenchidos, nao_encontrados: nao_encontrados};
            """

            try:
                res = driver.execute_script(js_preencher, materiais_norm)
                if res and res.get("sucesso"):
                    p_list = res.get("preenchidos", [])
                    n_list = res.get("nao_encontrados", [])
                    if n_list:
                        _status(f"Materiais preenchidos: {len(p_list)}. Não encontrados no site: {', '.join(n_list)}")
                    else:
                        _status(f"✅ {len(p_list)} material(is) preenchido(s) com sucesso na tabela!")
                else:
                    _status(f"⚠️ Atenção ao preencher materiais: {res.get('erro', 'erro desconhecido')}")
            except Exception as e_js:
                _status(f"⚠️ Falha ao preencher tabela via script: {e_js}")

            time.sleep(1.0)

        # 6. Salvar automaticamente ou deixar para conferência
        if dados.get("salvar_auto", False):
            _status("Gravando e salvando Ficha no WorkForce...")
            try:
                btn_salvar = None
                seletores = [
                    "//input[@value='Salvar']",
                    "//input[contains(@id, 'btnSalvar')]",
                    "//input[contains(@value, 'Salvar')]",
                    "//button[contains(., 'Salvar')]",
                    "//a[contains(., 'Salvar')]"
                ]
                for sel_xpath in seletores:
                    try:
                        els = driver.find_elements(By.XPATH, sel_xpath)
                        for el in els:
                            if el.is_displayed():
                                btn_salvar = el
                                break
                        if btn_salvar:
                            break
                    except Exception:
                        pass

                if btn_salvar:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
                    time.sleep(0.5)
                    try:
                        btn_salvar.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn_salvar)

                    # Aceita diálogo modal de confirmação do navegador (window.confirm / alert) se houver
                    try:
                        WebDriverWait(driver, 3).until(EC.alert_is_present())
                        driver.switch_to.alert.accept()
                    except Exception:
                        pass

                    # Aguarda redirecionamento ou mensagem de sucesso
                    time.sleep(3.5)
                    if "MedicaoFichaId=0" not in driver.current_url:
                        _status("✅ Ficha criada e salva com sucesso no WorkForce!")
                    else:
                        _status("✅ Ficha gravada no WorkForce!")
                else:
                    _status("⚠️ Ficha preenchida! Clique em Salvar no navegador.")
            except Exception as e_save:
                _status(f"⚠️ Ficha preenchida! Clique em Salvar no navegador ({e_save}).")
        else:
            _status("✅ Ficha e materiais preenchidos com sucesso! Confira e clique em Salvar no navegador.")

        # Mantém a referência do driver viva para impedir que o Chrome feche ao fim do thread
        _ACTIVE_WORKFORCE_DRIVERS.append(driver)
        return True, driver

    except Exception as erro:
        _status(f"⚠️ Erro durante o fluxo WorkForce: {erro}")
        return False, str(erro)


def _aguardar_estabilidade_ajax(driver, timeout=20):
    """Aguarda o término de requisições Ajax do OutSystems/jQuery e estabilidade do DOM."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script(
                "return (typeof OsAjaxActiveRequests === 'undefined' || OsAjaxActiveRequests === 0) && "
                "(typeof jQuery === 'undefined' || jQuery.active === 0);"
            )
        )
    except Exception:
        pass
    time.sleep(1.0)


def preencher_ficha_saida_workforce(
    equipe,
    doc_sap,
    materiais_agrupados,
    callback_status=None,
    login_usuario=None,
    login_senha=None,
    salvar_auto=False
):
    """
    Localiza a Ficha de Solicitação existente do recurso/equipe no WorkForce,
    preenche o campo 'Doc.' com o Documento SAP MIGO gerado,
    abre o modal de seriais para cada material correspondente inserindo os seriais separados por ';',
    atualiza a quantidade fornecida na tabela de materiais, e opcionalmente salva a ficha.
    
    Retorna (sucesso: bool, mensagem_ou_driver: Any).
    """
    def _status(msg):
        if callback_status:
            try:
                callback_status(msg)
            except Exception:
                pass

    if not login_usuario or not login_senha:
        cfg = carregar_credenciais_workforce()
        login_usuario = login_usuario or cfg.get("usuario", "19595")
        login_senha = login_senha or cfg.get("senha", "979648")

    url_lista = "https://monitoria.conectae.com.br/workforce/MaterialMedicaoFicha.aspx"

    if webdriver is None:
        webbrowser.open(url_lista, new=2)
        _status("⚠️ Selenium não instalado; página aberta no navegador padrão.")
        return False, "Selenium não instalado"

    driver = None
    try:
        if ChromeOptions is not None:
            options = ChromeOptions()
        else:
            options = webdriver.ChromeOptions()

        options.add_argument("--start-maximized")
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")

        _status("🌐 Iniciando navegador Chrome...")
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 25)

        # 1. Login no WorkForce
        _status("🔑 Acessando WorkForce...")
        driver.get(url_lista)

        try:
            # Verifica se está na tela de login
            usuario_el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[id*='wtUserNameInput'], input[type='text']"))
            )
            senha_el = driver.find_element(By.CSS_SELECTOR, "input[id*='wtPasswordInput'], input[type='password']")
            usuario_el.clear()
            usuario_el.send_keys(str(login_usuario))
            senha_el.clear()
            senha_el.send_keys(str(login_senha))

            btn_login = driver.find_element(
                By.CSS_SELECTOR, "input[id*='wt16'], input[type='submit'][value*='Entrar'], input[type='submit']"
            )
            _status(f"🔑 Autenticando como {login_usuario}...")
            btn_login.click()
            wait.until(EC.staleness_of(btn_login))
            time.sleep(2.0)
        except Exception:
            # Já pode estar logado
            pass

        # Garante que está na lista de fichas
        if "MaterialMedicaoFicha.aspx" not in driver.current_url:
            driver.get(url_lista)
            time.sleep(2.0)

        # 2. Localizar a Ficha da Equipe / Recurso
        recurso_alvo = str(equipe).strip().upper()
        _status(f"🔍 Buscando ficha da equipe {recurso_alvo} no WorkForce...")

        try:
            inp_rec = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_recurso')]")
                )
            )
            inp_rec.clear()
            inp_rec.send_keys(recurso_alvo)

            btn_buscar = driver.find_element(
                By.XPATH, "//input[@value='Buscar'] | //input[contains(@id, 'wt35')]"
            )
            btn_buscar.click()
            _aguardar_estabilidade_ajax(driver, timeout=15)
        except Exception as e_busca:
            _status(f"⚠️ Atenção ao filtrar recurso: {e_busca}. Procurando na lista geral...")

        # Procura a linha da equipe na tabela de fichas
        link_ficha = None
        for tentativa in range(3):
            try:
                links = driver.find_elements(By.XPATH, "//table//a[contains(@href, 'MaterialMedicaoFichaDetail.aspx')]")
                for a in links:
                    t_a = a.text.strip().upper()
                    if recurso_alvo in t_a or t_a in recurso_alvo:
                        link_ficha = a
                        break
                if link_ficha:
                    break
                time.sleep(1.0)
            except Exception:
                time.sleep(1.0)

        if not link_ficha:
            msg_erro = f"Ficha da equipe '{recurso_alvo}' não encontrada na lista do WorkForce na data de hoje."
            _status(f"⚠️ {msg_erro}")
            _ACTIVE_WORKFORCE_DRIVERS.append(driver)
            return False, msg_erro

        _status(f"📋 Abrindo ficha da equipe {recurso_alvo}...")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_ficha)
        time.sleep(0.5)
        try:
            link_ficha.click()
        except Exception:
            href = link_ficha.get_attribute("href")
            driver.get(href)

        # 3. Na tela de Detalhe da Ficha (MaterialMedicaoFichaDetail.aspx)
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[contains(@id, 'AlmoxarifadoMaterialMedicaoTable')]")
            )
        )
        _aguardar_estabilidade_ajax(driver, timeout=15)

        # Preenche o campo Doc.
        if doc_sap and str(doc_sap).strip() and str(doc_sap).strip() != "REGISTRADO":
            _status(f"📝 Preenchendo campo Doc.: {doc_sap}...")
            try:
                inp_doc = driver.find_element(
                    By.XPATH, "//input[contains(@id, 'almoxarifado_materialMedicaoFicha_doc')]"
                )
                driver.execute_script("""
                    var el = arguments[0];
                    el.focus();
                    el.value = arguments[1];
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    el.blur();
                """, inp_doc, str(doc_sap).strip())
                time.sleep(0.5)
            except Exception as e_doc:
                _status(f"⚠️ Campo Doc. não encontrado ou erro ao preencher: {e_doc}")

        # 4. Lança os Materiais e Seriais na tabela
        _status(f"📦 Verificando {len(materiais_agrupados)} material(is) a fornecer...")
        materiais_processados = 0

        for cod_mat, lista_itens in materiais_agrupados.items():
            cod_str = str(cod_mat).strip().upper()
            if not cod_str or cod_str == "—":
                continue

            seriais = [
                str(it.get("serial", "")).strip()
                for it in lista_itens
                if it.get("serial") and not it.get("eh_nao_serializado")
            ]
            
            # Quantidade total a fornecer
            if seriais:
                qtd_total = len(seriais)
            else:
                qtd_total = sum(float(it.get("quantidade_num", 0)) for it in lista_itens if it.get("eh_nao_serializado"))

            # Localiza a linha correspondente na tabela
            table_el = driver.find_element(By.XPATH, "//table[contains(@id, 'AlmoxarifadoMaterialMedicaoTable')]")
            rows = table_el.find_elements(By.XPATH, ".//tbody/tr | .//tr")
            linha_alvo = None

            cod_clean = cod_str.lstrip("0") or "0"
            for r in rows:
                cells = r.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 6:
                    txt_cod = cells[1].text.strip().upper()
                    txt_desc = cells[3].text.strip().upper() if len(cells) > 3 else ""
                    if cod_str == txt_cod or cod_clean == txt_cod or (len(cod_str) >= 3 and cod_str in txt_cod):
                        linha_alvo = r
                        break
                    elif cod_str in txt_desc:
                        linha_alvo = r
                        break

            if not linha_alvo:
                _status(f"⚠️ Material {cod_str} não encontrado na grade da Ficha.")
                continue

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha_alvo)
            time.sleep(0.5)

            # Se houver seriais, abre o ModalConsulta e insere
            if seriais:
                _status(f"🔍 Abrindo pop-up de seriais para {cod_str} ({len(seriais)} unid)...")
                # Procura o link de clique na célula 1 ou 0
                link_modal = None
                links_linha = linha_alvo.find_elements(By.TAG_NAME, "a")
                if links_linha:
                    link_modal = links_linha[0]
                else:
                    cells = linha_alvo.find_elements(By.TAG_NAME, "td")
                    if len(cells) > 1:
                        link_modal = cells[1]

                if link_modal:
                    driver.execute_script("arguments[0].click();", link_modal)
                    
                    # Aguarda o modal de seriais abrir
                    try:
                        txt_area_modal = WebDriverWait(driver, 15).until(
                            EC.visibility_of_element_located(
                                (By.XPATH, "//*[contains(@id, 'ModalConsulta')]//textarea")
                            )
                        )
                        _aguardar_estabilidade_ajax(driver, timeout=5)
                        
                        # Formata seriais com ponto e vírgula no final conforme WorkForce exige
                        texto_seriais = ";".join(seriais) + ";"
                        _status(f"📋 Inserindo {len(seriais)} seriais no WorkForce...")
                        
                        driver.execute_script("""
                            var el = arguments[0];
                            el.focus();
                            el.value = arguments[1];
                            el.dispatchEvent(new Event('input', {bubbles: true}));
                            el.dispatchEvent(new Event('change', {bubbles: true}));
                        """, txt_area_modal, texto_seriais)
                        time.sleep(0.5)

                        # Clica em 'Incluir Serial'
                        btn_incluir = driver.find_element(
                            By.XPATH, "//*[contains(@id, 'ModalConsulta')]//input[@value='Incluir Serial']"
                        )
                        driver.execute_script("arguments[0].click();", btn_incluir)
                        _status("⏳ Gravando seriais no modal...")
                        _aguardar_estabilidade_ajax(driver, timeout=25)
                        time.sleep(1.5)

                        # Fecha o modal no ícone 'x'
                        btn_fechar = None
                        fechar_els = driver.find_elements(
                            By.XPATH,
                            "//*[contains(@id, 'ModalConsulta')]//a[contains(@id, 'Icon')] | "
                            "//*[contains(@id, 'ModalConsulta')]//div[contains(@class, 'modal-close')]//a"
                        )
                        if fechar_els:
                            btn_fechar = fechar_els[0]

                        if btn_fechar:
                            driver.execute_script("arguments[0].click();", btn_fechar)
                            try:
                                WebDriverWait(driver, 8).until(
                                    EC.invisibility_of_element_located(
                                        (By.XPATH, "//*[contains(@id, 'ModalConsulta')]//textarea")
                                    )
                                )
                            except Exception:
                                pass
                        else:
                            # Fallback: pressiona ESC
                            try:
                                from selenium.webdriver.common.keys import Keys
                                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                            except Exception:
                                pass

                        _aguardar_estabilidade_ajax(driver, timeout=10)
                        _status(f"✓ Seriais do material {cod_str} gravados com sucesso!")
                    except Exception as e_modal:
                        _status(f"⚠️ Erro ao lançar seriais no modal para {cod_str}: {e_modal}")

            # Preenche a quantidade Fornecido (Cell 5)
            try:
                cells = linha_alvo.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 6:
                    inp_forn = cells[5].find_element(By.TAG_NAME, "input")
                    qtd_str = str(int(qtd_total) if qtd_total == int(qtd_total) else qtd_total)
                    driver.execute_script("""
                        var el = arguments[0];
                        el.focus();
                        el.value = arguments[1];
                        el.dispatchEvent(new Event('input', {bubbles: true}));
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                        el.blur();
                    """, inp_forn, qtd_str)
            except Exception:
                pass

            materiais_processados += 1

        # 5. Salvar automaticamente ou deixar na tela
        if salvar_auto:
            _status("💾 Salvando Ficha no WorkForce...")
            try:
                btn_salvar = driver.find_element(
                    By.XPATH, "//input[@value='Salvar'] | //input[contains(@id, 'btnGravar')]"
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", btn_salvar)

                # Aceita confirmação do navegador se houver
                try:
                    WebDriverWait(driver, 4).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                except Exception:
                    pass

                _aguardar_estabilidade_ajax(driver, timeout=20)
                time.sleep(2.0)
                _status("✅ Ficha gravada e salva com sucesso no WorkForce!")
            except Exception as e_save:
                _status(f"⚠️ Atenção ao salvar: {e_save}. Confira na tela do Chrome e clique em Salvar.")
        else:
            _status("✅ Doc e seriais preenchidos com sucesso! Confira na janela do Chrome e clique em Salvar.")

        _ACTIVE_WORKFORCE_DRIVERS.append(driver)
        return True, driver

    except Exception as erro:
        _status(f"⚠️ Erro durante o fluxo WorkForce: {erro}")
        return False, str(erro)

