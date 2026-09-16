# Conecta Hub 3.0


Python/Tkinter desktop app — SAP GUI automation hub.

## Entrypoint

- `conecta_hub2.py` — main app, creates the Tk root and notebook with 6 tabs.

## Modules

All under the repo root, no package structure. Each exports a `construir_aba_*` factory.

| File | Tab | SAP needed |
|---|---|---|
| `modulo_etiquetas.py` | Gerador de Etiquetas | No |
| `modulo_extrator.py` | Extrator de Seriais / IQ09 | Yes (IQ09) |
| `modulo_Sap.py` | Login SAP | Yes |
| `modulo_bandeirada.py` | Bandeirada | Yes |
| `modulo_extrair_OSME.py` | Extrator OSME | Yes |
| `modulo_conferencia.py` | Conferência | No |
| `modulo_chat.py` | Chatbot (popup) | No |
| `modulo_consultar_migo.py` | Consultar MIGO (Clandestinas RCL001..RCL008) | Yes (IW39 / MIGO) |

## Build

```powershell
py -m PyInstaller "Conecta Hub 2.0.spec" --clean --noconfirm
```

Output: `dist\Conecta Hub 2.0\Conecta Hub 2.0.exe`

### Spec file gotchas

- `excludes` list has caused bugs — **never add `numpy`, `PIL`, or `pandas` to it.
- Hidden imports required: `numpy`, `PIL`, `pandas`, `openpyxl`, `reportlab.*`, `qrcode`, `win32com`, `pythoncom`, `tkinter.*`.
- Data files to bundle: all `.vbs`, `.ico`, `.png`, `.xlsx` in the repo root.

### Resource resolution

- Bundled resources (icons, VBS scripts, logo) → `sys._MEIPASS` via `obter_caminho_recurso()`.
- User-editable files (Excel database, config JSONs) → exe directory (`os.path.dirname(sys.executable)`), with fallback to `_MEIPASS`.

## Dependencies

`requirements.txt` with all deps. Install: `py -m pip install -r requirements.txt`

## SAP dependency

OSME, Bandeirada, and IQ09 tabs check `verificar_sap_aberto()` (COM `GetObject("SAPGUI")`) before running. If SAP GUI is closed, a messagebox error is shown and the operation is blocked.

## Python

CPython 3.14, Windows-only (COM + cscript + SAP GUI).

---

# Chatbot — Guia de Funções

O chatbot (`modulo_chat.py`) usa **OpenRouter** para processar linguagem natural e chamar funções do programa.
API Key salva em `config_chat.txt`.

## Modelos disponíveis (OpenRouter)

| Modelo | Gratuito | Nota |
|---|---|---|
| DeepSeek V3 (Free) | Sim | Padrão, ótimo custo-benefício |
| Llama 3.3 70B (Free) | Sim | Bom para raciocínio |
| Gemma 3 27B (Free) | Sim | Multilíngue |
| Nemotron 70B (Free) | Sim | Longo contexto |
| MiMo V2.5 | Não (pago) | Melhor raciocínio geral |

## Funções disponíveis

### buscar_lote(lote)
Busca informacoes de um lote na planilha de controle de seriais.
- **Parâmetros:** `lote` (string) — ex: "CL03-002"
- **Retorna:** deposito, material, quantidade, seriais
- **Efeito colateral:** preenche os campos na aba Gerador de Etiquetas automaticamente

### gerar_doc_google(lote)
Cria um documento no Google Docs com os seriais do lote e retorna o link.
- **Parâmetros:** `lote` (string) — ex: "CL03-002"
- **Efeito:** busca o lote, gera o doc, preenche o campo QR Topo com o link

### gerar_etiqueta(deposito, material, descricao, qtd, unidade, lote)
Gera uma etiqueta em PDF para impressao.
- **Parâmetros:** `deposito` (obrigatório), demais opcionais
- **Retorna:** caminho do PDF gerado

### consultar_material(codigo)
Busca a descricao de um material pelo codigo no banco de dados local.
- **Parâmetros:** `codigo` (string) — ex: "18398041"
- **Retorna:** codigo e descricao do material

### extrair_seriais(texto)
Extrai numeros de serie de um texto, suportando intervalos (ex: "18398041 a 18398050").
- **Parâmetros:** `texto` (string) — texto com seriais ou intervalos
- **Efeito colateral:** cola o texto na aba Extrator e já extrai, troca para aba automaticamente

### copiar_seriais()
Copia os seriais extraidos para a area de transferencia.
- **Parâmetros:** nenhum

### listar_lotes()
Lista todos os lotes disponiveis na planilha de controle.
- **Parâmetros:** nenhum

### login_sap(usuario)
Faz login no SAP usando credenciais salvas.
- **Parâmetros:** `usuario` (string) — "Anderson", "Junior", "Pedro", "Adriana"

### info_sap()
Retorna informacoes do SAP atual (sistema, cliente, usuario, transacao).

## Fluxos típicos

1. **Buscar lote e gerar etiqueta:** buscar_lote → gerar_etiqueta
2. **Gerar documento no Drive:** buscar_lote → gerar_doc_google
3. **Extrair e copiar seriais:** extrair_seriais → copiar_seriais
4. **Logar no SAP:** login_sap com nome do usuario

## Segurança

- `logins_sap.json` e `config_chat.txt` estão no `.gitignore`
- Nunca commite senhas ou chaves de API
