# N.E.X.U.S

Python/CustomTkinter desktop app — SAP GUI automation hub.

---

## 🛡️ PROTOCOLO DE BLINDAGEM E ISOLAMENTO MODULAR (ESTRITAMENTE OBRIGATÓRIO)

Para assegurar a integridade do sistema e evitar regressões acidentais, **todo agente ou desenvolvedor deve seguir rigorosamente as regras abaixo**:

1. **Princípio do Escopo Estrito (Zero Efeito Colateral)**:
   - Ao receber uma tarefa direcionada a um módulo específico (ex: `modulo_saida_osme.py`, `modulo_contagem_carro.py`, etc.), é **terminantemente proibido modificar, renomear, mover ou refatorar qualquer outro módulo** (`modulo_*.py`) sem solicitação expressa do usuário.
   - Jamais realize limpezas "em massa" ou edições cosméticas em arquivos alheios ao escopo da solicitação atual.

2. **Desacoplamento de Componentes (Regra de Dependência Neutra)**:
   - Nenhum módulo de interface (`modulo_*.py`) deve depender diretamente de funções de outro módulo de interface.
   - Qualquer lógica compartilhada (automações Selenium, leituras de Excel, parsers de texto, credenciais) deve residir em módulos utilitários neutros (`utils_*.py`).

3. **Tolerância a Falhas no Hub (`conecta_hub2.py`)**:
   - Cada módulo deve ser importado e inicializado em bloco isolado de tratamento de erro (`try...except`).
   - Se um módulo falhar, for excluído ou estiver em manutenção, **todos os outros módulos devem continuar funcionando normalmente**.

4. **Isolamento de Sessão SAP**:
   - Automações SAP devem sempre iniciar e finalizar no menu principal (`/n` com Enter).
   - Um erro de transação em um módulo jamais deve deixar caixas de diálogo modais abertas que possam travar as ações de outro módulo.

---

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
| `modulo_saida_osme.py` | Saída OSME | Yes (MIGO Z61) |
| ~`modulo_transferencia_equipes.py`~ | Transf. Equipes | *(Desativado/Movido para modulos_desativados)* |
| `modulo_contagem_carro.py` | Contagem de Carro | Yes (ZMMCSE001 / IW39 / MIGO Z62 + Z61) |
| `modulo_atualizar_osme.py` | Atualizar OSME | Yes (IW39) |
| `modulo_contagem_deposito.py` | Contagem Depósito (Lote 0 ⟷ ZM29) | No (Offline / Excel) |

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

## Funções & Skills Disponíveis (Agente Orquestrador)

O chatbot atua como um **Agente Orquestrador Central** capaz de acionar módulos do N.E.X.U.S tanto via **OpenRouter (LLM Tool Calling)** quanto pelo **Motor Local Integrado (Offline / NLP Regex)**.

### executar_bandeirada(recursos)
Inicia o processo de Bandeirada no SAP para uma ou mais equipes/recursos.
- **Comandos naturais:** `"bandeira o csc300"`, `"bandeirar csc300, csc301"`, `"bandeirada csc300"`
- **Parâmetros:** `recursos` (lista de strings) — ex: `["CSC300"]`
- **Efeito:** Alterna para a aba Bandeirada, insere as equipes e aciona `iniciar_bandeirada()`.

### extrair_osme(recursos)
Inicia a extração de OSME no SAP para equipes.
- **Comandos naturais:** `"extrair osme csc300"`, `"osme csc300"`, `"osme do csc300"`
- **Parâmetros:** `recursos` (lista de strings) — ex: `["CSC300"]`
- **Efeito:** Alterna para a aba OSME, preenche o campo de equipes e inicia a extração.

### consultar_migo(equipe, ordem)
Consulta a ordem ativa na IW39 e os materiais/reservas na MIGO para Equipes Clandestinas (RCL001 a RCL008) ou equipes gerais.
- **Comandos naturais:** `"migo rcl002"`, `"consultar migo rcl001"`, `"ordem rcl003"`
- **Parâmetros:** `equipe` (string, ex: `"RCL001"`), `ordem` (string opcional)
- **Efeito:** Alterna para Consultar MIGO, seleciona a equipe e dispara o fluxo IW39 ➡️ MIGO.

### transferir_osme(equipes)
Prepara e inicia o processo de Transferência de OSME (Z62 e Z61 no SAP).
- **Comandos naturais:** `"transferir osme csc300"`, `"transferencia osme csc300"`
- **Parâmetros:** `equipes` (lista de strings)
- **Efeito:** Alterna para Transf. OSME, carrega a fila de equipes e inicia o fluxo.

### contagem_carro(equipe, seriais, acao)
Executa a auditoria e contagem física/balanço de medidores em veículos de equipes no SAP:
- **Comandos naturais:** `"contagem de carro csc300"`, `"contar carro csc300"`, `"balanço carro csc300"`
- **Parâmetros:** `equipe` (string, ex: `"CSC300"`), `seriais` (opcional), `acao` (`"abrir"`, `"auditar"`, `"executar"`)
- **Efeito:** Audita medidores na ZMCSE01 (`ZMMCSE001`), busca ordem mais recente na IW39, estorna via MIGO Z62 e reaponta via MIGO Z61, exibindo pop-up com os documentos gerados.

### mudar_recurso(origem, destino, acao)
Realiza consulta de saldo ou transferência integral entre equipes no SAP.
- **Comandos naturais:** `"mudar recurso do csc395 para cvc396"`, `"saldo origem csc395"`
- **Parâmetros:** `origem` (string), `destino` (string opcional), `acao` (`"consultar"` ou `"executar"`)
- **Efeito:** Alterna para Mudança de Recurso, define os campos e aciona a operação.

### saida_osme(equipe, seriais, material, quantidade)
Prepara a Saída de medidores ou materiais não serializados para uma OSME/equipe no SAP MIGO Z61.
- **Comandos naturais:** `"processa na osme do cvc358 510388 6 un"`, `"saida osme csc300 seriais 18398041 a 18398045"`
- **Parâmetros:** `equipe` (string), `seriais` (opcional), `material` (opcional), `quantidade` (opcional)
- **Efeito:** Carrega os dados na aba Saída OSME, valida contra FORNECIDOS e solicita confirmação. Ao responder `"confirmo"` ou `"sim"` no chat, dispara a saída no SAP automaticamente.

### transferir_entre_equipes(destino, seriais)
Prepara a transferência de medidores entre equipes (estorno Z62 imputado à OSME de origem e saída Z61 imputada à OSME de destino no SAP MIGO).
- **Comandos naturais:** `"transferir para o csc300 os medidores 18398041 a 18398045"`, `"transferir seriais para cvc395"`
- **Parâmetros:** `destino` (string, equipe destino), `seriais` (string ou lista)
- **Efeito:** Valida seriais em `FORNECIDOS`, descobre equipe e OSME de origem, alerta não-fornecidos, atualiza planilha e prepara transferência. Ao responder `"confirmo"` ou `"sim"` no chat, executa o estorno e saída no SAP.

### abrir_conferencia(base, fisico)
Abre a aba de Conferência física de seriais/materiais com bipagem ou comparação.
- **Comandos naturais:** `"conferir seriais"`, `"abrir conferencia"`, `"comparar seriais"`

### abrir_fichas(re, recurso)
Abre a aba de Fichas WorkForce e opcionalmente preenche o RE do eletricista e o recurso.
- **Comandos naturais:** `"abrir fichas"`, `"ficha re 19595"`, `"preencher ficha 19595"`

### abrir_modulo(modulo)
Navega instantaneamente para qualquer aba do sistema.
- **Comandos naturais:** `"abrir etiquetas"`, `"ir para sap"`, `"mostrar dashboard"`, `"abrir zm02"`

### consultar_iq09(seriais)
Pesquisa um ou múltiplos medidores/números de série (numéricos ou alfanuméricos como `UQM9HDN26L01533345`) ou intervalos diretamente na transação IQ09 do SAP.
- **Comandos naturais:** `"pesquise o medidor 18398041"`, `"olha esse medidor 123456789"`, `"dá uma olhada no medidor UQM9HDN26L01533345"`, `"checa esse equipamento 18400028"`, `"iq09 18400025 a 18400028"`, `"UQM9HDN26L01533345"` (código colado direto), blocos de seriais colados.
- **Parâmetros:** `seriais` (string ou lista, com suporte a alfanuméricos e intervalos)
- **Efeito:** Expande intervalos e listas, pesquisa no SAP IQ09, exibe Pop-up detalhado (`JanelaDetalheMedidor` para 1 medidor, `JanelaMultiplosMedidores` com tabela de resultados e exportação para múltiplos medidores) e retorna resumo no Chatbot com status (`LIDI DISP`), material, depósito e denominação.

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

1. **Bandeirada SAP por voz/chat:** "bandeira o csc300" → executa bandeirada no SAP e exibe status
2. **Consultar MIGO clandestina:** "migo rcl002" → busca ordem IW39 e abre MIGO no SAP
3. **Extração rápida de OSME:** "extrair osme csc300" → abre OSME e inicia extração
4. **Mudança de recurso:** "mudar recurso do csc395 para cvc396" → transfere materiais no SAP
5. **Buscar lote e gerar etiqueta:** buscar_lote → gerar_etiqueta
6. **Gerar documento no Drive:** buscar_lote → gerar_doc_google
7. **Extrair e copiar seriais:** extrair_seriais → copiar_seriais
8. **Logar no SAP:** login_sap com nome do usuario

## Como Criar Novas Skills em Markdown (.md)

O N.E.X.U.S suporta o carregamento de novas habilidades dinamicamente a partir de arquivos `.md` na pasta `skills/`.
**Você não precisa alterar o código-fonte nem recompilar o .exe para criar novas skills!**

### Estrutura de um arquivo `skills/minha_skill.md`:

```markdown
---
name: nome_da_skill
description: "Breve explicação do que a skill faz para a IA (OpenRouter e Motor Local)"
triggers:
  - gatilho principal
  - sinonimo 1
  - comando natural
module: nome_do_modulo_python   # opcional: modulo onde a funcao esta
action: nome_da_funcao          # opcional: funcao python para executar
parameters:
  parametro1:
    type: string
    description: "Descricao do parametro"
    required: true
---

# Titulo da Skill

Documentação detalhada e instruções de suporte para guiar o assistente virtual.
```

### Onde colocar os arquivos:
1. Em ambiente de desenvolvimento: na pasta `skills/` na raiz do projeto.
2. Em produção com o executável: basta criar ou manter a pasta `skills/` no mesmo diretório do `N.E.X.U.S.exe`!

## Segurança

- `logins_sap.json` e `config_chat.txt` estão no `.gitignore`
- Nunca commite senhas ou chaves de API

## Backlog de Melhorias & Ajustes Futuros

- [ ] **Módulo Transf. Equipes (Desativado / Aguardando Revisão)**:
  - Tratar/remover entradas duplicadas de seriais/itens (fazer deduplicação automática ao carregar ou processar na fila).
  - Adicionado botão de continuidade direto no topo (card de seriais), na tabela de conferência, no modal de alerta de não fornecidos e barra fixa no rodapé da janela para nunca sumir no scroll.
  - Campo opcional `📅 Data Lçto. (Fechamento SAP)` no Card de Destino e injeção da sub-rotina `PreencherDataLancamento` em ambos os movimentos MIGO (Z62 e Z61) para superar bloqueios de fechamento contábil.
- [x] **Módulo Saída OSME (`src/modulo_saida_osme.py`)**:
  - Campo opcional `📅 Data Lçto. (Fechamento SAP)` no Card de Equipe / OSME.
  - Injeção da sub-rotina `PreencherDataLancamento` no script MIGO Z61 (`ctxtGOHEAD-BUDAT`) para permitir lançamento contábil em períodos de fechamento do SAP.
- [x] **Módulo Transferência de OSME (`src/modulo_transferenciaosme.py` / `modulos/transferencia_osmes_final.py`)**:
  - Campo opcional `📅 Data Lçto. (Fechamento SAP)` na interface com validação estrita de calendário (`dd.mm.aaaa`).
  - Função `_preencher_data_lancamento` no backend Python COM para aplicar a data de lançamento na aba Geral do Cabeçalho da MIGO (Z62 e Z61), expandindo o cabeçalho se recolhido, disparando Enter e tratando modais para viabilizar fechamentos contábeis.
- [x] **Módulo Atualizar OSME (`src/modulo_atualizar_osme.py`)**:
  - Varredura automática na IW39 (SAP) recurso por recurso (`DY_PARVW="Z2"`, `DY_PARNR=equipe`).
  - Identifica ordem e vencimento tratando ordens únicas e múltiplas (grid ALV ordenado decrescente por `GLTRP`).
  - **Filtro Inteligente por Vencimento**: pula automaticamente equipes cuja OSME atual ainda está vigente (`Vencimento > Hoje`), reduzindo o tempo de varredura de ~25 minutos para ~2 minutos (apenas equipes vencidas ou sem ordem são consultadas no SAP).
  - **Gravação Dinâmica por Equipe em Tempo Real**: busca a linha da equipe pelo nome na Coluna E no momento da gravação, evitando sobrescrever linhas erradas se a planilha tiver sido classificada ou reordenada no Excel.
  - Gravação resiliente com suporte à planilha aberta no Excel local via COM ou gravação direta via openpyxl.
  - Fallback automático na Fila de Pendências (`utils_pendencias_planilha.py`) com banner de sincronização e timer de 20s em segundo plano quando a planilha estiver bloqueada ou em uso na rede.
- [x] **Gerenciamento de Bandeja do Windows & Fechamento Seguro (`src/utils_tray.py` / `src/conecta_hub2.py`)**:
  - Interceptação segura do botão 'X' (`WM_DELETE_WINDOW`): se houver operações pendentes de gravação, alerta o usuário e permite minimizar para a bandeja.
  - Ícone de bandeja do sistema (System Tray) com `pystray`: mantém o N.E.X.U.S rodando em segundo plano com timers de auto-sincronização ativos para gravar na planilha assim que o Excel for liberado.
  - Menu de contexto na bandeja: restaurar janela, sincronizar agora e encerrar com confirmação para nunca perder dados de fallback acidentalmente.
- [x] **Módulo Contagem de Carro (`src/modulo_contagem_carro.py`)**:
  - Campo opcional `Data de Lançamento (dd.mm.aaaa)` no Card 1 para superar bloqueios contábeis de fechamento de mês no SAP (dias 28 a 31).
  - Sub-rotina VBScript `PreencherDataLancamento` injetada com segurança em MIGO Z62 e Z61 (`ctxtGOHEAD-BUDAT` + `sendVKey 0`).
  - Botão `🧹 Remover Divergentes da Lista` no Card 2 para limpar seriais inválidos do campo de texto com um clique após auditoria na ZMCSE01.
  - Classificação estrita de divergências (fora de LIDI/DISP, com nota de serviço, ou sem ordem de origem) impedindo que quebrem a execução no SAP.
  - Fechamento proativo de popups informativos no loop de espera da grade ALV da ZMMCSE001.
- [x] **Botão de Emergência / "Interromper SAP" na Barra de Status (`src/conecta_hub2.py` / `src/modulo_Sap.py`)**:
  - Botão vermelho `🛑 Interromper SAP` exibido dinamicamente na barra de status sempre que qualquer automação SAP estiver ocupada (`is_sap_ocupado() == True`).
  - Mata processos filhos `cscript.exe` com segurança, aciona flag `is_abortar_solicitado()`, reseta estado de ocupado e libera a UI imediatamente.
- [x] **Redesign de Navegação — Abas Laterais / Sidebar (`src/sidebar_navegacao.py` / `src/conecta_hub2.py`)**:
  - Setas indicativas dinâmicas no botão de alternância (`◀` para recolher, `▶` para expandir) com tooltips contextuais.
  - Setas de categorias mais nítidas e bold (`▾` / `▸`), e centralização limpa de ícones no modo recolhido sem corte de texto.
  - Barra de busca rápida no topo da Sidebar (Omnibar com atalho `Ctrl+F`) para filtragem instantânea de módulos em tempo real.
  - Eliminação do vão preto à direita via layout dinâmico `grid` no `frame_corpo` (conteúdo expande automaticamente para 100% da largura).
- [x] **Módulo Extrator OSME (`src/modulo_extrair_OSME.py` / `scripts/Sciprt OSME2.vbs`)**:
  - Correção da substituição do campo de equipe na transação `ZUDWM_237` (`ctxtS_EQUIPE-LOW`), eliminando o bug que prendia a execução em `csc300`.
  - Ordenação decrescente por Ordem (`AUFNR` com `&SORT_DSC`) para capturar sempre a OSME ativa mais recente na linha 0.
  - Confirmação robusta de janelas de diálogo do SAP (`btn[0]` e `btn[11]` para substituição de arquivos).
  - Loop de espera COM resiliente na impressão automática do Excel evitando erro `(-2147221021, 'Operação não disponível')`.
  - Alerta sonoro nativo do Windows ao concluir a extração.

- [x] **Módulo Mudança de Recurso (`modulos/mudanca_recurso.py` e `src/modulo_mudanca_recurso.py`)**:
  - **Inversão de Fluxo com Fail-Fast Pré-Voo**: A consulta e validação da ordem ativa de destino na IW39 foi antecipada para a **Etapa 2/4** (antes de qualquer MIGO). Se a ordem de destino não for encontrada, o processo é abortado imediatamente com **zero movimentação contábil no SAP**, impedindo que o estorno Z62 seja executado sem destino garantido.
  - **Busca Resiliente na IW39**: Tratamento automático de janelas modais informativas (`wnd[1]`), fallback de casing minúsculo/maiúsculo para o parceiro Z2 (`cvc308` e `CVC308`), captura via Regex no título e busca na grade ALV ordenada por vencimento (`GLTRP`).
  - **Ordem de Destino Manual / Teste Pré-Voo**: Campo opcional `Ordem Destino (opcional / fallback)` e botão `🔍 Testar IW39` na interface gráfica para validar a ordem antes da execução.
  - **Modo de Recuperação Direta Saída Z61**: Botão `⚡ Concluir Saída Z61 (Recuperação)` para imputar no SAP materiais/seriais previamente estornados sem necessidade de retrabalho manual.
- [x] **Módulo Contagem de Depósito (`src/modulo_contagem_deposito.py`)**:
  - **Flexibilização de Entrada**: Suporte a upload de planilha idêntica à ZM29 com a coluna/aba 'Estoque Livre' preenchida com a contagem física do depósito, superando o modelo restrito à planilha beta de seriais.
  - **Suporte a Arquivo Único com Duas Abas**: Permite cruzar ZM29 e contagem física presentes no mesmo workbook (ex: aba 'ZM29' e aba 'Estoque Livre').
  - **Botão '📥 Puxar Lote 0 da Planilha Beta'**: Adicionado no topo do layout para extrair automaticamente os dados de Lote 0 da planilha `CONTROLE_SERIAIS.vBeta.xlsm`, consolidar as quantidades por material para o depósito ativo (`CL03`/`CL04`) e carregar automaticamente no campo de contagem física.
  - **Engine de Cruzamento Robusta**: Tratamento de unidades brasileiras, ordenação dinâmica por divergência na Treeview e exportação de relatório Excel (.xlsx) com métricas de acurácia e destaque visual de status (🟢 Bateu, 🔴 Faltou, 🟡 Sobrando).
- [x] **Sistema de Atualização Automática Híbrido & Patch Notes Acumulativas (`src/utils_updater.py` / `versao.json`)**:
  - **Consulta Híbrida**: Checagem de novas versões primeiro online via GitHub público (`raw.githubusercontent.com/.../versao.json`) com fallback transparente para a rede corporativa (`Drive X:`).
  - **Patch Notes Acumulativas (Forma A)**: O campo `notas` no `versao.json` acumula as versões recentes agrupadas por cabeçalho (ex: `[v2.2.2]`, `[v2.2.1]`), permitindo que operadores que pulem versões leiam todo o histórico de mudanças.
  - **Encerramento e Inicialização Seguros**: Script `.bat` de atualização higieniza as variáveis herdadas do PyInstaller (`_MEIPASS2`, `PYTHONHOME`, `PYTHONPATH`) evitando o erro `Failed to import encodings module`, e utiliza loop de encerramento em árvore (`taskkill /F /T`) com polling no `tasklist` até a liberação completa do mutex de instância única.
  - **Modal de Atualização Moderno (`540x510`)**: Layout espaçoso com botões fixos no rodapé (`side="bottom"`), barra de progresso em tempo real e caixa de texto com scroll para notas longas.
