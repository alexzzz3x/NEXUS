---
name: executar-transferencia-osmes
description: Executar, diagnosticar ou manter a transferencia completa de materiais de OSMEs vencidas para a ordem valida do mesmo recurso, usando ZUDWM_237, conferencia de seriais na Analise de equipe, MIGO Z62 e MIGO Z61 com captura dos documentos SAP. Usar quando o pedido mencionar transferencia de OSMEs, recurso/equipe, ordem vencida ou valida, LIDI/DISP, ZMMCSE004, MIGO, Z62, Z61, materiais serializados ou nao serializados, documentos SAP ou os modulos transferencia_osmes*.py.
---

# Executar Transferencia de OSMEs

## Objetivo

Transferir exclusivamente os materiais encontrados na consulta atual do recurso: primeiro devolver os itens das ordens vencidas ao deposito com Z62 e depois enviar os mesmos itens para a unica ordem valida com Z61.

Antes de alterar controles ou sequencias da MIGO, ler [a referencia do script SAP](../sap_script_base_referencia.md).

## Localizar a implementacao

- Consultar e montar os CSVs em `modulos/transferencia_osmes.py`.
- Validar a consulta atual e executar Z62/Z61 em `modulos/transferencia_osmes_final.py`.
- Usar `main.py` como unica interface do programa.

## Coletar e validar os itens

1. Receber uma equipe/recurso alfanumerico e manter esse valor como escopo imutavel.
2. Executar uma consulta nova na `ZUDWM_237`; nunca reutilizar CSV antigo.
3. Identificar exatamente uma ordem nao vencida como destino.
4. Coletar somente saldos positivos das ordens vencidas.
5. Incluir materiais serializados e nao serializados.
6. Extrair todos os seriais dos materiais serializados e consultar em lote por `ZMMCSE001 > Analise de equipe`.
7. Percorrer todo o grid por blocos para carregar linhas fora da area visivel.
8. Comparar seriais numericos ignorando zeros a esquerda. Exemplo: `18389335` equivale a `000000000018389335`.
9. Aprovar o serial somente quando o recurso for o consultado e os estados forem `LIDI`, `DISP` e movimento `Disponivel`.
10. Tratar serial nao localizado como consulta inconclusiva, nunca como comprovadamente inapto.
11. Bloquear o fluxo se equipe, destino, quantidade ou conferencia nao corresponderem a consulta atual.

## Executar a MIGO

Obter confirmacao explicita antes de registrar movimentos reais.

### Z62

1. Abrir `/nMIGO` com acao `A01`, referencia `R10` e movimento `Z62`.
2. Para cada item, usar sua ordem vencida como ordem contabil.
3. Informar operacao `0010`, centro `EMPR`, deposito `CL03` e recebedor `EMPR 1CL03`.
4. Preencher material, quantidade e seriais aprovados.
5. Avancar entre itens por `btnOK_NEXT_ITEM`; nao escrever diretamente na proxima linha do grid.
6. Verificar o documento, interromper em popup ou mensagem de erro e registrar somente se estiver `Documento O.K.`.
7. Capturar imediatamente o numero de `Doc.material` na barra de status.
8. Persistir o documento antes de iniciar a Z61.

### Z61

1. Iniciar somente depois de capturar e persistir o documento Z62.
2. Abrir a MIGO com acao `A07`, referencia `R10` e movimento `Z61`.
3. Repetir exatamente os itens efetivamente registrados na Z62.
4. Usar somente a ordem valida atual como ordem contabil de todos os itens.
5. Verificar, registrar e capturar o documento SAP da mesma forma usada na Z62.

## Preservar o historico

- Gravar Z62 e Z61 em `Documents/Assistente Almox/Historico_Transferencias_SAP.csv`.
- Registrar identificador da execucao, equipe, movimentos, documentos, origens, destino, itens, mensagem SAP e CSV de origem.
- Se o CSV principal estiver bloqueado, gravar em `Historico_Transferencias_SAP_Pendente_<execucao>.csv` e continuar o fluxo.
- Se a Z62 tiver sido registrada e ocorrer falha antes da Z61, nao reiniciar o processo completo. Recuperar o documento da barra de status ou do historico e retomar somente a Z61 mediante autorizacao.

## Diagnosticar com seguranca

- Em pedidos de diagnostico, limitar-se a leituras do SAP e dos CSVs; nao registrar movimentos.
- Distinguir controle ausente, linha nao carregada, serial com zeros a esquerda e divergencia real de status.
- Depois de ordenar um ALV, obter novamente o objeto do grid antes de ler celulas.
- Nunca misturar dados de equipes, consultas ou datas diferentes.
- Nao repetir Z62 ou Z61 sem confirmar no historico e na barra de status se o documento anterior existe.

## Validar alteracoes

1. Executar `py_compile` nos arquivos alterados.
2. Testar normalizacao, agrupamento, sequencia de itens e fallback do historico com mocks ou dados locais.
3. Nao usar uma execucao SAP real como teste automatico sem autorizacao explicita.
4. Recompilar o executavel final somente após aprovacao.
5. Relatar documentos capturados, arquivos gravados e qualquer etapa que tenha ficado pendente.
