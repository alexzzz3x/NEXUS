---
name: executar-bandeirada
description: Executar, diagnosticar ou manter a automacao SAP de Bandeirada deste projeto, cobrindo a selecao da ordem da equipe na IW39 e a acao correspondente na ZUDWM_213. Usar quando o pedido mencionar Bandeirada, IW39, ZUDWM_213, escolha de ordem por recurso/equipe, erro de controle SAP nesse fluxo ou alteracao do modulo bandeirada.py.
---

# Executar Bandeirada

## Objetivo

Operar e manter o fluxo `IW39 -> ZUDWM_213` sem reutilizar ordem de outra equipe e sem ocultar falhas do SAP.

## Localizar a implementacao

- Usar `modulos/bandeirada.py` como unica implementacao do programa.
- Tratar o VBScript temporario como artefato descartavel e sempre removê-lo ao final.

## Executar o fluxo

1. Receber e normalizar o codigo da equipe.
2. Exigir uma sessao SAP GUI aberta e autenticada. Nao implementar nem armazenar credenciais.
3. Abrir `/nIW39`.
4. Selecionar a funcao de parceiro `Z2`, informar somente a equipe solicitada e marcar `DY_MAB`.
5. Executar a consulta e ordenar o grid pela data `GLTRP` conforme o fluxo vigente.
6. Ler a ordem `AUFNR` da primeira linha valida. Quando o SAP abrir diretamente o detalhe de uma unica ordem, ler `CAUFVD-AUFNR`.
7. Interromper se nenhuma ordem for identificada; nunca usar uma ordem antiga ou de outro recurso.
8. Abrir `ZUDWM_213`, preencher `P_AUFNR` com a ordem obtida e executar.
9. Conferir a ordem exibida no detalhe antes da acao final.
10. Pressionar a acao da transacao e confirmar o popup somente quando o usuario tiver autorizado a execucao real.
11. Registrar no log a equipe, a ordem selecionada, o retorno do script e qualquer mensagem de erro.

## Diagnosticar e corrigir

- Para controle SAP ausente, verificar primeiro a transacao e a tela atualmente abertas; depois comparar o ID gravado com o ID vigente.
- Reobter grids depois de ordenar ou navegar, pois referencias COM podem ficar invalidas.
- Preferir esperas por controle com limite de tempo a pausas fixas longas.
- Tratar retorno diferente de zero do `cscript.exe` como falha, mesmo que exista texto de saida.
- Preservar o codigo da equipe informado durante toda a execucao.

## Validar alteracoes

1. Executar `py_compile` nos modulos alterados.
2. Verificar que nenhum dado de credencial foi introduzido.
3. Testar logica e geracao do script sem acionar o SAP quando a autorizacao for apenas para manutencao.
4. Para teste real, avisar que a etapa final modifica dados no SAP e obter confirmacao explicita.
5. Recompilar o executavel final e relatar o caminho gerado.
