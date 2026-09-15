# Referência do fluxo SAP MIGO

Este documento resume os padrões aproveitados do script VBA fornecido como
base para o fluxo completo publicado no programa final.

## Abertura da MIGO

1. Abrir `/nMIGO`.
2. Se o detalhe estiver recolhido em `SAPLMIGO:0002`, pressionar
   `btnBUTTON_DETAIL`.
3. Trabalhar com o detalhe expandido em `SAPLMIGO:0003`.

## Movimentos

- Z62, devolução da ordem vencida para o depósito:
  - ação `A01`;
  - referência `R10`;
  - ordem contábil: ordem vencida do item.
- Z61, saída do depósito para a ordem válida:
  - ação `A07`;
  - referência `R10`;
  - ordem contábil: ordem válida do recurso consultado.

## Preenchimento do item

Para cada item:

1. Aba Material: `GOITEM-MAKTX`.
2. Aba Quantidade: `GOITEM-ERFMG`.
3. Aba Destino: centro `EMPR`, depósito `CL03` e recebedor `EMPR 1CL03`.
4. Aba Classificação contábil: ordem correspondente e operação `0010`.
5. Aba Serial, quando aplicável.
6. Pressionar `btnOK_NEXT_ITEM` para criar/ativar o próximo item.

O avanço não deve ser feito escrevendo diretamente na linha seguinte da grade.
O controle suportado pelo fluxo-base é `btnOK_NEXT_ITEM`.

## Verificação e registro

1. Verificar com `btn[7]`.
2. Interromper se o SAP abrir uma mensagem de erro.
3. Registrar com `btn[23]`.
4. Ler imediatamente a barra de status.
5. Extrair e armazenar o número de `Doc.material` antes de iniciar o próximo
   movimento.

## Regras adicionais da automação

- Todo o fluxo fica vinculado à consulta atual do recurso na `ZUDWM_237`.
- Seriais só entram quando a conferência atual confirma o mesmo recurso,
  `LIDI`, `DISP`, movimento `Disponível` e `Apto=SIM`.
- A leitura percorre todo o grid e normaliza zeros à esquerda em seriais
  numéricos antes da comparação.
- A Z61 só começa após a captura e persistência do documento Z62.
- Se o histórico principal estiver bloqueado, o documento é preservado em um
  CSV pendente da mesma execução, sem interromper o fluxo.
- Falhas de validação ou registro interrompem o fluxo para evitar duplicidade.
