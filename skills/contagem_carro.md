---
name: contagem_carro
description: "Executa a Contagem de Carro no SAP: audita medidores na ZMCSE01 (ZMMCSE001), localiza a ordem mais recente na IW39, realiza o estorno MIGO Z62 da ordem atual e a re-saída MIGO Z61 para a ordem mais recente."
triggers:
  - contagem de carro
  - contagem carro
  - contar carro
  - balanço carro
  - balanco carro
  - auditoria carro
  - contagem veículo
  - contagem veiculo
module: modulo_contagem_carro
action: contagem_carro
parameters:
  equipe:
    type: string
    description: "Código da equipe/recurso a auditar (ex: CSC300)"
    required: true
  seriais:
    type: string
    description: "Lista de seriais bipados/contados no veículo"
    required: false
  acao:
    type: string
    description: "'abrir', 'auditar' (ZMCSE01) ou 'executar' (ciclo completo)"
    required: false
---

# Skill: Contagem de Carro (Balanço de Medidores)

Executa a rotina de inventário/contagem física de medidores nos veículos das equipes de campo no SAP:
1. Bipagem/preenchimento dos medidores contados fisicamente no veículo no Nexus (não puxa nada da ordem SAP).
2. Auditoria obrigatória na **ZMCSE01** (`ZMMCSE001` — Análise de Equipe) exclusivamente sobre os medidores bipados para validar status (`LIDI`, `DISP`), código do material (`MATNR`) e ordem atual (`AUFNR`). Medidores que **não forem LIDI / DISP** (bloqueados, sucata, montados ou não localizados) são alertados ao operador e **ignorados**, permitindo que o processo prossiga com os medidores válidos.
3. Localização na **IW39** da ordem ativa mais recente para a equipe (Parceiro Z2).
4. Devolução/estorno ao depósito via **MIGO Z62** imputando na ordem de origem onde o medidor estava.
5. Saída/imputação via **MIGO Z61** para a ordem ativa localizada na IW39 (seja a mesma ordem ou a mais atual).
6. Captura e exibição dos Documentos SAP gerados (`Doc Z62` e `Doc Z61`), com relatório de medidores processados e ignorados, e gravação no histórico CSV.
