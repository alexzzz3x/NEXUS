---
name: consultar_iq09
description: "Consulta um ou múltiplos medidores, números de série (inclusive alfanuméricos) ou intervalos diretamente na transação IQ09 do SAP"
triggers:
  - iq09
  - medidor
  - medidores
  - serial
  - seriais
  - equipamento
  - equipamentos
  - status sap
  - disp
  - lidi
module: conecta_hub2
action: consultar_iq09
parameters:
  seriais:
    type: string
    description: "Número do medidor, serial, intervalo (ex: '18400025 a 18400028') ou lista de seriais"
    required: true
---

# Skill: Consultar IQ09 no SAP

Pesquisa direta de registros de equipamentos e medidores no SAP pela transação IQ09.

## Capacidades:
- Suporta seriais numéricos (ex: 18398041) e alfanuméricos (ex: UQM9HDN26L01533345).
- Suporta intervalos com expansão automática (ex: 18400025 a 18400030).
- Lê o ALV Grid completo (Material, Equipamento, Centro, Depósito, Status do Sistema como LIDI DISP).
- Abre Pop-up detalhado na interface gráfica do N.E.X.U.S.
