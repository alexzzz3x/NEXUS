---
name: consultar_migo
description: "Consulta ordens ativas na IW39 e materiais/reservas na MIGO no SAP para Equipes Clandestinas (RCL001 a RCL008) ou equipes gerais"
triggers:
  - migo
  - consultar migo
  - clandestina
  - clandestinas
  - ordem rcl
  - rcl
module: conecta_hub2
action: consultar_migo
parameters:
  equipe:
    type: string
    description: "Código da equipe (ex: RCL001, RCL002)"
    required: true
  ordem:
    type: string
    description: "Número da ordem IW39 opcional"
    required: false
---

# Skill: Consultar MIGO / IW39 (Clandestinas)

Consulta ordens de serviço ativas para equipes clandestinas no SAP.
Alterna para a aba Consultar MIGO e realiza a busca encadeada IW39 -> MIGO.
