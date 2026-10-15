---
name: extrair_osme
description: "Executa a extração de OSME no SAP para uma ou mais equipes/recursos"
triggers:
  - osme
  - extrair osme
  - puxar osme
  - extracao osme
  - extração osme
module: conecta_hub2
action: extrair_osme
parameters:
  recursos:
    type: array
    items:
      type: string
    description: "Lista com códigos de recursos/equipes para extrair OSME (ex: ['CSC300'])"
    required: true
---

# Skill: Extrair OSME

Inicia a extração automatizada de OSME para equipes no SAP.
Alterna para a aba OSME e executa a rotina de busca de pendências de OSME.
