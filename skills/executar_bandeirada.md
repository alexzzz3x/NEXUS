---
name: executar_bandeirada
description: "Executa o processo de Bandeirada no SAP para uma ou mais equipes/recursos"
triggers:
  - bandeira
  - bandeirar
  - bandeirada
  - fazer bandeirada
  - iniciar bandeirada
module: conecta_hub2
action: executar_bandeirada
parameters:
  recursos:
    type: array
    items:
      type: string
    description: "Lista com códigos de recursos/equipes (ex: ['CSC300', 'CDC410'])"
    required: true
---

# Skill: Executar Bandeirada

Inicia a automação de Bandeirada no SAP para os recursos informados pelo usuário.
Alterna automaticamente para a aba Bandeirada e despacha a fila de equipes.
