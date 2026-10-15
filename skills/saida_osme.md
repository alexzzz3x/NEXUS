---
name: saida_osme
description: "Prepara a Saída para OSME no módulo de Saída OSME para uma equipe, validando seriais e calculando quantidades consolidadas para conferência prévia antes do envio ao SAP MIGO"
triggers:
  - saida osme
  - saída osme
  - saida para osme
  - saída para osme
  - processa na osme
  - processe na osme
  - processar na osme
  - baixa na osme
  - dar baixa na osme
  - baixa osme
  - lancar na osme
  - lançar na osme
module: conecta_hub2
action: saida_osme
parameters:
  equipe:
    type: string
    description: "Código da equipe/recurso (ex: CVC395, CSC300)"
    required: true
  seriais:
    type: string
    description: "Seriais ou intervalo de medidores (ex: '18398041 a 18398045')"
    required: false
  material:
    type: string
    description: "Código do material não serializado (ex: '330702', '510388')"
    required: false
  quantidade:
    type: string
    description: "Quantidade do material não serializado (ex: '50m', '50', '20 un')"
    required: false
---

# Skill: Saída para OSME

Esta habilidade prepara a tela e os dados de Saída para OSME:
- Preenche a equipe e localiza a OSME ativa, depósito, segmento e vencimento na aba CADASTRO.
- Valida seriais informados contra a aba FORNECIDOS para impedir duplicidades e contra a aba ESTOQUE.
- Permite a inclusão de materiais não serializados (cabos, postes, conectores) com busca automática de descrição e unidade.
- **Confirmação e Execução**: O robô prepara os dados e solicita confirmação ("Deseja processar a saída no SAP agora?"). Quando o usuário responder "confirmo", "sim" ou "pode executar", o assistente executa a saída no SAP MIGO automaticamente.

