# Runbook Operacional - ProcessoASO

## Pre-check
1. Confirmar acesso a rede e caminhos base.
2. Confirmar credenciais e variaveis de ambiente.
3. Confirmar disponibilidade dos sistemas integrados.
4. Confirmar acesso de escrita em `P:\ProcessoASO\logs`, `P:\ProcessoASO\json` e `P:\ProcessoASO\relatorios`.
5. Confirmar disponibilidade de Tesseract, Poppler, Outlook e Yube.

## Execucao
1. Executar o metodo oficial do robo (bat, exe ou script homologado).
2. Acompanhar o terminal no padrao ENESA e validar a etapa atual.
3. Acompanhar os logs tecnicos em `logs/`.
4. Validar manifesto em `json/` e resumo em `relatorios/` ao final.

## Pos-execucao
1. Conferir classificacao em `processados/` e `erros/`.
2. Registrar incidente se houver erro critico.
3. Atualizar evidencias de auditoria quando aplicavel.
4. Validar que o mesmo `execution_id` aparece em log, manifesto e relatorios.
5. Em caso de falha, anexar os caminhos dos artefatos na evidencia ou no ticket.
