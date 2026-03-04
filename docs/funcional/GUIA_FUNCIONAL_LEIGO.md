# Guia Funcional (Publico Leigo) - ProcessoASO

## O que este robo faz
Ler emails de ASO, identificar os documentos validos, extrair os dados principais e encaminhar os arquivos para o processamento automatizado.

## Quando ele roda
Roda conforme a rotina operacional homologada pela area responsavel. Pode ser executado manualmente ou pelo metodo oficial da operacao.

## O que entra
- Emails com anexos PDF de ASO
- Links de Google Drive, quando presentes no corpo do email
- Credenciais e configuracoes do ambiente

## O que sai
- Arquivos classificados em `processados/` ou `erros/`
- Log tecnico em `logs/`
- Manifesto de execucao em `json/`
- Resumo de execucao em `relatorios/`

## Como saber se deu certo
1. Relatorio de execucao gerado.
2. Itens em `processados/`.
3. Ausencia de erro critico no log.

## O que fazer se der erro
1. Verificar a pasta `erros/`.
2. Consultar manifesto, relatorio e log pelo mesmo `execution_id`.
3. Acionar suporte tecnico.
