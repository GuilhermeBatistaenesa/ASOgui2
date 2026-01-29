
# 📄 Documentação Técnica Oficial - RPA ASO Automation

**Versão Documento**: 1.0  
**Data**: 28/01/2026  
**Projeto**: Automação de Recebimento de Atestados de Saúde Ocupacional (ASO)

---

## 1. Visão Geral do Sistema
Este sistema é uma solução de automação robótica de processos (RPA) desenvolvida para otimizar o fluxo de recebimento e cadastro de ASOs. A solução monitora endereços de e-mail corporativos, identifica mensagens contendo atestados, realiza o processamento de imagens (OCR) para extração de dados e integra essas informações com o sistema de gestão via bot RPA.

### 1.1 Objetivo
Eliminar a triagem manual de e-mails e a digitação de dados de atestados médicos, garantindo agilidade e padronização no cadastro.

### 1.2 Arquitetura
O sistema opera em uma máquina local Windows e utiliza as seguintes tecnologias:
- **Linguagem**: Python 3.10+
- **Integração E-mail**: `pywin32` (Outlook COM Interface)
- **Processamento de Imagem**: `pdf2image` (Poppler) e `pytesseract` (Tesseract OCR)
- **RPA Web/Desktop**: Módulo proprietário `rpa_yube`

---

## 2. Estrutura de Módulos

### `main.py`
Núcleo da aplicação. Responsável por:
1. Conectar ao Outlook e filtrar e-mails da data corrente.
2. Validar anexos PDF e convertê-los em imagens.
3. Executar o OCR para identificar:
   - Nome do Funcionário
   - CPF
   - Data do Exame
   - Função/Cargo
4. Decidir se o arquivo é válido ou um "Rascunho".
5. Orquestrar a chamada ao módulo RPA.

### `custom_logger.py`
Gerenciador de logs estruturados (JSONL) e saída de console.
- **Formato**: JSON (arquivo) e Texto formatado com ícones (console).
- **Localização**: `logs/execution_log_YYYY-MM-DD.jsonl`

### `reporting.py`
Gerador de relatórios de execução.
- **Saída**: JSON completo e Resumo Markdown.
- **Localização**: `relatorios/`

### `notification.py`
Módulo responsável pelo envio do e-mail de resumo ao final do processamento, informando estatísticas de sucesso e erros.

---

## 3. Fluxo de Dados

1. **Entrada**: E-mail recebido no Outlook contendo "ASO" e "ADMISSIONAL" no assunto.
2. **Processamento**:
   - Download do anexo PDF.
   - Conversão PDF -> JPG (Memória).
   - OCR (Tesseract) -> Texto Bruto.
   - Regex Parsing -> Dados Estruturados (Nome, CPF, etc.).
3. **Saída Intermediária**: Arquivo PDF renomeado (`Nome - CPF.pdf`) salvo na pasta de rede.
4. **Integração**: Acionamento do `rpa_yube` apontando para a pasta processada.
5. **Relatório**: Compilação dos dados e envio de notificação.

---

## 4. Configuração e Variáveis (.env)

O sistema utiliza um arquivo `.env` na raiz para parametrização.

| Variável | Descrição | Exemplo |
| :--- | :--- | :--- |
| `TESSERACT_PATH` | Caminho absoluto para o executável do Tesseract. | `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `POPPLER_PATH` | Caminho para binários do Poppler. | `C:\Tools\poppler\bin` |
| `ASO_EMAIL_ACCOUNT` | E-mail alvo para monitoramento. | `aso@empresa.com.br` |
| `ASO_MAILBOX_NAME` | Nome da caixa no Outlook (exibição). | `Caixa ASO` |
| `ASO_DAYS_BACK` | Dias retroativos para busca (0 = Hoje). | `0` |

---

## 5. Guia de Manutenção e Troubleshooting

### 5.1 Erro: "TesseractNotFound"
**Causa**: O Python não encontrou o executável do Tesseract.
**Solução**: Verifique se o `TESSERACT_PATH` no `.env` está correto e se o arquivo existe.

### 5.2 Erro: "Outlook Interface Error"
**Causa**: O Outlook não está aberto ou bloqueou a conexão COM.
**Solução**: Reinicie o Outlook e garanta que o usuário esteja logado.

### 5.3 Baixa Precisão do OCR
**Diagnóstico**:
- Verificar se os arquivos são "Rascunho" (marca d'água atrapalha).
- Verificar log com contexto (`DEBUG OCR`).
**Ajuste**: Melhorar as Regex no arquivo `main.py`, função `extrair_dados_completos`.

---

## 6. Procedimento de Testes
O projeto possui uma suíte de testes unitários para validar a lógica de extração sem depender de arquivos reais.

**Comando para execução**:
```bash
python -m pytest tests/
```
Cobertura atual: Validação de OCR (Nome, CPF, Data, Detecção de Rascunho).

---
**Responsável Técnico**: Equipe de Automação / Desenvolvimento.
