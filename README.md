
# ASO Automation Project

## Visão Geral
Este projeto automatiza o processamento de Atestados de Saúde Ocupacional (ASO) recebidos via e-mail.
O sistema monitora uma caixa de entrada do Outlook, identifica e-mails contendo ASOs, extrai informações via OCR (Tesseract) e integra com um sistema RPA (Yube).

## Funcionalidades Principais
- **Monitoramento de E-mail**: Verifica a caixa de entrada em busca de e-mails com assuntos padronizados.
- **Processamento de PDF**: Converte PDFs em imagens e aplica OCR para extração de dados.
- **Extração de Dados**: Identifica Nome, CPF, Data do ASO e Função/Cargo.
- **Integração RPA**: Prepara os arquivos e aciona o bot RPA para cadastro.
- **Relatórios**: Gera relatórios JSON e envia resumos por e-mail.

## Pré-requisitos

### Sistema
- **Windows OS** (Necessário para automação via `win32com` Outlook).
- **Microsoft Outlook** instalado e configurado com a conta alvo.
- **Tesseract OCR**:
  - Instalar o [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract/wiki).
  - Adicionar ao PATH ou configurar no `.env`.
- **Poppler**:
  - Necessário para `pdf2image`.
  - Baixar e extrair o binário, configurar caminho no `.env`.

### Python
- Python 3.10+
- Dependências listadas em `requirements.txt`.

## Instalação

1. Clone o repositório.
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Configure o arquivo `.env`:
   - Copie `.env.example` para `.env`.
   - Preencha os caminhos do Tesseract, Poppler e conta de e-mail.

## Configuração (.env)
```ini
# Exemplo
TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe
POPPLER_PATH=C:\Installs\poppler-24.08.0\Library\bin
ASO_EMAIL_ACCOUNT=aso@enesa.com.br
ASO_MAILBOX_NAME=Aso
ASO_DAYS_BACK=0  # 0 = Apenas hoje
```

## Utilização

Para rodar o processo manualmente:
```bash
python main.py
```
O script irá:
1. Ler os e-mails do dia (ou janela configurada).
2. Baixar e processar anexos.
3. Gerar arquivos na pasta de rede configurada.
4. Acionar o RPA.
5. Enviar um e-mail de resumo ao final.

## Estrutura de Pastas
- `main.py`: Script principal de orquestração.
- `custom_logger.py`: Módulo de logs estruturados.
- `reporting.py`: Geração de relatórios JSON.
- `notification.py`: Envio de e-mail de resumo.
- `aso_admissional_email.py`: (Legado/Alternativo) Módulo de e-mail.
- `logs/`: Logs de execução.
- `relatorios/`: Relatórios gerados.
- `tests/`: Testes automatizados.

## Testes
Para executar os testes unitários:
```bash
pytest tests/
```
