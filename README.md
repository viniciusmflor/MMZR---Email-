# MMZR Family Office - Sistema de Relatorios

Sistema automatizado para geracao de relatorios HTML de performance de carteiras de investimento com integracao ao Microsoft Outlook.

**Versao:** 1.0.0  
**Plataforma:** Windows  

## Funcionalidades Principais

- Interface grafica intuitiva com selecao multipla de clientes
- Geracao automatizada de relatorios HTML profissionais
- Integracao nativa com Microsoft Outlook para envio de emails
- Diagnostico completo do sistema
- Processamento em lote para multiplos clientes

## Como Usar

### Interface Grafica (Recomendado)

```bash
# Usar o inicializador
iniciar_mmzr.bat

# Ou executar diretamente
python app.py
```

### Operacao da Interface

1. **Selecionar Clientes:** Use a lista com selecao multipla
   - Selecao individual: Clique nos clientes desejados (Ctrl+clique)
   - Selecionar todos: Botao "Selecionar Todos"
   - Limpar selecao: Botao "Limpar Selecao"

2. **Configurar Opcoes:** Marque "Preparar emails no Microsoft Outlook" se desejar envio automatico

3. **Gerar Relatorios:** Clique "Gerar Relatorios" e confirme se multiplos clientes estao selecionados

4. **Acompanhar Progresso:** Use o log do sistema para monitorar o processo

### Linha de Comando 

```bash
# Listar clientes disponiveis
python gerador.py --listar

# Gerar relatorio para cliente especifico
python gerador.py --cliente "Nome do Cliente"

# Gerar com integracao de email
python gerador.py --cliente "Nome do Cliente" --enviar
```

## Instalacao

### Pre-requisitos

- Windows 10/11
- Python 3.8+
- Microsoft Outlook instalado e configurado

### Instalacao das Dependencias

```bash
# Instalar dependencias
pip install -r requirements.txt

# Verificar instalacao
python -c "import pandas, openpyxl; print('Dependencias OK')"
```

### Estrutura de Dados

Colocar as planilhas Excel na pasta `documentos/dados/`:

**Planilha Base:** Deve conter aba "Base Clientes"
- Colunas obrigatorias: Nome cliente, Codigo carteira smart, Nome carteira, Estrategia carteira

**Planilha Rentabilidade:** Dados de performance
- Colunas: Rentabilidade Carteira Mes/Ano, Benchmark Mes/Ano, Variacao Relativa

## Integracao com Email

- Emails abrem automaticamente no Microsoft Outlook como rascunhos
- Pre-preenchimento completo: destinatario, assunto, corpo HTML
- Usuario revisa e envia manualmente
- Suporte a multiplos destinatarios simultaneos

## Diagnostico do Sistema

### Interface Grafica

- **Verificar Status:** Verificacao rapida dos componentes
- **Diagnostico Completo:** Analise detalhada do sistema

### Linha de Comando

```bash
# Verificacao rapida
python diagnostico.py --status

# Diagnostico completo
python diagnostico.py --diagnostico
```

## Arquivos Gerados

**Local:** Pasta principal do projeto  
**Formato:** `relatorio_mensal_NomeCliente_YYYYMMDD.html`  
**Conteudo:** Relatorio HTML profissional com dados de performance  

## Resolucao de Problemas

### Problemas Comuns

**"Nenhum cliente encontrado":**
- Verificar planilhas Excel em `documentos/dados/`
- Executar diagnostico completo na interface
- Verificar estrutura das planilhas

**"Microsoft Outlook nao disponivel":**
- Verificar se Microsoft Outlook esta instalado
- Configurar conta de email no Outlook
- Executar como administrador se necessario
- Instalar dependencia: `pip install pywin32`

**Interface nao abre:**
- Executar `iniciar_mmzr.bat` na pasta raiz do projeto
- Verificar se Python esta instalado corretamente
- Executar: `python app.py` diretamente

**Erro de dependencias:**
- Reinstalar dependencias: `pip install -r requirements.txt`
- Verificar versao do Python (3.8+)

### Logs do Sistema

- Interface grafica: Consultar "Log do Sistema" na aplicacao
- Linha de comando: Verificar saida do terminal
- Arquivos de log: Verificar logs do Python

## Estrutura do Projeto

```
MMZR-Email/
├── app.py                    # Interface grafica principal
├── gerador.py                # Sistema de geracao de relatorios
├── html_generator.py         # Gerador de HTML profissional
├── mmzr_email_sender.py      # Integracao com Outlook
├── diagnostico.py            # Verificacao e diagnostico
├── iniciar_mmzr.bat         # Script de inicializacao Windows
├── requirements.txt          # Dependencias Python
├── config_planilhas.json    # Configuracao do sistema
├── README.md                 # Este arquivo
└── documentos/dados/         # Planilhas Excel de dados
```

## Requisitos do Sistema

- Windows 10/11
- Microsoft Outlook instalado e configurado com conta de email
- Python 3.8 ou superior
- Dependencias instaladas via requirements.txt
- Planilhas atualizadas na pasta documentos/dados/

---

**MMZR Family Office**  
Versao 1.0.0 - Sistema de Relatorios Automatizados 