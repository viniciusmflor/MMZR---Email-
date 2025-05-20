# MMZR Family Office - Gerador de Relatórios de Performance

## Visão Geral

Este projeto implementa um sistema de geração automatizada de relatórios mensais de performance para a MMZR Family Office. O sistema processa dados de planilhas Excel, gera relatórios em formato HTML com visualização moderna e responsiva, e oferece a possibilidade de enviar esses relatórios automaticamente por email quando executado no Windows.

## Funcionalidades

- Geração de relatórios de performance em formato HTML moderno e responsivo
- Compatibilidade entre macOS (desenvolvimento) e Windows (produção)
- Integração nativa com Outlook no Windows para envio automático de emails
- Suporte a múltiplas carteiras por cliente
- Visualização de dados relevantes para o cliente:
  - Performance da carteira (mensal e anual)
  - Retorno financeiro integrado à tabela de performance
  - Estratégias de destaque (limitadas a 2)
  - Ativos promotores com indicação positiva (limitados a 2)
  - Ativos detratores (limitados a 2)
- Compatibilidade com modo escuro em dispositivos móveis

## Estrutura do Projeto

- `mmzr_integracao_real.py`: Script principal para integração com dados reais e envio de emails
- `mmzr_email_generator.py`: Classe principal para processamento de dados e geração de HTML
- `mmzr_compatibilidade.py`: Gerenciamento de compatibilidade entre macOS e Windows
- `documentos/`: Pasta com arquivos de dados e recursos visuais
  - `dados/`: Planilhas Excel com dados financeiros reais
  - `img/`: Imagens e logotipos usados nos relatórios

## Formato das Planilhas de Dados

O sistema utiliza duas planilhas principais:

1. **Planilha Inteli.xlsm**
   - Aba principal: `Base Clientes`
   - Dados: informações básicas dos clientes (código, nome, carteira, estratégia, benchmark)

2. **Planilha Inteli - dados de rentabilidade.xlsx**
   - Contém os dados de performance de cada cliente identificados pelo código da carteira

Se algum dado não for encontrado, o sistema usará valores simulados para preencher o relatório.

## Requisitos

- Python 3.6 ou superior
- Bibliotecas Python:
  - pandas
  - numpy
  - openpyxl (para leitura de arquivos Excel)
- Para envio de emails no Windows:
  - pywin32 (win32com)

## Instalação

1. Clone este repositório ou baixe os arquivos para sua máquina local
2. Instale as dependências necessárias:

```bash
pip install -r requirements.txt
```

3. No Windows, para envio de emails, instale pywin32:

```bash
pip install pywin32
```

## Uso

### Listar Clientes Disponíveis

Para listar todos os clientes disponíveis para geração de relatório:

```bash
python mmzr_integracao_real.py --listar
```

### Gerar Relatório para um Cliente Específico

Para gerar um relatório para um cliente específico usando seu código:

```bash
python mmzr_integracao_real.py --cliente [CÓDIGO]
```

### Gerar e Enviar Relatório (Windows)

Para gerar e enviar por email (somente no Windows):

```bash
python mmzr_integracao_real.py --cliente [CÓDIGO] --enviar
```

### Gerar Relatórios para Todos os Clientes

Para gerar relatórios para todos os clientes:

```bash
python mmzr_integracao_real.py
```

Quando solicitado, você pode escolher se deseja enviar os emails (funciona apenas no Windows).

### Verificar Compatibilidade do Sistema

Para verificar se o sistema está corretamente configurado:

```bash
python mmzr_compatibilidade.py
```

## Customização do Relatório

O design do relatório usa a cor #0D2035, correspondente à identidade visual da MMZR.

Para personalizar o estilo visual do relatório, edite as funções de geração de HTML na classe `MMZREmailGenerator`:

- `generate_html_email`: Template principal do e-mail
- `generate_portfolio_section`: Seção de cada carteira
- `generate_performance_table`: Tabela de performance
- `generate_financial_return_section`: Seção de retorno financeiro
- `generate_highlight_strategies_section`: Seção de estratégias de destaque
- `generate_promoter_assets_section`: Seção de ativos promotores
- `generate_detractor_assets_section`: Seção de ativos detratores

## Compatibilidade com Dispositivos Móveis

O sistema inclui otimizações para garantir que os emails sejam exibidos corretamente em dispositivos móveis, incluindo:

- Design responsivo que se adapta a diferentes tamanhos de tela
- Metadados específicos para forçar o modo claro em dispositivos com tema escuro
- Estilos CSS inline para máxima compatibilidade com clientes de email

## Integração com Outlook (Windows)

No Windows, o sistema utiliza a biblioteca win32com para se integrar diretamente com o Microsoft Outlook:

1. Cria um novo email usando a API do Outlook
2. Define o destinatário, assunto e corpo HTML
3. Insere o conteúdo HTML completo do relatório
4. Envia o email automaticamente através da conta configurada no Outlook

No macOS, o sistema simula o envio de email para fins de desenvolvimento e teste.

## Licença

Este projeto é de propriedade da MMZR Family Office e deve ser usado de acordo com os termos acordados. 