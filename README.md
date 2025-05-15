# MMZR Family Office - Gerador de Relatórios de Performance

## Visão Geral

Este projeto implementa um sistema de geração automatizada de relatórios mensais de performance para a MMZR Family Office. O sistema processa dados de planilhas Excel e gera relatórios em formato HTML com visualização moderna e responsiva, melhorando significativamente a experiência do usuário em comparação com o formato anterior.

## Funcionalidades

- Geração de relatórios de performance em formato HTML moderno e responsivo
- Suporte a múltiplas carteiras por cliente
- Visualização de dados relevantes para o cliente:
  - Performance da carteira (mensal, anual e últimos 12 meses)
  - Retorno financeiro
  - Estratégias de destaque
  - Ativos promotores (que contribuíram positivamente)
  - Ativos detratores (que contribuíram negativamente)
- Interface gráfica para configuração e geração de relatórios
- Configuração salva entre sessões

## Estrutura do Projeto

- `mmzr_email_generator.py`: Classe principal para processamento de dados e geração de HTML
- `mmzr_app.py`: Aplicativo com interface gráfica para configuração e geração de relatórios
- `gerar_planilha_teste.py`: Script para gerar planilhas de teste para desenvolvimento
- `testar_extracao_dados.py`: Script para testar a extração de dados das planilhas
- `documentos/`: Pasta com arquivos de dados e recursos visuais
  - `dados/`: Planilhas Excel com dados financeiros reais
  - `img/`: Imagens e logotipos usados nos relatórios
- `planilhas_teste/`: Contém planilhas geradas para teste

## Formato Esperado dos Dados

O gerador espera encontrar nas planilhas Excel os seguintes dados (em qualquer ordem):

1. **Performance**: Tabela com períodos (colunas: Período, Carteira, Benchmark, Diferença)
2. **Retorno Financeiro**: Valor numérico
3. **Estratégias de Destaque**: Lista de estratégias que se destacaram positivamente
4. **Ativos Promotores**: Lista de ativos que contribuíram positivamente
5. **Ativos Detratores**: Lista de ativos que contribuíram negativamente

Se algum desses dados não for encontrado, o sistema usará valores simulados para preencher o relatório.

## Requisitos

- Python 3.6 ou superior
- Bibliotecas Python:
  - pandas
  - numpy
  - tkinter
  - openpyxl (para leitura de arquivos Excel)

## Instalação

1. Clone este repositório ou baixe os arquivos para sua máquina local
2. Instale as dependências necessárias:

```bash
pip install -r requirements.txt
```

## Uso

### Interface Gráfica

Para usar a interface gráfica:

1. Execute o arquivo `mmzr_app.py`:

```bash
python mmzr_app.py
```

2. Na interface:
   - Selecione o arquivo Excel com os dados
   - Preencha as informações do cliente
   - Configure as carteiras conforme necessário
   - Clique em "Gerar Relatório"

### Criação de Planilhas de Teste

Para criar planilhas de teste que sigam o formato esperado:

```bash
python gerar_planilha_teste.py
```

Isso criará um arquivo em `planilhas_teste/dados_teste_mmzr.xlsx` com várias abas para teste.

### Teste da Extração de Dados

Para testar se a extração de dados está funcionando corretamente:

```bash
python testar_extracao_dados.py
```

Este script mostrará os dados extraídos de cada aba das planilhas disponíveis e gerará um relatório de teste.

### Uso via Código

É possível usar diretamente o módulo `mmzr_email_generator.py`:

```python
from mmzr_email_generator import process_and_generate_report

# Configuração do cliente
client = {
    'name': 'Nome do Cliente',
    'email': 'cliente@example.com',
    'portfolios': [
        {
            'name': 'Carteira Moderada',
            'type': 'Renda Variável + Renda Fixa',
            'sheet_name': 'Base Consolidada',
            'benchmark_name': 'IPCA+5%'
        }
    ]
}

# Processar e gerar relatório
result = process_and_generate_report('caminho/para/planilha.xlsx', client)
```

## Customização do Relatório

### Modificando o Estilo Visual

Para personalizar o estilo visual do relatório, edite as funções de geração de HTML na classe `MMZREmailGenerator`:

- `generate_html_email`: Template principal do e-mail
- `generate_portfolio_section`: Seção de cada carteira
- `generate_performance_table`: Tabela de performance
- `generate_financial_return_section`: Seção de retorno financeiro
- `generate_highlight_strategies_section`: Seção de estratégias de destaque
- `generate_promoter_assets_section`: Seção de ativos promotores
- `generate_detractor_assets_section`: Seção de ativos detratores

### Modificando a Extração de Dados

Para ajustar como os dados são extraídos das planilhas, edite os métodos de extração:

- `extract_performance_data`: Extração de dados de performance
- `extract_financial_return`: Extração do retorno financeiro
- `extract_highlight_strategies`: Extração das estratégias de destaque
- `extract_promoter_assets`: Extração dos ativos promotores
- `extract_detractor_assets`: Extração dos ativos detratores

## Notas sobre Adaptação para Dados Reais

O sistema foi adaptado para funcionar tanto com as planilhas simuladas de teste quanto com as planilhas reais fornecidas pelo cliente. As funções de extração são suficientemente robustas para lidar com diferentes formatos de dados, mas podem precisar de ajustes caso a estrutura das planilhas reais seja muito diferente do esperado.

Para garantir a compatibilidade com os dados reais, recomenda-se:

1. Executar `testar_extracao_dados.py` com as planilhas reais
2. Verificar se todos os dados estão sendo extraídos corretamente
3. Ajustar as funções de extração conforme necessário

## Licença

Este projeto é de propriedade da MMZR Family Office e deve ser usado de acordo com os termos acordados. 