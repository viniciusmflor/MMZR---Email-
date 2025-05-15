# Integração de Dados Reais no Sistema MMZR Email

Este documento explica como o sistema de geração de emails da MMZR Family Office funciona com dados reais.

## Estrutura de Dados

O sistema utiliza duas planilhas principais:

1. **Planilha Inteli.xlsm**
   - Localização: `documentos/dados/Planilha Inteli.xlsm`
   - Aba principal: `Base Clientes`
   - Dados contidos:
     - Código carteira smart
     - Nome cliente
     - Nome carteira
     - Estratégia carteira
     - Benchmark

2. **Planilha Inteli - dados de rentabilidade.xlsx**
   - Localização: `documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx`
   - Aba principal: `Sheet1`
   - Dados contidos:
     - Código carteira smart (chave de ligação)
     - Rentabilidade Carteira Mês
     - Rentabilidade Carteira No Ano
     - Benchmark Mês
     - Benchmark No Ano
     - Variação Relativa Mês
     - Variação Relativa No Ano
     - Retorno Financeiro
     - Estratégia de Destaque 1 e 2
     - Ativo Promotor 1 e 2
     - Ativo Detrator 1 e 2

## Processo de Integração

O processo de integração é realizado pelo script `mmzr_integracao_real.py`, que:

1. Carrega as duas planilhas
2. Identifica os clientes que existem em ambas as planilhas (usando o código carteira smart como chave)
3. Para cada cliente encontrado, extrai:
   - Informações básicas da planilha base (nome, carteira, estratégia, benchmark)
   - Dados de performance da planilha de rentabilidade (rentabilidades, ativos, etc.)
4. Gera um relatório HTML personalizado para cada cliente
5. Salva os relatórios como arquivos HTML

## Como Executar a Integração

Para gerar relatórios com dados reais:

```python
python mmzr_integracao_real.py
```

Para gerar um relatório para um cliente específico (usando o código da carteira):

```python
# Edite o arquivo mmzr_integracao_real.py e descomente a linha:
# gerar_relatorio_integrado(planilha_base, planilha_rentabilidade, 11719)
# Substitua 11719 pelo código do cliente desejado
```

## Formato do Relatório

O relatório gerado inclui:

1. **Performance da carteira**:
   - Mês atual (com nome do mês)
   - No ano
   - Retorno financeiro (integrado na tabela de performance)

2. **Estratégias de Destaque** (limitadas a 2)

3. **Ativos Promotores** (limitados a 2, com valores positivos)

4. **Ativos Detratores** (limitados a 2, com valores negativos)

## Testes e Verificação

Para verificar se o sistema está funcionando corretamente, você pode usar os seguintes scripts:

- `testar_planilha_real.py`: Testa a extração de dados das planilhas reais
- `testar_relatorio_real.py`: Testa a geração de um relatório usando os dados reais

## Notas Importantes

- O sistema agora usa todas as melhorias visuais implementadas (limite de estratégias, formatação positiva, etc.)
- A extração de dados reais é robusta e detecta automaticamente os clientes presentes em ambas as planilhas
- Os dados de performance são extraídos diretamente sem necessidade de manipulação adicional 