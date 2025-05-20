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

## Compatibilidade entre Plataformas

O sistema foi projetado para funcionar em dois ambientes:

1. **macOS**: Para desenvolvimento e testes
   - Simula o envio de emails (mostra mensagens no terminal)
   - Detecta automaticamente que está no macOS

2. **Windows**: Para produção
   - Integra-se automaticamente com o Microsoft Outlook
   - Envia emails diretamente usando win32com
   - Requer a biblioteca pywin32 instalada

A classe `MMZRCompatibilidade` gerencia essas diferenças entre plataformas.

## Processo de Integração

O processo de integração é realizado pelo script `mmzr_integracao_real.py`, que:

1. Carrega as duas planilhas
2. Identifica os clientes que existem em ambas as planilhas (usando o código carteira smart como chave)
3. Para cada cliente encontrado, extrai:
   - Informações básicas da planilha base (nome, carteira, estratégia, benchmark)
   - Dados de performance da planilha de rentabilidade (rentabilidades, ativos, etc.)
4. Gera um relatório HTML personalizado para cada cliente
5. Salva os relatórios como arquivos HTML
6. Opcionalmente, envia os relatórios por email (Windows)

## Como Executar a Integração

Para listar os clientes disponíveis:

```bash
python mmzr_integracao_real.py --listar
```

Para gerar um relatório para um cliente específico:

```bash
python mmzr_integracao_real.py --cliente [CÓDIGO]
```

Para gerar um relatório e enviá-lo por email (Windows):

```bash
python mmzr_integracao_real.py --cliente [CÓDIGO] --enviar
```

Para gerar relatórios para todos os clientes:

```bash
python mmzr_integracao_real.py
```

## Formato do Relatório

O relatório gerado inclui:

1. **Performance da carteira**:
   - Mês atual (com nome do mês)
   - No ano
   - Retorno financeiro (integrado na tabela de performance)

2. **Estratégias de Destaque** (limitadas a 2)

3. **Ativos Promotores** (limitados a 2, com valores positivos e símbolo "+")

4. **Ativos Detratores** (limitados a 2, com valores negativos)

## Identidade Visual

O relatório segue a identidade visual da MMZR Family Office:

- Cor principal: #0D2035 (azul escuro)
- Logo centralizada e em destaque
- Tipografia adequada para leitura em diversos dispositivos

## Compatibilidade com Dispositivos Móveis

O sistema inclui suporte específico para visualização em dispositivos móveis:

1. **Design Responsivo** que se adapta a diferentes tamanhos de tela

2. **Compatibilidade com Modo Escuro**:
   - Metatags especiais: `<meta name="color-scheme" content="light">`
   - Estilos CSS específicos para forçar modo claro em tela escura
   - Classes CSS para sobrescrever estilos impostos pelo sistema

3. **Otimizações para Clientes de Email Mobile**:
   - Estilos inline para máxima compatibilidade
   - Estrutura de tabelas aninhadas para suporte amplo
   - Cores explícitas para elementos críticos

## Verificação do Sistema

Para verificar se o sistema está configurado corretamente:

```bash
python mmzr_compatibilidade.py
```

Este comando testará:
- O sistema operacional em uso
- Acesso às planilhas necessárias
- Disponibilidade do win32com (em Windows)

## Notas Importantes

- Os emails são gerados no formato HTML moderno, mas com compatibilidade para clientes de email mais antigos
- O sistema detecta automaticamente o ambiente (macOS/Windows) e adapta seu comportamento
- Os arquivos são salvos localmente mesmo quando enviados por email
- O sistema mostra mensagens detalhadas no terminal para acompanhamento do processo
- Agora com suporte completo a tema escuro em dispositivos móveis 