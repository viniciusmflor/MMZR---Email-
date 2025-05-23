# Integração de Dados Reais no Sistema MMZR Email

Este documento explica como o sistema de geração de emails da MMZR Family Office funciona com dados reais.

## Estrutura de Dados

O sistema utiliza duas planilhas principais:

1. **Planilha Inteli.xlsm**
   - Localização: `documentos/dados/Planilha Inteli.xlsm`
   - Abas utilizadas:
     - `Base Clientes`: Dados principais dos clientes e suas carteiras
     - `Base Consolidada`: Informações adicionais, incluindo dados dos bankers
   - Dados contidos:
     - Nome cliente (identificador principal)
     - Email cliente (identificador alternativo)
     - Código carteira smart
     - Nome carteira
     - Estratégia carteira
     - Benchmark
     - Informações do banker (na aba Base Consolidada)

2. **Planilha Inteli - dados de rentabilidade.xlsx**
   - Localização: `documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx`
   - Aba principal: `Sheet1`
   - Dados contidos:
     - Código carteira smart (chave de ligação com a carteira)
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
2. Permite buscar clientes por nome ou email
3. Agrupa os clientes pelo nome, permitindo identificar clientes com múltiplas carteiras
4. Para cada cliente encontrado, extrai:
   - Informações básicas da planilha base (nome, carteira, estratégia, benchmark)
   - Dados de performance da planilha de rentabilidade (rentabilidades, ativos, etc.)
   - Dados do banker a partir da aba Base Consolidada
5. Gera um relatório HTML personalizado para cada cliente, incluindo todas as suas carteiras em um único relatório
6. Salva os relatórios como arquivos HTML
7. Opcionalmente, envia os relatórios por email (Windows)

## Como Executar a Integração

Para listar os clientes disponíveis:

```bash
python mmzr_integracao_real.py --listar
```

Para gerar um relatório para um cliente específico usando nome ou email:

```bash
python mmzr_integracao_real.py --cliente "Nome do Cliente"
```

ou

```bash
python mmzr_integracao_real.py --cliente "email@cliente.com"
```

Para gerar um relatório e enviá-lo por email (Windows):

```bash
python mmzr_integracao_real.py --cliente "Nome do Cliente" --enviar
```

Para gerar relatórios para todos os clientes:

```bash
python mmzr_integracao_real.py
```

Para criar dados de exemplo para testes:

```bash
python mmzr_integracao_real.py --criar-exemplo
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

5. **Novas funcionalidades**:
   - Link para a carta mensal (gerado automaticamente com base no mês/ano atual)
   - Observação sobre os bankers em cópia (extraído da aba Base Consolidada)
   - Múltiplas carteiras agrupadas no mesmo relatório

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
   - Ajustes específicos para iOS e Android
   - Controle de proporções para diferentes dispositivos

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

- Os emails são gerados no formato HTML moderno, com compatibilidade para clientes de email mais antigos
- O sistema detecta automaticamente o ambiente (macOS/Windows) e adapta seu comportamento
- Os arquivos são salvos localmente mesmo quando enviados por email
- O sistema mostra mensagens detalhadas no terminal para acompanhamento do processo
- Compatibilidade otimizada para tema escuro em dispositivos móveis 
- A identificação do cliente é feita pelo nome ou email, permitindo visualizar todas as carteiras de um mesmo cliente em um único relatório
- Inclusão de observação sobre bankers em cópia, extraindo essa informação da aba Base Consolidada
- Geração automática do link para a carta mensal com base no mês e ano atuais 