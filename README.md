# MMZR Family Office - Gerador de Relatórios

Sistema automatizado para gerar relatórios HTML de performance de carteiras de investimento.

## Instalação

1. Instalar dependências:
```bash
pip install -r requirements.txt
```

2. Colocar as planilhas Excel na pasta:
```
documentos/dados/
├── qualquer_nome_clientes.xlsm     # Planilha com aba "Base Clientes"
└── qualquer_nome_rentabilidade.xlsx # Planilha com dados de performance
```

O sistema detecta automaticamente qual planilha é qual baseado no conteúdo.

## Uso Básico

### Gerar relatório para um cliente específico
```bash
python mmzr_integracao_real.py --cliente "Nome do Cliente"
```

### Ver todos os clientes disponíveis
```bash
python mmzr_integracao_real.py --listar
```

### Gerar relatórios para todos os clientes
```bash
python mmzr_integracao_real.py
```

## Verificar Sistema

### Ver status e planilhas detectadas
```bash
python mmzr_compatibilidade.py --status
```

### Executar diagnóstico completo
```bash
python mmzr_compatibilidade.py --diagnostico
```

## Arquivos Gerados

Os relatórios são salvos como arquivos HTML na pasta principal:
- `relatorio_mensal_Cliente_YYYYMMDD.html`

## Detecção Automática

O sistema identifica as planilhas automaticamente:

**Planilha Base**: Aquela que contém a aba "Base Clientes"
- Deve ter aba "Base Clientes" com dados dos clientes e carteiras
- Pode ter aba "Base Consolidada" com informações dos bankers

**Planilha Rentabilidade**: A outra planilha Excel na pasta
- Contém dados de performance, estratégias e ativos

**Nomes de arquivo**: Podem ser qualquer um (ex: `dados.xlsx`, `clientes.xlsm`, etc.)

## Solução de Problemas

### "Nenhum arquivo Excel encontrado"
Verifique se há arquivos `.xlsx`, `.xlsm` ou `.xls` em `documentos/dados/`

### "Cliente não encontrado"
Verifique se o nome está correto usando `--listar`

### "Aba 'Base Clientes' não encontrada"
Uma das planilhas deve ter a aba "Base Clientes" com dados dos clientes

## Envio por Email

No Windows com Outlook instalado, adicione `--enviar`:
```bash
python mmzr_integracao_real.py --cliente "Nome do Cliente" --enviar
```

Isso criará um rascunho no Outlook para revisão antes do envio.

## Estrutura dos Relatórios

Cada relatório inclui:
- Performance mensal e anual vs benchmark
- Retorno financeiro em reais
- Estratégias de destaque
- Ativos promotores e detratores
- Informações dos bankers responsáveis

## Exemplo de Uso

```bash
# Ver status e planilhas detectadas
python mmzr_compatibilidade.py --status

# Ver clientes disponíveis
python mmzr_integracao_real.py --listar

# Gerar relatório para cliente específico
python mmzr_integracao_real.py --cliente "João Silva"

# Gerar e preparar email (Windows)
python mmzr_integracao_real.py --cliente "João Silva" --enviar
```

O sistema funciona com qualquer nome de planilha, detectando automaticamente qual é qual baseado no conteúdo. 