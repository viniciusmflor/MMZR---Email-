# MMZR Family Office - Sistema de Relatórios

Sistema automatizado para geração de relatórios HTML de performance de carteiras de investimento.

**Versão:** 1.0.0  
**Autor:** MMZR Family Office  

## Como Usar

### Para não técnicos - estagiários (Recomendado)
1. Duplo clique em `iniciar_mmzr.bat`
2. Selecione o cliente na lista
3. Clique em "Gerar Relatório"

### Para Técnicos (manipulação pelo terminal)
```bash
# Listar clientes
python gerador.py --listar

# Gerar relatório
python gerador.py --cliente "Nome do Cliente"

# Com email
python gerador.py --cliente "Nome do Cliente" --enviar
```

## Instalação

1. Instalar Python 3.8+
2. Instalar dependências:
```bash
pip install -r requirements.txt
```
3. Colocar planilhas Excel em `documentos/dados/`

## Estrutura de Planilhas

**Planilha Base:** Deve ter aba "Base Clientes"
- Colunas: Nome cliente, Código carteira smart, Nome carteira, Estratégia carteira

**Planilha Rentabilidade:** Dados de performance
- Rentabilidade Carteira (Mês/Ano), Benchmark (Mês/Ano), Variação Relativa

## Verificação do Sistema

```bash
# Status rápido
python diagnostico.py --status

# Diagnóstico completo
python diagnostico.py --diagnostico
```

## Arquivos Gerados

- **Local:** Pasta principal do projeto
- **Formato:** `relatorio_mensal_NomeCliente_YYYYMMDD.html`

## Resolução de Problemas

**Erro "Nenhum cliente encontrado":**
- Verificar arquivos Excel em `documentos/dados/`
- Executar `python diagnostico.py --diagnostico`

**Interface não abre:**
- Executar `iniciar_mmzr.bat` na pasta correta do projeto

**Python não encontrado:**
- Instalar Python 3.8+ e marcar "Add to PATH"

## Estrutura do Projeto

```
MMZR-Email/
├── app.py                 # Interface gráfica principal
├── gerador.py             # Sistema de geração de relatórios
├── html_generator.py      # Gerador de HTML
├── diagnostico.py         # Verificação do sistema
├── iniciar_mmzr.bat      # Script de inicialização
└── documentos/dados/      # Planilhas Excel
```

Para informações detalhadas, consulte `MANUAL.md`.

---

**MMZR Family Office | Desenvolvido para Windows** 