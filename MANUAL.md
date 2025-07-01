# MMZR Family Office - Manual Técnico

Documentação técnica detalhada do sistema de relatórios.

## Instalação Detalhada

### Requisitos
- Python 3.8+
- Microsoft Windows (para integração Outlook)
- Conexão com internet para dependências

### Instalação do Python
1. Download: [python.org/downloads](https://www.python.org/downloads/)
2. **Importante:** Marcar "Add Python to PATH" durante instalação
3. Reiniciar computador após instalação

### Dependências
```bash
pip install -r requirements.txt
```

**Bibliotecas instaladas:**
- pandas >= 2.0.0 (manipulação de dados)
- openpyxl >= 3.1.0 (arquivos Excel)
- python-dateutil >= 2.8.0 (datas)
- pywin32 >= 306 (Windows/Outlook)

### Verificação da Instalação
```bash
python diagnostico.py --status
python gerador.py --listar
```

## Estrutura Detalhada das Planilhas

### Planilha Base (Clientes)

**Aba "Base Clientes" (obrigatória):**
- `Nome cliente` - Nome completo
- `Código carteira smart` - ID da carteira
- `Nome carteira` - Nome descritivo
- `Estratégia carteira` - Tipo da estratégia
- `Comentários` (opcional) - Observações

**Aba "Base Consolidada" (opcional):**
- `NomeCompletoCliente` - Para correspondência
- `EmailCliente` - Email do cliente
- `Banker` - Código do banker
- `NomePronomeBanker` - Nome do banker

### Planilha Rentabilidade

**Dados por linha (uma carteira por linha):**

Performance:
- `Rentabilidade Carteira Mês`
- `Rentabilidade Carteira No Ano`
- `Benchmark Mês`
- `Benchmark No Ano`
- `Variação Relativa Mês`
- `Variação Relativa No Ano`

Dados adicionais:
- `Retorno Financeiro` (em reais)
- `Estratégia de Destaque 1`
- `Estratégia de Destaque 2`
- `Ativo Promotor 1`
- `Ativo Promotor 2`
- `Ativo Detrator 1`
- `Ativo Detrator 2`

## Linha de Comando

### Comandos Principais
```bash
# Listar todos os clientes
python gerador.py --listar

# Gerar relatório específico
python gerador.py --cliente "Nome do Cliente"

# Gerar com preparação de email
python gerador.py --cliente "Nome do Cliente" --enviar

# Gerar para todos os clientes
python gerador.py

# Verificar sistema
python diagnostico.py --status

# Diagnóstico completo
python diagnostico.py --diagnostico
```

### Códigos de Saída
- 0: Sucesso
- 1: Erro (verificar logs)

## Configuração de Email (Windows)

### Funcionamento
1. Sistema gera HTML
2. Cria rascunho no Outlook
3. Abre para revisão manual
4. Usuário edita e envia

### Vantagens
- Controle total sobre envio
- Revisão antes de enviar
- Compliance e segurança

## Arquivos de Configuração

### config_planilhas.json
```json
{
    "auto_detectar": true,
    "planilhas": {
        "planilha_base": "",
        "planilha_rentabilidade": ""
    }
}
```

## Resolução de Problemas Técnicos

### Logs do Sistema
- **Local:** Pasta raiz do projeto
- **Formato:** Timestamp + Nível + Mensagem
- **Comando:** Consultar área de status na interface

### Problemas Comuns

**ImportError/ModuleNotFoundError:**
```bash
pip install --upgrade -r requirements.txt
```

**"Planilha não encontrada":**
1. Verificar `documentos/dados/`
2. Confirmar extensões (.xlsx, .xlsm, .xls)
3. Executar `python diagnostico.py --diagnostico`

**"Python não é reconhecido":**
1. Reinstalar Python marcando "Add to PATH"
2. Reiniciar terminal/computador

**Interface não responde:**
1. Aguardar processamento
2. Verificar logs na área de status
3. Reiniciar aplicação se necessário

### Debugging
```bash
# Logs detalhados
python diagnostico.py --diagnostico

# Teste de cliente específico
python gerador.py --cliente "ClienteTeste" 2>&1 | tee debug.log
```

## Estrutura Técnica

### Arquitetura
- `app.py` - Interface gráfica (tkinter)
- `gerador.py` - Processamento principal
- `html_generator.py` - Geração de HTML
- `diagnostico.py` - Verificações do sistema

### Dependências entre Módulos
```
app.py
├── gerador.py
│   ├── html_generator.py
│   └── diagnostico.py
└── diagnostico.py
```

### Threads
- Interface principal: Thread UI
- Processamento: Thread separada (não bloqueia UI)
- Carregamento: Thread assíncrona

## Personalização

### Modificar Templates
- Editar `html_generator.py`
- Função `generate_html_email()`

### Adicionar Validações
- Editar `diagnostico.py`
- Função `verificar_estrutura_dados()`

### Configurar Planilhas
- Editar `config_planilhas.json`
- Ou usar detecção automática

---

**MMZR Family Office | Manual Técnico v1.0.0** 