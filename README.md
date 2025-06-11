# MMZR Family Office - Sistema de Relatórios Financeiros

Sistema automatizado para geração de relatórios de performance financeira em HTML para clientes da MMZR Family Office.

## Características Principais

- **Flexibilidade Total**: Aceita qualquer nome de planilha Excel
- **Detecção Automática**: Identifica planilhas automaticamente
- **Configuração Personalizada**: Suporte a nomes específicos de arquivos
- **Compatibilidade**: Funciona em macOS e Windows
- **Relatórios Profissionais**: Gera HTMLs completos com gráficos e indicadores

## Instalação Rápida

```bash
# 1. Instalar dependências
pip install -r requirements.txt

# 2. Colocar suas planilhas na pasta
documentos/dados/
├── sua_planilha_clientes.xlsx
├── seus_dados_performance.xlsx

# 3. Executar o sistema
python3 mmzr_email_generator.py
```

## Planilhas Necessárias

O sistema pode trabalhar com **qualquer nome de planilha Excel**. Você tem 3 opções:

### Opção 1: Detecção Automática (Recomendada)
Simplesmente coloque suas planilhas Excel na pasta `documentos/dados/` e o sistema detectará automaticamente:

```
documentos/dados/
├── minha_planilha_clientes.xlsx     # Será detectada como planilha base
├── dados_rentabilidade.xlsx        # Será detectada como planilha de performance
```

### Opção 2: Configuração Personalizada
Edite o arquivo `config_planilhas.json` na raiz do projeto:

```json
{
    "auto_detectar": false,
    "planilhas": {
        "planilha_base": "meus_clientes.xlsx",
        "planilha_rentabilidade": "performance_carteiras.xlsx"
    }
}
```

### Opção 3: Nomes Padrão (Compatibilidade)
Use os nomes originais (mantém compatibilidade):
- `Planilha Inteli.xlsm`
- `Planilha Inteli - dados de rentabilidade.xlsx`

### Estrutura Necessária das Planilhas

**Planilha Base** (qualquer nome) deve conter:
- Aba "Base Clientes": dados dos clientes e carteiras
- Aba "Base Consolidada": emails dos clientes (opcional)

**Planilha Rentabilidade** (qualquer nome) deve conter:
- Aba com dados de performance das carteiras
- Estratégias, ativos promotores/detratores
- Retorno financeiro

## Funcionalidades do Relatório

### Seção Principal
- Logo MMZR personalizada
- Período de análise
- Performance da carteira vs benchmarks
- Gráfico de evolução patrimonial

### Observações
- Dados do IPCA automáticos
- Comentários de Felipe e Fernandito
- Observações personalizadas da planilha

### Principais Indicadores
- CDI, Selic, Ibovespa
- Dólar, Euro, Bitcoin
- IFIX, S&P 500, NASDAQ

### Recursos Adicionais
- Botão para carta mensal
- Estratégias detalhadas
- Ativos promotores e detratores
- Compatibilidade total macOS/Windows

### Envio de Email (Windows)
- **Cria rascunho no Outlook** para revisão manual
- **Abre automaticamente** para o usuário verificar
- **Seguro**: não envia automaticamente
- **Prático**: usuário pode editar antes de enviar

## Estrutura do Projeto

```
MMZR - Email/
├── mmzr_email_generator.py      # Sistema principal
├── mmzr_compatibilidade.py      # Compatibilidade macOS/Windows
├── mmzr_integracao_real.py      # Integração com APIs
├── config_planilhas.json        # Configuração de planilhas
├── requirements.txt             # Dependências Python
├── documentos/
│   └── dados/                   # Suas planilhas Excel aqui
└── exemplo_uso_planilhas.md     # Exemplos práticos
```

## Teste de Compatibilidade

Para verificar se o sistema funcionará no seu ambiente:

```bash
python3 mmzr_compatibilidade.py
```

## Versão e Suporte

- **Versão**: 1.0.0
- **Python**: 3.8+
- **Sistemas**: macOS, Windows
- **Formato de saída**: HTML responsivo
- **Desenvolvido para**: MMZR Family Office

---

Para mais exemplos práticos, consulte `exemplo_uso_planilhas.md`. 