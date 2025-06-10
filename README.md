# MMZR Family Office - Gerador de Relatórios

Sistema de geração automática de relatórios mensais para clientes da MMZR Family Office.

## 📋 Funcionalidades

- ✅ **Geração automática de relatórios HTML** a partir de planilhas Excel
- ✅ **Comentários personalizados** da planilha integrados automaticamente
- ✅ **Envio de emails** via Outlook (Windows) ou simulação (macOS/Linux)
- ✅ **Múltiplas carteiras por cliente** em um único relatório
- ✅ **Interface compatível** com clientes de email
- ✅ **Processamento em lote** ou cliente específico

## 🚀 Instalação

### Requisitos
- Python 3.8+
- pandas
- win32com (apenas Windows, para integração com Outlook)

### Instalação
```bash
pip install -r requirements.txt
```

## 📁 Estrutura de Arquivos

```
MMZR - Email/
├── mmzr_integracao_real.py     # Script principal
├── mmzr_email_generator.py     # Gerador de HTML
├── mmzr_compatibilidade.py     # Compatibilidade multi-OS
├── requirements.txt            # Dependências Python
├── documentos/
│   └── dados/
│       ├── Planilha Inteli.xlsm
│       └── Planilha Inteli - dados de rentabilidade.xlsx
└── recursos_email/
    └── logo-MMZR-azul.png
```

## 💻 Uso

### Comando Básico
```bash
# Listar clientes disponíveis
python mmzr_integracao_real.py --listar

# Gerar relatório para cliente específico
python mmzr_integracao_real.py --cliente "Nome do Cliente"

# Gerar e enviar por email
python mmzr_integracao_real.py --cliente "Nome do Cliente" --enviar

# Modo interativo
python mmzr_integracao_real.py
```

### Opções de Linha de Comando
- `--listar`: Lista todos os clientes disponíveis
- `--cliente "NOME"`: Gera relatório para cliente específico
- `--enviar`: Envia o relatório por email automaticamente
- `--help`: Mostra ajuda

## 📊 Estrutura das Planilhas

### Aba "Base Clientes"
Colunas obrigatórias:
- `Nome cliente`
- `Email cliente` (ou obtido da aba "Base Consolidada")
- `Código carteira smart`
- `Nome carteira`
- `Estratégia carteira`
- `Comentários` (opcional)

### Aba "Base Consolidada"
Para emails e informações de bankers:
- `NomeCompletoCliente`
- `EmailCliente`
- `Banker`
- `NomePronomeBanker`

### Planilha de Rentabilidade
Colunas obrigatórias:
- `Código carteira smart`
- `Rentabilidade Carteira Mês`
- `Benchmark Mês`
- `Variação Relativa Mês`
- `Rentabilidade Carteira No Ano`
- `Benchmark No Ano`
- `Variação Relativa No Ano`
- `Retorno Financeiro`
- `Estratégia de Destaque 1`
- `Estratégia de Destaque 2`
- `Ativo Promotor 1`
- `Ativo Promotor 2`
- `Ativo Detrator 1`
- `Ativo Detrator 2`

## 📧 Funcionalidade de Comentários

Os comentários da coluna "Comentários" na planilha aparecem automaticamente na seção de observações do email no formato:

```
Observações:
• Obs.: Eventuais ajustes retroativos do IPCA...
• Obs.: Conforme solicitado, deixo o Felipe e Renato em cópia...
• Comentário [Nome da Carteira]: [Texto da planilha]
```

## 🖥️ Compatibilidade

### Windows
- Integração completa com Microsoft Outlook
- Envio automático de emails

### macOS/Linux
- Simulação de envio de emails
- Geração completa de relatórios HTML

## 🔧 Configuração

O sistema detecta automaticamente:
- Sistema operacional
- Disponibilidade do Outlook (Windows)
- Caminhos das planilhas

## 📝 Saída

Para cada cliente, o sistema gera:
- Arquivo HTML otimizado para email: `relatorio_mensal_[Cliente]_[Data].html`
- Logo incorporada em base64 para compatibilidade
- Layout responsivo compatível com clientes de email

## 🚨 Solução de Problemas

### Erro: "Cliente não encontrado"
- Verifique se o nome está exatamente como na planilha
- Use aspas ao especificar nomes com espaços

### Erro: "Planilha não encontrada"
- Confirme que as planilhas estão em `documentos/dados/`
- Verifique os nomes dos arquivos

### Problemas de Email (Windows)
- Certifique-se que o Outlook está instalado
- Execute o script como administrador se necessário

## 🔄 Versão

**Versão Final Refatorada** - Código otimizado e simplificado para produção.

## 📞 Suporte

Para suporte técnico, consulte a documentação interna da MMZR Family Office. 