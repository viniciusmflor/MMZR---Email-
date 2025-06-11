# MMZR - Guia de Instalação

Guia rápido para configurar o sistema MMZR de relatórios financeiros.

## Pré-requisitos

- **Python 3.8+** instalado
- **Planilhas Excel** com dados dos clientes

## Instalação

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. Configurar planilhas (3 opções)

**Opção A: Detecção Automática** ⭐ Recomendada
```
documentos/dados/
├── suas_planilhas_clientes.xlsx     # Qualquer nome
├── dados_performance.xlsx           # Qualquer nome
```

**Opção B: Configuração Manual**
Editar `config_planilhas.json`:
```json
{
    "auto_detectar": false,
    "planilhas": {
        "planilha_base": "meus_dados.xlsx",
        "planilha_rentabilidade": "performance.xlsx"
    }
}
```

**Opção C: Nomes Padrão**
- Planilha Inteli.xlsm
- Planilha Inteli - dados de rentabilidade.xlsx

### 3. Testar sistema
```bash
# Verificar compatibilidade
python3 mmzr_compatibilidade.py

# Listar clientes
python3 mmzr_integracao_real.py --listar

# Gerar relatório teste
python3 mmzr_integracao_real.py --cliente "Nome Cliente"
```

## Estrutura Necessária

**Planilha Base**: Aba "Base Clientes" (dados dos clientes)
**Planilha Performance**: Dados de rentabilidade das carteiras

## Windows - Envio de Emails

```bash
pip install pywin32
```

**Como funciona**:
- Sistema **cria rascunho** no Outlook automaticamente
- **Abre o email** para você revisar
- Você pode **editar e enviar** manualmente
- **Mais seguro** que envio automático

## Resolução de Problemas

**Erro: "Planilha não encontrada"**
- Verificar arquivos em `documentos/dados/`
- Conferir `config_planilhas.json`

**Erro: "Aba não encontrada"**
- Verificar aba "Base Clientes" na planilha base

**Testar compatibilidade**
```bash
python3 mmzr_compatibilidade.py
```

---

**MMZR Family Office v1.0.0 | Sistema Testado macOS/Windows** 