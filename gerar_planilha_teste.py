import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

# Garantir que o diretório existe
os.makedirs('planilhas_teste', exist_ok=True)

# Criar um objeto ExcelWriter para salvar múltiplas abas
excel_writer = pd.ExcelWriter('planilhas_teste/dados_teste_mmzr.xlsx', engine='openpyxl')

# Datas para o relatório
data_atual = datetime.now()
data_15dias_atras = data_atual - timedelta(days=15)

# Função para criar dados fictícios de uma carteira
def criar_dados_carteira(nome_carteira, tipo_carteira, benchmark):
    # DataFrame principal
    df = pd.DataFrame()
    
    # Informações básicas
    df.loc[0, 0] = f"MMZR Family Office - Relatório {nome_carteira}"
    df.loc[1, 0] = f"Data de Referência: {data_atual.strftime('%d/%m/%Y')}"
    df.loc[2, 0] = f"Tipo: {tipo_carteira}"
    
    # Seção Performance
    df.loc[5, 0] = "Performance"
    df.loc[6, 0] = "Período"
    df.loc[6, 1] = "Carteira"
    df.loc[6, 2] = benchmark
    df.loc[6, 3] = "Carteira vs. Benchmark"
    
    # Dados de performance
    periodos = ["Últimos 15 dias", "Mês atual", "Ano atual", "Últimos 12 meses"]
    # Gerar valores aleatórios
    carteira_valores = np.round(np.random.uniform(-1.5, 3.5, 4), 2)
    benchmark_valores = np.round(np.random.uniform(-1.0, 2.8, 4), 2)
    diferenca_valores = np.round(carteira_valores - benchmark_valores, 2)
    
    for i, periodo in enumerate(periodos):
        df.loc[7+i, 0] = periodo
        df.loc[7+i, 1] = carteira_valores[i]
        df.loc[7+i, 2] = benchmark_valores[i]
        df.loc[7+i, 3] = diferenca_valores[i]
    
    # Retorno Financeiro
    df.loc[12, 0] = "Retorno Financeiro"
    df.loc[12, 1] = np.round(np.random.uniform(500, 5000, 1)[0], 2)
    
    # Estratégias de Destaque
    df.loc[14, 0] = "Estratégias de Destaque"
    estrategias = [
        f"RETORNO ABSOLUTO ({np.round(np.random.uniform(1, 10, 1)[0], 2)}%) > RENDA VARIÁVEL ({np.round(np.random.uniform(5, 15, 1)[0], 2)}%)",
        f"FUNDOS MULTIMERCADOS ({np.round(np.random.uniform(3, 8, 1)[0], 2)}%) > FUNDOS DE AÇÕES ({np.round(np.random.uniform(5, 12, 1)[0], 2)}%)"
    ]
    df.loc[15, 0] = estrategias[0]
    df.loc[16, 0] = estrategias[1]
    
    # Ativos Promotores
    df.loc[18, 0] = "Ativos Promotores"
    ativos_promotores = [
        f"OCEANA LONG BIASED ADVISORY FIC FI MULT ({np.round(np.random.uniform(5, 15, 1)[0], 2)}%)",
        f"SHARP LB ADVISORY FIC FIA ({np.round(np.random.uniform(7, 18, 1)[0], 2)}%)",
        f"VERDE AM LONG BIAS FIC FIA ({np.round(np.random.uniform(4, 12, 1)[0], 2)}%)"
    ]
    df.loc[19, 0] = ativos_promotores[0]
    df.loc[20, 0] = ativos_promotores[1]
    df.loc[21, 0] = ativos_promotores[2]
    
    # Ativos Detratores
    df.loc[23, 0] = "Ativos Detratores"
    ativos_detratores = [
        f"ALLOCATION GLOBAL EQUITIES FI MULT ({np.round(np.random.uniform(-5, -0.5, 1)[0], 2)}%)",
        f"BTG PACTUAL DISCOVERY FI MULT ({np.round(np.random.uniform(-3, -0.1, 1)[0], 2)}%)"
    ]
    df.loc[24, 0] = ativos_detratores[0]
    df.loc[25, 0] = ativos_detratores[1]
    
    return df

# Criar várias carteiras
carteiras = [
    {"nome": "Carteira Moderada", "tipo": "Renda Variável + Renda Fixa", "benchmark": "IPCA+5%"},
    {"nome": "Carteira Conservadora", "tipo": "Renda Fixa", "benchmark": "CDI"},
    {"nome": "Carteira Arrojada", "tipo": "Renda Variável", "benchmark": "IBOVESPA"}
]

# Criar dados para cada cliente
clientes = ["João Silva", "Maria Oliveira", "Pedro Santos"]

for cliente in clientes:
    # Criar uma planilha para o cliente
    cliente_sheet = pd.DataFrame()
    cliente_sheet.loc[0, 0] = f"Dados do Cliente: {cliente}"
    cliente_sheet.loc[1, 0] = f"Email: {cliente.lower().replace(' ', '.')}@example.com"
    cliente_sheet.loc[2, 0] = f"Data de Referência: {data_atual.strftime('%d/%m/%Y')}"
    
    # Salvar na planilha
    cliente_sheet_name = cliente.replace(' ', '_')
    cliente_sheet.to_excel(excel_writer, sheet_name=cliente_sheet_name, index=False, header=False)

    # Criar carteiras para o cliente
    for carteira in carteiras:
        df_carteira = criar_dados_carteira(
            carteira["nome"], 
            carteira["tipo"],
            carteira["benchmark"]
        )
        
        # Nome da aba: Cliente_TipoCarteira
        sheet_name = f"{cliente_sheet_name}_{carteira['nome'].replace(' ', '_')}"
        if len(sheet_name) > 31:  # Limite de caracteres para nome de aba no Excel
            sheet_name = sheet_name[:31]
            
        # Salvar na planilha
        df_carteira.to_excel(excel_writer, sheet_name=sheet_name, index=False, header=False)

# Criar uma aba de estrutura similar às planilhas reais do cliente
base_consolidada = pd.DataFrame()
base_consolidada.loc[0, 0] = "MMZR Family Office - Relatório Base Consolidada"
base_consolidada.loc[1, 0] = f"Data de Referência: {data_atual.strftime('%d/%m/%Y')}"
base_consolidada.loc[2, 0] = "Tipo: Consolidado"

# Seção Performance
base_consolidada.loc[5, 0] = "Performance"
base_consolidada.loc[6, 0] = "Período"
base_consolidada.loc[6, 1] = "Carteira"
base_consolidada.loc[6, 2] = "IPCA+5%"
base_consolidada.loc[6, 3] = "Carteira vs. Benchmark"

# Dados de performance
periodos = ["Mês atual", "Ano atual", "Últimos 12 meses"]
carteira_valores = [2.38, 8.76, 14.52]
benchmark_valores = [1.45, 5.32, 9.87]
diferenca_valores = [0.93, 3.44, 4.65]

for i, periodo in enumerate(periodos):
    base_consolidada.loc[7+i, 0] = periodo
    base_consolidada.loc[7+i, 1] = carteira_valores[i]
    base_consolidada.loc[7+i, 2] = benchmark_valores[i]
    base_consolidada.loc[7+i, 3] = diferenca_valores[i]

# Retorno Financeiro
base_consolidada.loc[11, 0] = "Retorno Financeiro"
base_consolidada.loc[11, 1] = 1140.27

# Estratégias de Destaque
base_consolidada.loc[13, 0] = "Estratégias de Destaque"
base_consolidada.loc[14, 0] = "RETORNO ABSOLUTO (3,12%) > RENDA VARIÁVEL (8,54%)"
base_consolidada.loc[15, 0] = "OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%) > SHARP LB ADVISORY FIC FIA (12,57%)"

# Ativos Promotores
base_consolidada.loc[17, 0] = "Ativos Promotores"
base_consolidada.loc[18, 0] = "OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%)"
base_consolidada.loc[19, 0] = "SHARP LB ADVISORY FIC FIA (12,57%)"

# Ativos Detratores
base_consolidada.loc[21, 0] = "Ativos Detratores"
base_consolidada.loc[22, 0] = "ALLOCATION GLOBAL EQUITIES FI MULT (-1,66%)"
base_consolidada.loc[23, 0] = "BTG PACTUAL DISCOVERY FI MULT (-0,85%)"

# Salvar na planilha
base_consolidada.to_excel(excel_writer, sheet_name="Base Consolidada", index=False, header=False)

# Salvar o arquivo Excel
excel_writer.close()
print("Planilha de teste criada com sucesso: planilhas_teste/dados_teste_mmzr.xlsx") 