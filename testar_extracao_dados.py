import os
import pandas as pd
from mmzr_email_generator import MMZREmailGenerator
from datetime import datetime

def testar_extracao(planilha_path=None):
    """Testa a extração de dados da planilha"""
    
    # Caminho para a planilha de teste
    if not planilha_path:
        planilha_teste = 'planilhas_teste/dados_teste_mmzr.xlsx'
        
        # Verificar se existe a planilha de teste
        if not os.path.exists(planilha_teste):
            print(f"Erro: A planilha {planilha_teste} não foi encontrada.")
            return
    else:
        planilha_teste = planilha_path
        
    # Criar o gerador de email
    generator = MMZREmailGenerator()
    
    # Carregar a planilha
    excel_file = generator.load_excel_data(planilha_teste)
    if not excel_file:
        print("Erro ao carregar o arquivo Excel.")
        return
    
    # Exibir as abas disponíveis
    print(f"Abas disponíveis: {excel_file.sheet_names}")
    
    # Testar extração de dados de cada aba
    for sheet_name in excel_file.sheet_names:
        print(f"\n===== Testando extração da aba: {sheet_name} =====")
        
        # Ler os dados da aba
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Extrair dados de performance
        performance_data = generator.extract_performance_data(df)
        print(f"\nDados de Performance extraídos ({len(performance_data)} entradas):")
        for item in performance_data:
            print(f"  - {item['periodo']}: Carteira: {item['carteira']}%, "
                  f"Benchmark: {item['benchmark']}%, "
                  f"Diferença: {item['diferenca']}pp")
        
        # Extrair retorno financeiro
        retorno_financeiro = generator.extract_financial_return(df)
        print(f"\nRetorno Financeiro extraído: {generator.format_currency(retorno_financeiro)}")
        
        # Extrair estratégias de destaque
        estrategias_destaque = generator.extract_highlight_strategies(df)
        print(f"\nEstratégias de Destaque extraídas ({len(estrategias_destaque)} entradas):")
        for estrategia in estrategias_destaque:
            print(f"  - {estrategia}")
        
        # Extrair ativos promotores
        ativos_promotores = generator.extract_promoter_assets(df)
        print(f"\nAtivos Promotores extraídos ({len(ativos_promotores)} entradas):")
        for ativo in ativos_promotores:
            print(f"  - {ativo}")
        
        # Extrair ativos detratores
        ativos_detratores = generator.extract_detractor_assets(df)
        print(f"\nAtivos Detratores extraídos ({len(ativos_detratores)} entradas):")
        for ativo in ativos_detratores:
            print(f"  - {ativo}")
        
        print("\n")  # Linha em branco para separar as abas

def testar_com_dados_reais():
    """Testa a extração de dados das planilhas reais"""
    dados_reais = [
        "documentos/dados/Planilha Inteli.xlsm",
        "documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx"
    ]
    
    for planilha in dados_reais:
        if os.path.exists(planilha):
            print(f"\n===== Testando com dados reais: {planilha} =====")
            testar_extracao(planilha)
        else:
            print(f"Arquivo não encontrado: {planilha}")

def gerar_relatório_teste():
    """Gera um relatório de teste usando dados reais ou simulados"""
    # Verificar se os dados reais estão disponíveis
    dados_reais = "documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx"
    planilha_teste = "planilhas_teste/dados_teste_mmzr.xlsx"
    
    planilha_a_usar = dados_reais if os.path.exists(dados_reais) else planilha_teste
    
    print(f"Gerando relatório usando: {planilha_a_usar}")
    
    # Configuração do cliente
    client_config = {
        'name': 'João Silva',
        'email': 'joao.silva@example.com',
        'portfolios': [
            {
                'name': 'Carteira Moderada',
                'type': 'Renda Variável + Renda Fixa',
                'sheet_name': 'João_Silva_Carteira_Moderada' if "teste" in planilha_a_usar else 'Base Consolidada',
                'benchmark_name': 'IPCA+5%'
            }
        ]
    }
    
    # Processar e gerar relatório
    from mmzr_email_generator import process_and_generate_report
    result = process_and_generate_report(planilha_a_usar, client_config)
    
    if result:
        print(f"Relatório gerado com sucesso: {result}")
    else:
        print("Erro ao gerar relatório.")

if __name__ == "__main__":
    print("===== Testando extração de dados das planilhas =====")
    testar_extracao()
    
    print("\n\n===== Testando extração de dados das planilhas reais =====")
    testar_com_dados_reais()
    
    print("\n\n===== Gerando relatório de teste =====")
    gerar_relatório_teste() 