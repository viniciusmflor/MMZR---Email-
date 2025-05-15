import json
import pandas as pd
from mmzr_email_generator import MMZREmailGenerator, process_and_generate_report

def testar_extracao_dados_reais():
    """Testa a extração de dados das planilhas reais"""
    
    # Caminho das planilhas
    planilha_base = "documentos/dados/Planilha Inteli.xlsm"
    planilha_rentabilidade = "documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx"
    
    # Criando o gerador
    generator = MMZREmailGenerator()
    
    print("=== TESTE DE EXTRAÇÃO DE DADOS REAIS ===")
    
    # Testando a extração da planilha base
    try:
        excel_base = generator.load_excel_data(planilha_base)
        if excel_base:
            print(f"\n1. Planilha base carregada com sucesso.")
            print(f"   Abas disponíveis: {excel_base.sheet_names}")
            
            # Verificando a aba Base Clientes
            if "Base Clientes" in excel_base.sheet_names:
                df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
                print(f"\n2. Aba 'Base Clientes' encontrada.")
                print(f"   Colunas disponíveis: {df_clientes.columns.tolist()}")
                print(f"   Primeiras linhas:")
                print(df_clientes.head(3))
            else:
                print("\n2. ATENÇÃO: Aba 'Base Clientes' não encontrada!")
        else:
            print("\n1. ERRO: Não foi possível carregar a planilha base.")
    except Exception as e:
        print(f"\n1. ERRO ao carregar planilha base: {str(e)}")
    
    # Testando a extração da planilha de rentabilidade
    try:
        excel_rent = generator.load_excel_data(planilha_rentabilidade)
        if excel_rent:
            print(f"\n3. Planilha de rentabilidade carregada com sucesso.")
            print(f"   Abas disponíveis: {excel_rent.sheet_names}")
            
            # Vamos verificar a primeira aba (geralmente contém dados de rentabilidade)
            primeira_aba = excel_rent.sheet_names[0]
            df_rent = pd.read_excel(excel_rent, sheet_name=primeira_aba)
            print(f"\n4. Aba '{primeira_aba}' carregada.")
            print(f"   Colunas disponíveis: {df_rent.columns.tolist()}")
            print(f"   Primeiras linhas:")
            print(df_rent.head(3))
        else:
            print("\n3. ERRO: Não foi possível carregar a planilha de rentabilidade.")
    except Exception as e:
        print(f"\n3. ERRO ao carregar planilha de rentabilidade: {str(e)}")
    
    # Testando a extração de dados específicos usando as funções do gerador
    try:
        # Vamos testar uma aba específica (usando a primeira aba da planilha de rentabilidade)
        if 'excel_rent' in locals() and excel_rent:
            primeira_aba = excel_rent.sheet_names[0]
            df_teste = pd.read_excel(excel_rent, sheet_name=primeira_aba)
            
            print("\n5. Testando extração de dados específicos:")
            
            # Performance
            performance = generator.extract_performance_data(df_teste)
            print(f"\n   5.1 Performance: {performance}")
            
            # Retorno Financeiro
            retorno = generator.extract_financial_return(df_teste)
            print(f"\n   5.2 Retorno Financeiro: {retorno}")
            
            # Estratégias de Destaque
            estrategias = generator.extract_highlight_strategies(df_teste)
            print(f"\n   5.3 Estratégias de Destaque: {estrategias}")
            
            # Ativos Promotores
            promotores = generator.extract_promoter_assets(df_teste)
            print(f"\n   5.4 Ativos Promotores: {promotores}")
            
            # Ativos Detratores
            detratores = generator.extract_detractor_assets(df_teste)
            print(f"\n   5.5 Ativos Detratores: {detratores}")
        else:
            print("\n5. ERRO: Não foi possível testar a extração de dados específicos.")
    except Exception as e:
        print(f"\n5. ERRO ao testar extração de dados específicos: {str(e)}")
    
    print("\n=== TESTE CONCLUÍDO ===")


if __name__ == "__main__":
    testar_extracao_dados_reais() 