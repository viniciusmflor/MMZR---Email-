import json
from mmzr_email_generator import process_and_generate_report

def testar_geracao_relatorio_real():
    """Testa a geração de um relatório completo usando os dados reais"""
    
    # Caminhos das planilhas
    planilha_real = "documentos/dados/Planilha Inteli - dados de rentabilidade.xlsx"
    
    # Configuração do cliente
    client_config = {
        'name': 'Cliente Real',
        'email': 'cliente.real@example.com',
        'portfolios': [
            {
                'name': 'Carteira Moderada',
                'type': 'Renda Variável + Renda Fixa',
                'sheet_name': 'Sheet1',
                'benchmark_name': 'IPCA+5%'
            }
        ]
    }
    
    print("=== TESTE DE GERAÇÃO DE RELATÓRIO COM DADOS REAIS ===")
    
    try:
        # Processar e gerar relatório
        output_file = process_and_generate_report(planilha_real, client_config)
        
        if output_file:
            print(f"\nRelatório gerado com sucesso em: {output_file}")
        else:
            print("\nERRO: Não foi possível gerar o relatório.")
    except Exception as e:
        print(f"\nERRO ao gerar relatório: {str(e)}")
    
    print("\n=== TESTE CONCLUÍDO ===")


if __name__ == "__main__":
    testar_geracao_relatorio_real() 