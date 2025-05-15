import json
from mmzr_email_generator import process_and_generate_report

# Carregar configuração
with open('mmzr_config.json', 'r') as f:
    client_config = json.load(f)

# Gerar relatório
output_file = process_and_generate_report('planilhas_teste/dados_teste_mmzr.xlsx', client_config)
print(f"Relatório gerado: {output_file}") 