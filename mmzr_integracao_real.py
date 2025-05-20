import os
import json
import pandas as pd
from datetime import datetime
from mmzr_email_generator import MMZREmailGenerator, process_and_generate_report
from mmzr_compatibilidade import MMZRCompatibilidade

def gerar_relatorio_integrado(planilha_base=None, planilha_rentabilidade=None, codigo_cliente=None, enviar_email=False):
    """
    Gera um relatório integrando dados das duas planilhas reais:
    - Planilha Base: contém informações dos clientes
    - Planilha de Rentabilidade: contém dados de performance
    
    Args:
        planilha_base: Caminho para a planilha Inteli.xlsm
        planilha_rentabilidade: Caminho para a planilha de rentabilidade
        codigo_cliente: Código específico do cliente (opcional)
        enviar_email: Se True, tenta enviar o email (no Windows) ou simula (no macOS)
    """
    # Iniciando o gerador
    generator = MMZREmailGenerator()
    
    # Obter caminhos das planilhas se não forem fornecidos
    if not planilha_base or not planilha_rentabilidade:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
    
    print("=== INICIANDO INTEGRAÇÃO DAS PLANILHAS REAIS ===")
    print(f"Sistema operacional: {MMZRCompatibilidade.get_os_info()['sistema']}")
    
    # Carregando as planilhas
    try:
        # Carregar planilha base
        excel_base = pd.ExcelFile(planilha_base)
        print(f"Planilha base carregada: {planilha_base}")
        
        # Carregar a aba Base Clientes
        if "Base Clientes" in excel_base.sheet_names:
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
            print(f"Aba Base Clientes carregada, encontrados {len(df_clientes)} clientes")
        else:
            print("ERRO: Aba 'Base Clientes' não encontrada na planilha base")
            return
        
        # Carregar planilha de rentabilidade
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        print(f"Planilha de rentabilidade carregada: {planilha_rentabilidade}")
        
        # Carregar dados de rentabilidade (primeira aba)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        print(f"Dados de rentabilidade carregados, encontrados {len(df_rentabilidade)} registros")
        
        # Filtrar cliente específico se fornecido
        if codigo_cliente:
            df_cliente = df_clientes[df_clientes['Código carteira smart'] == codigo_cliente]
            if len(df_cliente) == 0:
                print(f"ERRO: Cliente com código {codigo_cliente} não encontrado")
                return
            
            df_rent_cliente = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo_cliente]
            if len(df_rent_cliente) == 0:
                print(f"ERRO: Dados de rentabilidade para cliente com código {codigo_cliente} não encontrados")
                return
            
            # Processar um único cliente
            processar_cliente(df_cliente.iloc[0], df_rent_cliente.iloc[0], generator, enviar_email)
        else:
            # Processar todos os clientes que existem em ambas as planilhas
            codigos_clientes = set(df_clientes['Código carteira smart']).intersection(
                set(df_rentabilidade['Código carteira smart']))
            
            print(f"Encontrados {len(codigos_clientes)} clientes em comum entre as planilhas")
            
            for codigo in codigos_clientes:
                cliente = df_clientes[df_clientes['Código carteira smart'] == codigo].iloc[0]
                rent = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo].iloc[0]
                processar_cliente(cliente, rent, generator, enviar_email)
        
    except Exception as e:
        print(f"ERRO durante a integração: {str(e)}")
        import traceback
        traceback.print_exc()

def processar_cliente(dados_cliente, dados_rentabilidade, generator, enviar_email=False):
    """Processa os dados de um cliente e gera o relatório"""
    try:
        nome_cliente = dados_cliente['Nome cliente']
        nome_carteira = dados_cliente['Nome carteira']
        estrategia = dados_cliente['Estratégia carteira']
        benchmark = dados_cliente['Benchmark']
        codigo = dados_cliente['Código carteira smart']
        
        print(f"\nProcessando cliente: {nome_cliente}, Carteira: {nome_carteira} (Código: {codigo})")
        
        # Configuração do cliente
        client_config = {
            'name': nome_cliente,
            'email': f"{nome_cliente.lower().replace(' ', '.')}@example.com",
            'portfolios': [
                {
                    'name': nome_carteira,
                    'type': estrategia,
                    'sheet_name': 'Sheet1',  # Usamos a aba da planilha de rentabilidade
                    'benchmark_name': benchmark
                }
            ]
        }
        
        # Criar DataFrame a partir dos dados de rentabilidade para permitir a extração
        df_rent = pd.DataFrame([dados_rentabilidade])
        
        # Adicionar cabeçalhos para a extração de dados
        df_rent['Performance'] = 'Performance'
        df_rent['Mês atual'] = 'Mês atual'
        df_rent['Ano atual'] = 'Ano atual'
        df_rent['Retorno Financeiro'] = 'Retorno Financeiro'
        df_rent['Estratégias de Destaque'] = 'Estratégias de Destaque'
        df_rent['Ativos Promotores'] = 'Ativos Promotores'
        df_rent['Ativos Detratores'] = 'Ativos Detratores'
        
        # Criar dados de performance
        performance_data = [
            {
                'periodo': f"{generator.meses_pt[datetime.now().month]}:",
                'carteira': dados_rentabilidade['Rentabilidade Carteira Mês'],
                'benchmark': dados_rentabilidade['Benchmark Mês'],
                'diferenca': dados_rentabilidade['Variação Relativa Mês']
            },
            {
                'periodo': "No ano:",
                'carteira': dados_rentabilidade['Rentabilidade Carteira No Ano'],
                'benchmark': dados_rentabilidade['Benchmark No Ano'],
                'diferenca': dados_rentabilidade['Variação Relativa No Ano']
            }
        ]
        
        # Extrair estratégias de destaque
        estrategias = []
        if pd.notna(dados_rentabilidade['Estratégia de Destaque 1']):
            estrategias.append(dados_rentabilidade['Estratégia de Destaque 1'])
        if pd.notna(dados_rentabilidade['Estratégia de Destaque 2']):
            estrategias.append(dados_rentabilidade['Estratégia de Destaque 2'])
        
        # Extrair ativos promotores
        promotores = []
        if pd.notna(dados_rentabilidade['Ativo Promotor 1']):
            promotores.append(dados_rentabilidade['Ativo Promotor 1'])
        if pd.notna(dados_rentabilidade['Ativo Promotor 2']):
            promotores.append(dados_rentabilidade['Ativo Promotor 2'])
        
        # Extrair ativos detratores
        detratores = []
        if pd.notna(dados_rentabilidade['Ativo Detrator 1']):
            detratores.append(dados_rentabilidade['Ativo Detrator 1'])
        if pd.notna(dados_rentabilidade['Ativo Detrator 2']):
            detratores.append(dados_rentabilidade['Ativo Detrator 2'])
        
        # Criar dados da carteira
        portfolio_data = {
            'name': nome_carteira,
            'type': estrategia,
            'data': {
                'performance': performance_data,
                'retorno_financeiro': dados_rentabilidade['Retorno Financeiro'] if pd.notna(dados_rentabilidade['Retorno Financeiro']) else 0,
                'estrategias_destaque': estrategias if estrategias else ["Sem estratégias de destaque"],
                'ativos_promotores': promotores if promotores else ["Sem ativos promotores"],
                'ativos_detratores': detratores if detratores else ["Sem ativos detratores"]
            }
        }
        
        # Gerar o relatório
        portfolios_data = [portfolio_data]
        html_content = generator.generate_html_email(nome_cliente, datetime.now(), portfolios_data)
        
        # Salvar o relatório
        output_file = generator.save_email_to_file(html_content, nome_cliente)
        print(f"Relatório gerado com sucesso: {output_file}")
        
        # Enviar email se solicitado
        if enviar_email:
            email_dest = client_config['email']
            assunto = f"Relatório Mensal MMZR - {generator.meses_pt[datetime.now().month]} de {datetime.now().year}"
            
            enviado = MMZRCompatibilidade.enviar_email(
                destinatario=email_dest, 
                assunto=assunto, 
                caminho_html=output_file
            )
            
            if enviado:
                print(f"Email enviado com sucesso para {email_dest}")
            else:
                print(f"ATENÇÃO: Email não enviado para {email_dest}")
        
        return output_file
        
    except Exception as e:
        print(f"ERRO ao processar cliente {dados_cliente['Nome cliente']}: {str(e)}")
        return None


def listar_clientes_disponiveis():
    """Lista os clientes disponíveis para relatório"""
    try:
        # Obter caminhos das planilhas
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        # Carregar planilha base
        excel_base = pd.ExcelFile(planilha_base)
        df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
        
        # Carregar planilha de rentabilidade
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        
        # Identificar clientes em comum
        codigos_clientes = set(df_clientes['Código carteira smart']).intersection(
            set(df_rentabilidade['Código carteira smart']))
        
        # Preparar lista de clientes
        clientes_disponiveis = []
        print("\n=== CLIENTES DISPONÍVEIS PARA RELATÓRIO ===")
        print(f"{'Código':<10} | {'Nome Cliente':<30} | {'Carteira':<20}")
        print("-" * 70)
        
        for codigo in codigos_clientes:
            cliente = df_clientes[df_clientes['Código carteira smart'] == codigo].iloc[0]
            nome = cliente['Nome cliente']
            carteira = cliente['Nome carteira']
            
            clientes_disponiveis.append({
                'codigo': codigo,
                'nome': nome,
                'carteira': carteira
            })
            
            print(f"{codigo:<10} | {nome[:30]:<30} | {carteira[:20]:<20}")
        
        print("-" * 70)
        print(f"Total: {len(clientes_disponiveis)} clientes disponíveis")
        print("\nPara gerar um relatório específico, use o comando:")
        print("python mmzr_integracao_real.py --cliente [CÓDIGO]")
        
        return clientes_disponiveis
        
    except Exception as e:
        print(f"ERRO ao listar clientes: {str(e)}")
        return []


if __name__ == "__main__":
    import sys
    
    # Verificar compatibilidade
    compat = MMZRCompatibilidade.testar_compatibilidade()
    if not compat['paths_ok']:
        print("ERRO: Os caminhos para as planilhas não estão corretos")
        sys.exit(1)
    
    # Processar argumentos de linha de comando
    if len(sys.argv) > 1:
        # Se for --help, mostrar ajuda
        if sys.argv[1] == "--help" or sys.argv[1] == "-h":
            print("\n=== AJUDA DO MMZR INTEGRAÇÃO REAL ===")
            print("Uso: python mmzr_integracao_real.py [opções]")
            print("\nOpções:")
            print("  --cliente [CÓDIGO]    Gera relatório apenas para o cliente com o código especificado")
            print("  --enviar              Envia o relatório por email (apenas Windows)")
            print("  --listar              Lista todos os clientes disponíveis para relatório")
            print("  --help, -h            Mostra esta ajuda")
            sys.exit(0)
        
        # Se for --listar, listar clientes disponíveis
        if sys.argv[1] == "--listar":
            listar_clientes_disponiveis()
            sys.exit(0)
        
        # Se for --cliente, processar cliente específico
        if sys.argv[1] == "--cliente" and len(sys.argv) > 2:
            codigo_cliente = int(sys.argv[2])
            
            # Verificar se deve enviar email
            enviar_email = "--enviar" in sys.argv
            
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
            gerar_relatorio_integrado(planilha_base, planilha_rentabilidade, codigo_cliente, enviar_email)
            sys.exit(0)
    
    # Por padrão, listar clientes disponíveis
    clientes = listar_clientes_disponiveis()
    
    # Perguntar ao usuário qual cliente processar
    if clientes:
        try:
            codigo = input("\nDigite o código do cliente para gerar o relatório (ou Enter para todos): ")
            enviar = input("Enviar por email? (s/N): ").lower() == 's'
            
            if codigo.strip():
                # Processar um cliente específico
                gerar_relatorio_integrado(codigo_cliente=int(codigo), enviar_email=enviar)
            else:
                # Processar todos os clientes
                gerar_relatorio_integrado(enviar_email=enviar)
        except ValueError:
            print("Código inválido. Gerando para todos os clientes...")
            gerar_relatorio_integrado(enviar_email=False)
    else:
        print("Nenhum cliente disponível para processamento.")
        sys.exit(1) 