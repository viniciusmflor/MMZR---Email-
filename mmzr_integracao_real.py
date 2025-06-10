# import os
import json
import pandas as pd
from datetime import datetime
from mmzr_email_generator import MMZREmailGenerator
from mmzr_compatibilidade import MMZRCompatibilidade

def gerar_relatorio_integrado(planilha_base=None, planilha_rentabilidade=None, nome_ou_email_cliente=None, enviar_email=False):
    """Gera relatório integrando dados das planilhas"""
    generator = MMZREmailGenerator()
    
    if not planilha_base or not planilha_rentabilidade:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
    
    try:
        # Carregar planilha base
        excel_base = pd.ExcelFile(planilha_base)
        
        # Carregar a aba Base Clientes
        if "Base Clientes" in excel_base.sheet_names:
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
            df_clientes['Nome cliente'] = df_clientes['Nome cliente'].str.strip()
            df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
            
            # Carregar a aba Base Consolidada para obter os emails
            if "Base Consolidada" in excel_base.sheet_names:
                df_consolidada = pd.read_excel(excel_base, sheet_name="Base Consolidada")
                df_consolidada['NomeCompletoCliente'] = df_consolidada['NomeCompletoCliente'].str.strip()
                
                df_clientes = df_clientes.merge(
                    df_consolidada[['NomeCompletoCliente', 'EmailCliente']], 
                    left_on='Nome cliente', 
                    right_on='NomeCompletoCliente', 
                    how='left'
                )
                
                df_clientes['Email cliente'] = df_clientes['EmailCliente']
                
                # Criar emails fictícios para clientes sem email
                clientes_sem_email = df_clientes['Email cliente'].isna()
                if clientes_sem_email.any():
                    df_clientes.loc[clientes_sem_email, 'Email cliente'] = df_clientes.loc[clientes_sem_email, 'Nome cliente'].apply(
                        lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
                    )
            else:
                df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
                    lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
                )
        else:
            print("ERRO: Aba 'Base Clientes' não encontrada na planilha base")
            return
        
        # Carregar planilha de rentabilidade
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        
        # Processar cliente específico ou todos
        if nome_ou_email_cliente:
            nome_ou_email_cliente = nome_ou_email_cliente.strip()
            df_cliente = df_clientes[
                (df_clientes['Nome cliente'] == nome_ou_email_cliente) | 
                (df_clientes['Email cliente'] == nome_ou_email_cliente)
            ]
            
            if len(df_cliente) == 0:
                print(f"ERRO: Cliente '{nome_ou_email_cliente}' não encontrado")
                return
            
            clientes_agrupados = df_cliente.groupby('Nome cliente')
        else:
            clientes_agrupados = df_clientes.groupby('Nome cliente')
        
        # Processar cada cliente
        for nome_cliente, carteiras_cliente in clientes_agrupados:
            email_cliente = carteiras_cliente['Email cliente'].iloc[0]
            portfolios_data = []
            
            # Processar cada carteira do cliente
            for _, cliente_row in carteiras_cliente.iterrows():
                codigo_carteira = cliente_row['Código carteira smart']
                df_rent_cliente = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo_carteira]
                
                if len(df_rent_cliente) == 0:
                    continue
                
                portfolio_data = obter_dados_carteira(cliente_row, df_rent_cliente.iloc[0], generator)
                if portfolio_data:
                    portfolios_data.append(portfolio_data)
            
            # Gerar relatório se há dados
            if portfolios_data:
                html_content = generator.generate_html_email(nome_cliente, datetime.now(), portfolios_data)
                output_file = generator.save_email_to_file(html_content, nome_cliente)
                print(f"Relatório gerado: {output_file}")
                
                # Enviar email se solicitado
                if enviar_email:
                    assunto = generator.generate_email_subject(datetime.now())
                    enviado = MMZRCompatibilidade.enviar_email(
                        destinatario=email_cliente, 
                        assunto=assunto, 
                        caminho_html=output_file
                    )
                    
                    if enviado:
                        print(f"Email enviado para {email_cliente}")
        
    except Exception as e:
        print(f"ERRO: {str(e)}")

def obter_dados_carteira(dados_cliente, dados_rentabilidade, generator):
    """Processa os dados de uma carteira e retorna os dados formatados"""
    try:
        nome_carteira = dados_cliente['Nome carteira']
        estrategia = dados_cliente['Estratégia carteira']
        codigo = dados_cliente['Código carteira smart']
        
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
        
        # Extrair comentários da planilha
        comentarios_cliente = None
        if 'Comentários' in dados_cliente:
            comentarios_raw = dados_cliente['Comentários']
            if pd.notna(comentarios_raw) and str(comentarios_raw).strip():
                comentarios_cliente = str(comentarios_raw).strip()
        
        # Criar dados da carteira
        portfolio_data = {
            'name': nome_carteira,
            'type': estrategia,
            'data': {
                'performance': performance_data,
                'retorno_financeiro': dados_rentabilidade['Retorno Financeiro'] if pd.notna(dados_rentabilidade['Retorno Financeiro']) else 0,
                'estrategias_destaque': estrategias if estrategias else ["Sem estratégias de destaque"],
                'ativos_promotores': promotores if promotores else ["Sem ativos promotores"],
                'ativos_detratores': detratores if detratores else ["Sem ativos detratores"],
                'comentarios': comentarios_cliente
            }
        }
        
        return portfolio_data
        
    except Exception as e:
        print(f"ERRO ao processar carteira {dados_cliente['Nome carteira']}: {str(e)}")
        return None

def listar_clientes_disponiveis():
    """Lista os clientes disponíveis para relatório"""
    try:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        excel_base = pd.ExcelFile(planilha_base)
        df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
        df_clientes['Nome cliente'] = df_clientes['Nome cliente'].str.strip()
        df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
        
        if "Base Consolidada" in excel_base.sheet_names:
            df_consolidada = pd.read_excel(excel_base, sheet_name="Base Consolidada")
            df_consolidada['NomeCompletoCliente'] = df_consolidada['NomeCompletoCliente'].str.strip()
            
            df_clientes = df_clientes.merge(
                df_consolidada[['NomeCompletoCliente', 'EmailCliente']], 
                left_on='Nome cliente', 
                right_on='NomeCompletoCliente', 
                how='left'
            )
            
            df_clientes['Email cliente'] = df_clientes['EmailCliente']
            
            clientes_sem_email = df_clientes['Email cliente'].isna()
            if clientes_sem_email.any():
                df_clientes.loc[clientes_sem_email, 'Email cliente'] = df_clientes.loc[clientes_sem_email, 'Nome cliente'].apply(
                    lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
                )
        else:
            df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
                lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
            )
        
        # Carregar planilha de rentabilidade
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        
        # Identificar clientes com dados de rentabilidade disponíveis
        codigos_com_rentabilidade = set(df_rentabilidade['Código carteira smart'])
        df_clientes_com_rentabilidade = df_clientes[df_clientes['Código carteira smart'].isin(codigos_com_rentabilidade)]
        
        clientes_por_nome = df_clientes_com_rentabilidade.groupby('Nome cliente')
        
        print("\n=== CLIENTES DISPONÍVEIS ===")
        print(f"{'Nome Cliente':<30} | {'Email':<30} | {'Qtd Carteiras'}")
        print("-" * 80)
        
        for nome, grupo in clientes_por_nome:
            carteiras = grupo['Nome carteira'].tolist()
            email = grupo['Email cliente'].iloc[0]
            print(f"{nome[:30]:<30} | {email[:30]:<30} | {len(carteiras)}")
        
        print("-" * 80)
        print(f"Total: {len(clientes_por_nome)} clientes disponíveis")
        
        return list(clientes_por_nome.groups.keys())
        
    except Exception as e:
        print(f"ERRO ao listar clientes: {str(e)}")
        return []

if __name__ == "__main__":
    import sys
    
    # Verificar compatibilidade
    compat = MMZRCompatibilidade.testar_compatibilidade()
    
    # Processar argumentos de linha de comando
    if len(sys.argv) > 1:
        if sys.argv[1] == "--help" or sys.argv[1] == "-h":
            print("\n=== MMZR GERADOR DE RELATÓRIOS ===")
            print("Uso: python mmzr_integracao_real.py [opções]")
            print("\nOpções:")
            print("  --cliente \"[NOME OU EMAIL]\"  Gera relatório para cliente específico")
            print("  --enviar                    Envia o relatório por email")
            print("  --listar                    Lista clientes disponíveis")
            print("  --help, -h                  Mostra esta ajuda")
            sys.exit(0)
        
        if sys.argv[1] == "--listar":
            listar_clientes_disponiveis()
            sys.exit(0)
        
        if sys.argv[1] == "--cliente" and len(sys.argv) > 2:
            nome_ou_email_cliente = sys.argv[2]
            enviar_email = "--enviar" in sys.argv
            
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
            gerar_relatorio_integrado(planilha_base, planilha_rentabilidade, nome_ou_email_cliente, enviar_email)
            sys.exit(0)
    
    # Por padrão, listar clientes disponíveis
    clientes = listar_clientes_disponiveis()
    
    if clientes:
        try:
            nome_ou_email = input("\nDigite o nome ou email do cliente (ou Enter para todos): ")
            enviar = input("Enviar por email? (s/N): ").lower() == 's'
            
            if nome_ou_email.strip():
                gerar_relatorio_integrado(nome_ou_email_cliente=nome_ou_email, enviar_email=enviar)
            else:
                gerar_relatorio_integrado(enviar_email=enviar)
        except KeyboardInterrupt:
            print("\nOperação cancelada.")
            sys.exit(1)
    else:
        print("Nenhum cliente disponível para processamento.")
        sys.exit(1) 