# import os
import json
import pandas as pd
from datetime import datetime
from mmzr_email_generator import MMZREmailGenerator
from mmzr_compatibilidade import MMZRCompatibilidade

def gerar_relatorio_integrado(planilha_base=None, planilha_rentabilidade=None, nome_ou_email_cliente=None, enviar_email=False):
    """
    Gera um relatório integrando dados das duas planilhas reais:
    - Planilha Base: contém informações dos clientes
    - Planilha de Rentabilidade: contém dados de performance
    
    Args:
        planilha_base: Caminho para a planilha Inteli.xlsm
        planilha_rentabilidade: Caminho para a planilha de rentabilidade
        nome_ou_email_cliente: Nome ou email específico do cliente (opcional)
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
            
            # Limpar espaços em branco nos nomes dos clientes
            df_clientes['Nome cliente'] = df_clientes['Nome cliente'].str.strip()
            
            # Filtrar apenas clientes reais (remover dados de template como "Nome Cliente")
            df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
            
            # Se a coluna 'Email cliente' não existir, criar uma com emails fictícios
            if 'Email cliente' not in df_clientes.columns:
                df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
                    lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
                )
            
            print(f"Aba Base Clientes carregada, encontrados {len(df_clientes)} clientes (após filtrar dados de template)")
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
        if nome_ou_email_cliente:
            # Limpar espaços em branco do termo de busca fornecido
            nome_ou_email_cliente = nome_ou_email_cliente.strip()
            
            # Filtrar por nome ou email do cliente
            df_cliente = df_clientes[
                (df_clientes['Nome cliente'] == nome_ou_email_cliente) | 
                (df_clientes['Email cliente'] == nome_ou_email_cliente)
            ]
            
            if len(df_cliente) == 0:
                print(f"ERRO: Cliente com nome ou email '{nome_ou_email_cliente}' não encontrado")
                return
            
            # Agrupar carteiras por nome de cliente
            clientes_agrupados = df_cliente.groupby('Nome cliente')
            
            # Processar cada cliente
            for nome_cliente, carteiras_cliente in clientes_agrupados:
                email_cliente = carteiras_cliente['Email cliente'].iloc[0]
                print(f"\nProcessando cliente: {nome_cliente}, Email: {email_cliente} (Total de carteiras: {len(carteiras_cliente)})")
                
                # Lista para armazenar os dados de todas as carteiras do cliente
                portfolios_data = []
                
                # Para cada carteira do cliente, buscar os dados de rentabilidade
                for _, cliente_row in carteiras_cliente.iterrows():
                    codigo_carteira = cliente_row['Código carteira smart']
                    df_rent_cliente = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo_carteira]
                    
                    if len(df_rent_cliente) == 0:
                        print(f"ERRO: Dados de rentabilidade para carteira com código {codigo_carteira} não encontrados")
                        continue
                    
                    # Obter dados da carteira
                    portfolio_data = obter_dados_carteira(cliente_row, df_rent_cliente.iloc[0], generator)
                    if portfolio_data:
                        portfolios_data.append(portfolio_data)
                
                # Gerar um único relatório com todas as carteiras do cliente
                if portfolios_data:
                    # Gerar o relatório HTML
                    html_content = generator.generate_html_email(nome_cliente, datetime.now(), portfolios_data)
                    
                    # Salvar o relatório
                    output_file = generator.save_email_to_file(html_content, nome_cliente)
                    print(f"Relatório com {len(portfolios_data)} carteiras gerado com sucesso: {output_file}")
                    
                    # Enviar email se solicitado
                    if enviar_email:
                        assunto = generator.generate_email_subject(datetime.now())
                        
                        enviado = MMZRCompatibilidade.enviar_email(
                            destinatario=email_cliente, 
                            assunto=assunto, 
                            caminho_html=output_file
                        )
                        
                        if enviado:
                            print(f"Email enviado com sucesso para {email_cliente}")
                        else:
                            print(f"ATENÇÃO: Email não enviado para {email_cliente}")
        else:
            # Agrupar os clientes pelo nome para identificar aqueles com múltiplas carteiras
            clientes_por_nome = df_clientes.groupby('Nome cliente')
            
            for nome, grupo in clientes_por_nome:
                email_cliente = grupo['Email cliente'].iloc[0]
                print(f"\nProcessando cliente: {nome}, Email: {email_cliente} (Total de carteiras: {len(grupo)})")
                
                # Lista para armazenar os dados de todas as carteiras do cliente
                portfolios_data = []
                
                # Para cada carteira do cliente, buscar os dados de rentabilidade correspondentes
                for _, cliente_row in grupo.iterrows():
                    codigo_carteira = cliente_row['Código carteira smart']
                    df_rent_cliente = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo_carteira]
                    
                    if len(df_rent_cliente) > 0:
                        # Obter dados da carteira
                        portfolio_data = obter_dados_carteira(cliente_row, df_rent_cliente.iloc[0], generator)
                        if portfolio_data:
                            portfolios_data.append(portfolio_data)
                    else:
                        print(f"AVISO: Dados de rentabilidade para carteira com código {codigo_carteira} não encontrados")
                
                # Gerar um único relatório com todas as carteiras do cliente
                if portfolios_data:
                    # Gerar o relatório HTML
                    html_content = generator.generate_html_email(nome, datetime.now(), portfolios_data)
                    
                    # Salvar o relatório
                    output_file = generator.save_email_to_file(html_content, nome)
                    print(f"Relatório com {len(portfolios_data)} carteiras gerado com sucesso: {output_file}")
                    
                    # Enviar email se solicitado
                    if enviar_email:
                        assunto = generator.generate_email_subject(datetime.now())
                        
                        enviado = MMZRCompatibilidade.enviar_email(
                            destinatario=email_cliente, 
                            assunto=assunto, 
                            caminho_html=output_file
                        )
                        
                        if enviado:
                            print(f"Email enviado com sucesso para {email_cliente}")
                        else:
                            print(f"ATENÇÃO: Email não enviado para {email_cliente}")
        
    except Exception as e:
        print(f"ERRO durante a integração: {str(e)}")
        import traceback
        traceback.print_exc()

def obter_dados_carteira(dados_cliente, dados_rentabilidade, generator):
    """Processa os dados de uma carteira e retorna os dados formatados"""
    try:
        nome_carteira = dados_cliente['Nome carteira']
        estrategia = dados_cliente['Estratégia carteira']
        benchmark = dados_cliente['Benchmark']
        codigo = dados_cliente['Código carteira smart']
        
        print(f"  - Processando carteira: {nome_carteira} (Código: {codigo})")
        
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
        
        return portfolio_data
        
    except Exception as e:
        print(f"ERRO ao processar carteira {dados_cliente['Nome carteira']}: {str(e)}")
        return None

def listar_clientes_disponiveis():
    """Lista os clientes disponíveis para relatório"""
    try:
        # Obter caminhos das planilhas
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        # Carregar planilha base
        excel_base = pd.ExcelFile(planilha_base)
        df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
        
        # Limpar espaços extras nos nomes dos clientes
        df_clientes['Nome cliente'] = df_clientes['Nome cliente'].str.strip()
        
        # Filtrar apenas clientes reais (remover dados de template como "Nome Cliente")
        df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
        
        # Se a coluna 'Email cliente' não existir, criar uma com emails fictícios
        if 'Email cliente' not in df_clientes.columns:
            df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
                lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
            )
        
        # Carregar planilha de rentabilidade
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        
        # Mostrar os códigos disponíveis para debug
        print("\nCódigos de carteira disponíveis na planilha de rentabilidade:")
        for codigo in df_rentabilidade['Código carteira smart'].unique():
            print(f"- {codigo}")
        
        # Identificar clientes com dados de rentabilidade disponíveis
        codigos_com_rentabilidade = set(df_rentabilidade['Código carteira smart'])
        df_clientes_com_rentabilidade = df_clientes[df_clientes['Código carteira smart'].isin(codigos_com_rentabilidade)]
        
        print(f"Total de clientes com rentabilidade: {len(df_clientes_com_rentabilidade)}")
        
        # Agrupar clientes pelo nome para identificar aqueles com múltiplas carteiras
        clientes_por_nome = df_clientes_com_rentabilidade.groupby('Nome cliente')
        
        # Preparar lista de clientes
        clientes_disponiveis = []
        print("\n=== CLIENTES DISPONÍVEIS PARA RELATÓRIO ===")
        print(f"{'Nome Cliente':<30} | {'Email':<30} | {'Qtd Carteiras':<15} | {'Carteiras'}")
        print("-" * 110)
        
        for nome, grupo in clientes_por_nome:
            carteiras = grupo['Nome carteira'].tolist()
            carteiras_str = ", ".join(carteiras)
            email = grupo['Email cliente'].iloc[0]
            
            clientes_disponiveis.append({
                'nome': nome,
                'email': email,
                'carteiras': carteiras,
                'qtd_carteiras': len(carteiras)
            })
            
            print(f"{nome[:30]:<30} | {email[:30]:<30} | {len(carteiras):<15} | {carteiras_str}")
        
        print("-" * 110)
        print(f"Total: {len(clientes_disponiveis)} clientes disponíveis")
        print("\nPara gerar um relatório específico, use o comando:")
        print("python mmzr_integracao_real.py --cliente \"[NOME OU EMAIL DO CLIENTE]\"")
        
        return clientes_disponiveis
        
    except Exception as e:
        print(f"ERRO ao listar clientes: {str(e)}")
        return []

def criar_dados_exemplo():
    """Cria dados de exemplo mais completos com emails únicos de clientes"""
    try:
        # Obter caminhos das planilhas
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        # Verificar se os diretórios existem
        base_dir = os.path.dirname(planilha_base)
        if not os.path.exists(base_dir):
            os.makedirs(base_dir, exist_ok=True)
            print(f"Diretório de dados criado: {base_dir}")
        
        # Dados de exemplo para clientes
        clientes = [
            {
                'Nome cliente': 'Carlos Almeida',
                'Email cliente': 'carlos.almeida@exemplo.com.br',
                'Código carteira smart': 'CA001',
                'Nome carteira': 'Carteira Conservadora',
                'Estratégia carteira': 'Renda Fixa',
                'Benchmark': 'CDI'
            },
            {
                'Nome cliente': 'Carlos Almeida',
                'Email cliente': 'carlos.almeida@exemplo.com.br',
                'Código carteira smart': 'CA002',
                'Nome carteira': 'Carteira Moderada',
                'Estratégia carteira': 'Renda Fixa + Renda Variável',
                'Benchmark': 'IPCA+5%'
            },
            {
                'Nome cliente': 'Maria Silva',
                'Email cliente': 'maria.silva@exemplo.com.br',
                'Código carteira smart': 'MS001',
                'Nome carteira': 'Carteira Agressiva',
                'Estratégia carteira': 'Renda Variável',
                'Benchmark': 'Ibovespa'
            },
            {
                'Nome cliente': 'Pedro Santos',
                'Email cliente': 'pedro.santos@exemplo.com.br',
                'Código carteira smart': 'PS001',
                'Nome carteira': 'Carteira Internacional',
                'Estratégia carteira': 'Renda Variável Internacional',
                'Benchmark': 'S&P 500'
            }
        ]
        
        # Dados de exemplo para rentabilidade
        rentabilidade = []
        for cliente in clientes:
            # Gerar dados de rentabilidade aleatórios
            import random
            
            # Valores para o mês
            rent_carteira_mes = round(random.uniform(-3.0, 8.0), 2)
            bench_mes = round(random.uniform(0.5, 3.0), 2)
            var_rel_mes = round(rent_carteira_mes - bench_mes, 2)
            
            # Valores para o ano
            rent_carteira_ano = round(random.uniform(-5.0, 20.0), 2)
            bench_ano = round(random.uniform(3.0, 10.0), 2)
            var_rel_ano = round(rent_carteira_ano - bench_ano, 2)
            
            # Retorno financeiro (valor em reais)
            retorno_financeiro = round(random.uniform(-10000, 50000), 2)
            
            # Estratégias de destaque
            estrategias = [
                "Alocação em títulos públicos", "Diversificação em FIIs",
                "Alocação em ações dividendos", "Proteção com derivativos",
                "Exposição à dólar", "ETFs internacionais"
            ]
            
            # Ativos promotores (sempre com rentabilidade positiva)
            promotores = [
                f"PETR4 (+{round(random.uniform(3.0, 15.0), 2)}%)",
                f"VALE3 (+{round(random.uniform(3.0, 15.0), 2)}%)",
                f"WEGE3 (+{round(random.uniform(3.0, 15.0), 2)}%)",
                f"BBDC4 (+{round(random.uniform(3.0, 15.0), 2)}%)"
            ]
            
            # Ativos detratores (sempre com rentabilidade negativa)
            detratores = [
                f"MGLU3 (-{round(random.uniform(3.0, 15.0), 2)}%)",
                f"IRBR3 (-{round(random.uniform(3.0, 15.0), 2)}%)",
                f"COGN3 (-{round(random.uniform(3.0, 15.0), 2)}%)",
                f"BPAC11 (-{round(random.uniform(3.0, 15.0), 2)}%)"
            ]
            
            # Escolher estratégias, promotores e detratores aleatoriamente
            random.shuffle(estrategias)
            random.shuffle(promotores)
            random.shuffle(detratores)
            
            rentabilidade.append({
                'Código carteira smart': cliente['Código carteira smart'],
                'Rentabilidade Carteira Mês': rent_carteira_mes,
                'Benchmark Mês': bench_mes,
                'Variação Relativa Mês': var_rel_mes,
                'Rentabilidade Carteira No Ano': rent_carteira_ano,
                'Benchmark No Ano': bench_ano,
                'Variação Relativa No Ano': var_rel_ano,
                'Retorno Financeiro': retorno_financeiro,
                'Estratégia de Destaque 1': estrategias[0],
                'Estratégia de Destaque 2': estrategias[1],
                'Ativo Promotor 1': promotores[0],
                'Ativo Promotor 2': promotores[1],
                'Ativo Detrator 1': detratores[0],
                'Ativo Detrator 2': detratores[1]
            })
        
        # Criar DataFrames
        df_clientes = pd.DataFrame(clientes)
        df_rentabilidade = pd.DataFrame(rentabilidade)
        
        # Salvar as planilhas
        df_clientes.to_excel(planilha_base, sheet_name="Base Clientes", index=False)
        df_rentabilidade.to_excel(planilha_rentabilidade, sheet_name="Sheet1", index=False)
        
        print("\n=== DADOS DE EXEMPLO CRIADOS COM SUCESSO ===")
        print(f"Planilha base salva em: {planilha_base}")
        print(f"Planilha de rentabilidade salva em: {planilha_rentabilidade}")
        print(f"Total de clientes: {len(df_clientes['Nome cliente'].unique())}")
        print(f"Total de carteiras: {len(df_clientes)}")
        
        return True
    except Exception as e:
        print(f"ERRO ao criar dados de exemplo: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    import sys
    
    # Verificar compatibilidade
    compat = MMZRCompatibilidade.testar_compatibilidade()
    
    # Processar argumentos de linha de comando
    if len(sys.argv) > 1:
        # Se for --help, mostrar ajuda
        if sys.argv[1] == "--help" or sys.argv[1] == "-h":
            print("\n=== AJUDA DO MMZR INTEGRAÇÃO REAL ===")
            print("Uso: python mmzr_integracao_real.py [opções]")
            print("\nOpções:")
            print("  --cliente \"[NOME OU EMAIL]\"  Gera relatório apenas para o cliente com o nome ou email especificado")
            print("  --enviar                    Envia o relatório por email (apenas Windows)")
            print("  --listar                    Lista todos os clientes disponíveis para relatório")
            print("  --criar-exemplo             Cria dados de exemplo com emails para testes")
            print("  --help, -h                  Mostra esta ajuda")
            sys.exit(0)
        
        # Se for --criar-exemplo, criar dados de exemplo
        if sys.argv[1] == "--criar-exemplo":
            criar_dados_exemplo()
            sys.exit(0)
        
        # Se for --listar, listar clientes disponíveis
        if sys.argv[1] == "--listar":
            listar_clientes_disponiveis()
            sys.exit(0)
        
        # Se for --cliente, processar cliente específico
        if sys.argv[1] == "--cliente" and len(sys.argv) > 2:
            nome_ou_email_cliente = sys.argv[2]
            
            # Verificar se deve enviar email
            enviar_email = "--enviar" in sys.argv
            
            if not compat['paths_ok']:
                print("AVISO: Caminhos das planilhas não encontrados. Criando dados de exemplo...")
                criar_dados_exemplo()
            
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
            gerar_relatorio_integrado(planilha_base, planilha_rentabilidade, nome_ou_email_cliente, enviar_email)
            sys.exit(0)
    
    # Se os caminhos das planilhas não estiverem corretos, sugerir criar dados de exemplo
    if not compat['paths_ok']:
        print("AVISO: Os caminhos para as planilhas não estão corretos.")
        resposta = input("Deseja criar dados de exemplo para testes? (s/N): ").lower()
        if resposta == 's':
            criar_dados_exemplo()
    
    # Por padrão, listar clientes disponíveis
    clientes = listar_clientes_disponiveis()
    
    # Perguntar ao usuário qual cliente processar
    if clientes:
        try:
            nome_ou_email = input("\nDigite o nome ou email do cliente para gerar o relatório (ou Enter para todos): ")
            enviar = input("Enviar por email? (s/N): ").lower() == 's'
            
            if nome_ou_email.strip():
                # Processar um cliente específico
                gerar_relatorio_integrado(nome_ou_email_cliente=nome_ou_email, enviar_email=enviar)
            else:
                # Processar todos os clientes
                gerar_relatorio_integrado(enviar_email=enviar)
        except ValueError:
            print("Nome ou email inválido. Gerando para todos os clientes...")
            gerar_relatorio_integrado(enviar_email=False)
    else:
        print("Nenhum cliente disponível para processamento.")
        resposta = input("Deseja criar dados de exemplo para testes? (s/N): ").lower()
        if resposta == 's':
            criar_dados_exemplo()
            # Após criar dados de exemplo, listar clientes novamente
            clientes = listar_clientes_disponiveis()
            if clientes:
                nome_ou_email = input("\nDigite o nome ou email do cliente para gerar o relatório (ou Enter para todos): ")
                enviar = input("Enviar por email? (s/N): ").lower() == 's'
                
                if nome_ou_email.strip():
                    # Processar um cliente específico
                    gerar_relatorio_integrado(nome_ou_email_cliente=nome_ou_email, enviar_email=enviar)
                else:
                    # Processar todos os clientes
                    gerar_relatorio_integrado(enviar_email=enviar)
            else:
                print("Não foi possível criar clientes de exemplo.")
                sys.exit(1)
        else:
            sys.exit(1) 