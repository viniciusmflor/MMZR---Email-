"""
MMZR Family Office - Integração com Planilhas
Sistema integrado de geração de relatórios

Autor: MMZR Family Office
"""

import pandas as pd
from datetime import datetime
from mmzr_email_generator import MMZREmailGenerator
from mmzr_compatibilidade import MMZRCompatibilidade


def gerar_relatorio_integrado(planilha_base=None, planilha_rentabilidade=None, nome_ou_email_cliente=None, enviar_email=False):
    """Gera relatório integrando dados das planilhas."""
    generator = MMZREmailGenerator()
    
    if not planilha_base or not planilha_rentabilidade:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
    
    try:
        # Carregar dados base
        df_clientes = _carregar_dados_clientes(planilha_base)
        df_rentabilidade = _carregar_dados_rentabilidade(planilha_rentabilidade)
        
        # Processar cliente específico ou todos
        clientes_agrupados = _filtrar_clientes(df_clientes, nome_ou_email_cliente)
        
        # Processar cada cliente
        for nome_cliente, carteiras_cliente in clientes_agrupados:
            _processar_cliente(
                nome_cliente, 
                carteiras_cliente, 
                df_rentabilidade, 
                generator, 
                planilha_base, 
                enviar_email
            )
    
    except Exception as e:
        print(f"ERRO: {str(e)}")


def _carregar_dados_clientes(planilha_base):
    """Carrega e processa dados dos clientes."""
    excel_base = pd.ExcelFile(planilha_base)
    
    if "Base Clientes" not in excel_base.sheet_names:
        raise ValueError("Aba 'Base Clientes' não encontrada na planilha base")
    
    df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
    
    # Verificar se coluna existe e tem valores válidos
    if 'Nome cliente' not in df_clientes.columns:
        raise ValueError("Coluna 'Nome cliente' não encontrada na planilha")
    
    # Tratar valores nulos na coluna Nome cliente
    df_clientes['Nome cliente'] = df_clientes['Nome cliente'].fillna('').astype(str).str.strip()
    df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
    df_clientes = df_clientes[df_clientes['Nome cliente'] != '']
    
    # Integrar dados consolidados se disponível
    if "Base Consolidada" in excel_base.sheet_names:
        df_consolidada = pd.read_excel(excel_base, sheet_name="Base Consolidada")
        
        # Verificar e tratar coluna NomeCompletoCliente
        if 'NomeCompletoCliente' in df_consolidada.columns:
            df_consolidada['NomeCompletoCliente'] = df_consolidada['NomeCompletoCliente'].fillna('').astype(str).str.strip()
        
        df_clientes = df_clientes.merge(
            df_consolidada[['NomeCompletoCliente', 'EmailCliente', 'Banker']], 
            left_on='Nome cliente', 
            right_on='NomeCompletoCliente', 
            how='left'
        )
        
        df_clientes['Email cliente'] = df_clientes['EmailCliente']
        df_clientes['Banker Cliente'] = df_clientes['Banker']
        
        # Gerar emails fictícios para clientes sem email
        clientes_sem_email = df_clientes['Email cliente'].isna()
        if clientes_sem_email.any():
            df_clientes.loc[clientes_sem_email, 'Email cliente'] = df_clientes.loc[clientes_sem_email, 'Nome cliente'].apply(
                lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
            )
    else:
        df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
            lambda nome: f"{nome.lower().replace(' ', '.')}@example.com"
        )
        df_clientes['Banker Cliente'] = None
    
    return df_clientes


def _carregar_dados_rentabilidade(planilha_rentabilidade):
    """Carrega dados de rentabilidade."""
    excel_rent = pd.ExcelFile(planilha_rentabilidade)
    primeira_aba = excel_rent.sheet_names[0]
    return pd.read_excel(excel_rent, sheet_name=primeira_aba)


def _filtrar_clientes(df_clientes, nome_ou_email_cliente):
    """Filtra clientes baseado no critério fornecido."""
    if nome_ou_email_cliente:
        nome_ou_email_cliente = nome_ou_email_cliente.strip()
        df_filtrado = df_clientes[
            (df_clientes['Nome cliente'] == nome_ou_email_cliente) | 
            (df_clientes['Email cliente'] == nome_ou_email_cliente)
        ]
        
        if len(df_filtrado) == 0:
            raise ValueError(f"Cliente '{nome_ou_email_cliente}' não encontrado")
        
        return df_filtrado.groupby('Nome cliente')
    
    return df_clientes.groupby('Nome cliente')


def _processar_cliente(nome_cliente, carteiras_cliente, df_rentabilidade, generator, planilha_base, enviar_email):
    """Processa um cliente específico."""
    email_cliente = carteiras_cliente['Email cliente'].iloc[0]
    banker_cliente = carteiras_cliente['Banker Cliente'].iloc[0] if 'Banker Cliente' in carteiras_cliente.columns else None
    
    # Extrair informações dos bankers
    bankers_info = generator.extract_banker_info(planilha_base, banker_cliente)
    
    # Processar carteiras
    portfolios_data = []
    for _, cliente_row in carteiras_cliente.iterrows():
        codigo_carteira = cliente_row['Código carteira smart']
        df_rent_cliente = df_rentabilidade[df_rentabilidade['Código carteira smart'] == codigo_carteira]
        
        if len(df_rent_cliente) > 0:
            portfolio_data = _obter_dados_carteira(cliente_row, df_rent_cliente.iloc[0], generator)
            if portfolio_data:
                portfolios_data.append(portfolio_data)
    
    if portfolios_data:
        _gerar_e_salvar_relatorio(nome_cliente, email_cliente, portfolios_data, bankers_info, generator, enviar_email)
    else:
        print(f"❌ ERRO: Nenhuma carteira encontrada para {nome_cliente}")


def _obter_dados_carteira(dados_cliente, dados_rentabilidade, generator):
    """Processa dados de uma carteira específica."""
    try:
        # Extrair comentários se disponível
        comentarios = None
        if 'Comentários' in dados_cliente and pd.notna(dados_cliente['Comentários']) and dados_cliente['Comentários'] is not None:
            comentarios = str(dados_cliente['Comentários']).strip()
        
        # Criar dados da carteira
        portfolio_data = {
            'name': dados_cliente['Nome carteira'],
            'type': dados_cliente['Estratégia carteira'],
            'comentarios': comentarios,
            'data': {
                'performance': _criar_dados_performance(dados_rentabilidade, generator),
                'retorno_financeiro': dados_rentabilidade['Retorno Financeiro'] if pd.notna(dados_rentabilidade['Retorno Financeiro']) else 0,
                'estrategias_destaque': _extrair_estrategias(dados_rentabilidade),
                'ativos_promotores': _extrair_ativos(dados_rentabilidade, 'Promotor'),
                'ativos_detratores': _extrair_ativos(dados_rentabilidade, 'Detrator')
            }
        }
        
        return portfolio_data
        
    except Exception as e:
        print(f"ERRO ao processar carteira {dados_cliente['Nome carteira']}: {str(e)}")
        return None


def _criar_dados_performance(dados_rentabilidade, generator):
    """Cria dados de performance formatados."""
    return [
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


def _extrair_estrategias(dados_rentabilidade):
    """Extrai estratégias de destaque."""
    estrategias = []
    for i in [1, 2]:
        col = f'Estratégia de Destaque {i}'
        if col in dados_rentabilidade and pd.notna(dados_rentabilidade[col]):
            estrategias.append(dados_rentabilidade[col])
    
    return estrategias if estrategias else ["Sem estratégias de destaque"]


def _extrair_ativos(dados_rentabilidade, tipo):
    """Extrai ativos promotores ou detratores."""
    ativos = []
    for i in [1, 2]:
        col = f'Ativo {tipo} {i}'
        if col in dados_rentabilidade and pd.notna(dados_rentabilidade[col]):
            ativo = dados_rentabilidade[col]
            if str(ativo) != '-':
                ativos.append(str(ativo))
    
    return ativos if ativos else [f"Sem ativos {tipo.lower()}es"]


def _gerar_e_salvar_relatorio(nome_cliente, email_cliente, portfolios_data, bankers_info, generator, enviar_email):
    """Gera e salva o relatório final."""
    data_ref = datetime.now()
    
    # Gerar HTML
    html_content = generator.generate_html_email(nome_cliente, data_ref, portfolios_data, bankers_info)
    
    # Salvar arquivo
    output_file = generator.save_email_to_file(html_content, nome_cliente)
    
    # Exibir resultados
    print(f"✅ Relatório gerado: {output_file}")
    print(f"📧 Cliente: {nome_cliente} ({email_cliente})")
    print(f"🏦 Bankers: {bankers_info['banker_padrao']} e {bankers_info['outro_banker']}")
    print(f"📊 Carteiras: {len(portfolios_data)}")
    
    # Enviar email se solicitado
    if enviar_email:
        try:
            from mmzr_email_sender import enviar_email_outlook
            assunto = generator.generate_email_subject(data_ref)
            
            sucesso = enviar_email_outlook(
                para=email_cliente,
                assunto=assunto,
                corpo_html=html_content
            )
            
            print(f"📧 Email {'preparado' if sucesso else 'com erro'} para {nome_cliente}")
            
        except ImportError:
            print("⚠️  Módulo de envio de email não disponível. Apenas HTML foi gerado.")
    
    print("-" * 60)


def listar_clientes_disponiveis():
    """Lista clientes disponíveis para relatório."""
    try:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        df_clientes = _carregar_dados_clientes(planilha_base)
        df_rentabilidade = _carregar_dados_rentabilidade(planilha_rentabilidade)
        
        # Filtrar clientes com dados de rentabilidade
        codigos_com_rentabilidade = set(df_rentabilidade['Código carteira smart'])
        df_clientes_com_rentabilidade = df_clientes[df_clientes['Código carteira smart'].isin(codigos_com_rentabilidade)]
        
        clientes_por_nome = df_clientes_com_rentabilidade.groupby('Nome cliente')
        
        print("\n=== CLIENTES DISPONÍVEIS ===")
        print(f"{'Nome Cliente':<30} | {'Email':<30} | {'Qtd Carteiras'}")
        print("-" * 80)
        
        for nome, grupo in clientes_por_nome:
            email = grupo['Email cliente'].iloc[0]
            qtd_carteiras = len(grupo)
            print(f"{nome[:30]:<30} | {email[:30]:<30} | {qtd_carteiras}")
        
        print("-" * 80)
        print(f"Total: {len(clientes_por_nome)} clientes disponíveis")
        
        return list(clientes_por_nome.groups.keys())
        
    except Exception as e:
        print(f"ERRO ao listar clientes: {str(e)}")
        return []


# Interface de linha de comando
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--listar":
            listar_clientes_disponiveis()
        elif sys.argv[1] == "--cliente" and len(sys.argv) > 2:
            cliente = sys.argv[2]
            enviar = "--enviar" in sys.argv
            gerar_relatorio_integrado(nome_ou_email_cliente=cliente, enviar_email=enviar)
        else:
            print("Uso: python mmzr_integracao_real.py [--listar | --cliente 'Nome Cliente' [--enviar]]")
    else:
        # Gerar relatórios para todos os clientes
        gerar_relatorio_integrado() 