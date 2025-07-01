"""
MMZR Family Office - Sistema de Integração Principal
Processamento integrado de dados de planilhas para geração de relatórios

Autor: MMZR Family Office
Versão: 1.0.0
"""

import pandas as pd
import logging
from datetime import datetime
from typing import List, Tuple, Optional, Dict, Any
from html_generator import MMZREmailGenerator
from diagnostico import MMZRCompatibilidade

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def gerar_relatorio_integrado(
    planilha_base: Optional[str] = None, 
    planilha_rentabilidade: Optional[str] = None, 
    nome_ou_email_cliente: Optional[str] = None, 
    enviar_email: bool = False
) -> bool:
    """
    Gera relatório integrando dados das planilhas.
    
    Args:
        planilha_base: Caminho da planilha base (opcional, detectado automaticamente)
        planilha_rentabilidade: Caminho da planilha de rentabilidade (opcional, detectado automaticamente)
        nome_ou_email_cliente: Nome ou email do cliente específico (opcional, gera para todos se None)
        enviar_email: Se True, prepara email no Outlook
        
    Returns:
        bool: True se sucesso, False se erro
    """
    try:
        logger.info("Iniciando processo de geração de relatórios")
        generator = MMZREmailGenerator()
        
        # Detectar planilhas automaticamente se não fornecidas
        if not planilha_base or not planilha_rentabilidade:
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
            logger.info(f"Planilhas detectadas - Base: {planilha_base}, Rentabilidade: {planilha_rentabilidade}")
        
        # Carregar dados
        df_clientes = _carregar_dados_clientes(planilha_base)
        df_rentabilidade = _carregar_dados_rentabilidade(planilha_rentabilidade)
        
        # Filtrar clientes
        clientes_agrupados = _filtrar_clientes(df_clientes, nome_ou_email_cliente)
        
        # Processar cada cliente
        sucesso_total = True
        for nome_cliente, carteiras_cliente in clientes_agrupados:
            sucesso_cliente = _processar_cliente(
                nome_cliente, 
                carteiras_cliente, 
                df_rentabilidade, 
                generator, 
                planilha_base, 
                enviar_email
            )
            if not sucesso_cliente:
                sucesso_total = False
        
        logger.info("Processo de geração concluído")
        return sucesso_total
        
    except Exception as e:
        error_msg = f"ERRO no processo principal: {str(e)}"
        logger.error(error_msg)
        print(error_msg)
        return False


def _carregar_dados_clientes(planilha_base: str) -> pd.DataFrame:
    """
    Carrega e processa dados dos clientes da planilha base.
    
    Args:
        planilha_base: Caminho da planilha base
        
    Returns:
        pd.DataFrame: DataFrame com dados dos clientes processados
        
    Raises:
        ValueError: Se estrutura da planilha não for válida
    """
    logger.info(f"Carregando dados de clientes de: {planilha_base}")
    
    try:
        excel_base = pd.ExcelFile(planilha_base)
        
        if "Base Clientes" not in excel_base.sheet_names:
            raise ValueError("Aba 'Base Clientes' não encontrada na planilha base")
        
        df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
        
        # Validar e limpar coluna Nome cliente
        if 'Nome cliente' not in df_clientes.columns:
            raise ValueError("Coluna 'Nome cliente' não encontrada na planilha")
        
        df_clientes['Nome cliente'] = df_clientes['Nome cliente'].fillna('').astype(str).str.strip()
        df_clientes = df_clientes[df_clientes['Nome cliente'] != 'Nome Cliente']
        df_clientes = df_clientes[df_clientes['Nome cliente'] != '']
        
        # Integrar dados consolidados se disponível
        if "Base Consolidada" in excel_base.sheet_names:
            logger.info("Integrando dados da aba 'Base Consolidada'")
            df_consolidada = pd.read_excel(excel_base, sheet_name="Base Consolidada")
            
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
                        lambda nome: f"{nome.lower().replace(' ', '.')}@cliente.com"
                    )
        else:
            # Se não há base consolidada, gerar emails fictícios
            df_clientes['Email cliente'] = df_clientes['Nome cliente'].apply(
                lambda nome: f"{nome.lower().replace(' ', '.')}@cliente.com"
            )
            df_clientes['Banker Cliente'] = None
        
        logger.info(f"Dados de clientes carregados: {len(df_clientes)} registros")
        return df_clientes
        
    except Exception as e:
        logger.error(f"Erro ao carregar dados de clientes: {str(e)}")
        raise


def _carregar_dados_rentabilidade(planilha_rentabilidade: str) -> pd.DataFrame:
    """
    Carrega dados de rentabilidade da planilha.
    
    Args:
        planilha_rentabilidade: Caminho da planilha de rentabilidade
        
    Returns:
        pd.DataFrame: DataFrame com dados de rentabilidade
    """
    logger.info(f"Carregando dados de rentabilidade de: {planilha_rentabilidade}")
    
    try:
        excel_rent = pd.ExcelFile(planilha_rentabilidade)
        primeira_aba = excel_rent.sheet_names[0]
        df_rentabilidade = pd.read_excel(excel_rent, sheet_name=primeira_aba)
        
        logger.info(f"Dados de rentabilidade carregados: {len(df_rentabilidade)} registros")
        return df_rentabilidade
        
    except Exception as e:
        logger.error(f"Erro ao carregar dados de rentabilidade: {str(e)}")
        raise


def _filtrar_clientes(df_clientes: pd.DataFrame, nome_ou_email_cliente: Optional[str]):
    """
    Filtra clientes baseado no critério fornecido.
    
    Args:
        df_clientes: DataFrame com dados dos clientes
        nome_ou_email_cliente: Nome ou email do cliente para filtrar (None para todos)
        
    Returns:
        DataFrameGroupBy: Clientes agrupados por nome
        
    Raises:
        ValueError: Se cliente específico não for encontrado
    """
    if nome_ou_email_cliente:
        nome_ou_email_cliente = nome_ou_email_cliente.strip()
        logger.info(f"Filtrando cliente: {nome_ou_email_cliente}")
        
        df_filtrado = df_clientes[
            (df_clientes['Nome cliente'] == nome_ou_email_cliente) | 
            (df_clientes['Email cliente'] == nome_ou_email_cliente)
        ]
        
        if len(df_filtrado) == 0:
            raise ValueError(f"Cliente '{nome_ou_email_cliente}' não encontrado na planilha")
        
        logger.info(f"Cliente encontrado: {len(df_filtrado)} carteira(s)")
        return df_filtrado.groupby('Nome cliente')
    
    logger.info(f"Processando todos os clientes: {len(df_clientes.groupby('Nome cliente'))} cliente(s)")
    return df_clientes.groupby('Nome cliente')


def _processar_cliente(
    nome_cliente: str, 
    carteiras_cliente: pd.DataFrame, 
    df_rentabilidade: pd.DataFrame, 
    generator: MMZREmailGenerator, 
    planilha_base: str, 
    enviar_email: bool
) -> bool:
    """
    Processa um cliente específico gerando seu relatório.
    
    Args:
        nome_cliente: Nome do cliente
        carteiras_cliente: DataFrame com carteiras do cliente
        df_rentabilidade: DataFrame com dados de rentabilidade
        generator: Instância do gerador de emails
        planilha_base: Caminho da planilha base para extrair bankers
        enviar_email: Se deve preparar email
        
    Returns:
        bool: True se sucesso, False se erro
    """
    try:
        logger.info(f"Processando cliente: {nome_cliente}")
        
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
            return _gerar_e_salvar_relatorio(nome_cliente, email_cliente, portfolios_data, bankers_info, generator, enviar_email)
        else:
            error_msg = f"ERRO: Nenhuma carteira com dados de rentabilidade encontrada para {nome_cliente}"
            logger.error(error_msg)
            print(error_msg)
            return False
            
    except Exception as e:
        error_msg = f"ERRO ao processar cliente {nome_cliente}: {str(e)}"
        logger.error(error_msg)
        print(error_msg)
        return False


def _obter_dados_carteira(dados_cliente: pd.Series, dados_rentabilidade: pd.Series, generator: MMZREmailGenerator) -> Optional[Dict[str, Any]]:
    """
    Processa dados de uma carteira específica.
    
    Args:
        dados_cliente: Série com dados do cliente
        dados_rentabilidade: Série com dados de rentabilidade
        generator: Instância do gerador para formatação
        
    Returns:
        Dict[str, Any]: Dados da carteira formatados ou None se erro
    """
    try:
        # Extrair comentários se disponível
        comentarios = None
        if 'Comentários' in dados_cliente and pd.notna(dados_cliente['Comentários']):
            comentarios = str(dados_cliente['Comentários']).strip()
            if comentarios:
                logger.info(f"Comentário encontrado para carteira {dados_cliente['Nome carteira']}")
        
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
        
        logger.info(f"Carteira processada: {dados_cliente['Nome carteira']}")
        return portfolio_data
        
    except Exception as e:
        logger.error(f"ERRO ao processar carteira {dados_cliente['Nome carteira']}: {str(e)}")
        return None


def _criar_dados_performance(dados_rentabilidade: pd.Series, generator: MMZREmailGenerator) -> List[Dict[str, Any]]:
    """
    Cria dados de performance formatados.
    
    Args:
        dados_rentabilidade: Série com dados de rentabilidade
        generator: Instância do gerador para acesso aos meses
        
    Returns:
        List[Dict]: Lista com dados de performance do mês e ano
    """
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


def _extrair_estrategias(dados_rentabilidade: pd.Series) -> List[str]:
    """
    Extrai estratégias de destaque dos dados.
    
    Args:
        dados_rentabilidade: Série com dados de rentabilidade
        
    Returns:
        List[str]: Lista de estratégias de destaque
    """
    estrategias = []
    for i in [1, 2]:
        col = f'Estratégia de Destaque {i}'
        if col in dados_rentabilidade and pd.notna(dados_rentabilidade[col]):
            estrategia = str(dados_rentabilidade[col]).strip()
            if estrategia:
                estrategias.append(estrategia)
    
    return estrategias if estrategias else ["Sem estratégias de destaque disponíveis"]


def _extrair_ativos(dados_rentabilidade: pd.Series, tipo: str) -> List[str]:
    """
    Extrai ativos promotores ou detratores.
    
    Args:
        dados_rentabilidade: Série com dados de rentabilidade
        tipo: 'Promotor' ou 'Detrator'
        
    Returns:
        List[str]: Lista de ativos do tipo especificado
    """
    ativos = []
    for i in [1, 2]:
        col = f'Ativo {tipo} {i}'
        if col in dados_rentabilidade and pd.notna(dados_rentabilidade[col]):
            ativo = str(dados_rentabilidade[col]).strip()
            if ativo and ativo != '-':
                ativos.append(ativo)
    
    return ativos if ativos else [f"Sem ativos {tipo.lower()}es identificados"]


def _gerar_e_salvar_relatorio(
    nome_cliente: str, 
    email_cliente: str, 
    portfolios_data: List[Dict[str, Any]], 
    bankers_info: Dict[str, str], 
    generator: MMZREmailGenerator, 
    enviar_email: bool
) -> bool:
    """
    Gera e salva o relatório final para o cliente.
    
    Args:
        nome_cliente: Nome do cliente
        email_cliente: Email do cliente
        portfolios_data: Dados das carteiras
        bankers_info: Informações dos bankers
        generator: Instância do gerador
        enviar_email: Se deve preparar email
        
    Returns:
        bool: True se sucesso, False se erro
    """
    try:
        data_ref = datetime.now()
        logger.info(f"Gerando relatório para {nome_cliente}")
        
        # Gerar HTML
        html_content = generator.generate_html_email(nome_cliente, data_ref, portfolios_data, bankers_info)
        
        # Salvar arquivo
        output_file = generator.save_email_to_file(html_content, nome_cliente)
        
        # Exibir resultados
        print(f"SUCESSO: Relatório gerado - {output_file}")
        print(f"Cliente: {nome_cliente} ({email_cliente})")
        print(f"Bankers: {bankers_info['banker_padrao']} e {bankers_info['outro_banker']}")
        print(f"Carteiras processadas: {len(portfolios_data)}")
        
        # Preparar email se solicitado
        if enviar_email:
            try:
                from mmzr_email_sender import enviar_email_outlook
                assunto = generator.generate_email_subject(data_ref)
                
                sucesso_email = enviar_email_outlook(
                    para=email_cliente,
                    assunto=assunto,
                    corpo_html=html_content
                )
                
                status_email = "preparado no Outlook" if sucesso_email else "com erro"
                print(f"Email {status_email} para {nome_cliente}")
                
            except ImportError:
                print("AVISO: Módulo de envio de email não disponível. Apenas HTML foi gerado.")
        
        print("-" * 60)
        return True
        
    except Exception as e:
        error_msg = f"ERRO ao gerar relatório para {nome_cliente}: {str(e)}"
        logger.error(error_msg)
        print(error_msg)
        return False


def listar_clientes_disponiveis() -> List[str]:
    """
    Lista todos os clientes disponíveis para geração de relatórios.
    
    Returns:
        List[str]: Lista com nomes dos clientes disponíveis
    """
    try:
        logger.info("Listando clientes disponíveis")
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        df_clientes = _carregar_dados_clientes(planilha_base)
        df_rentabilidade = _carregar_dados_rentabilidade(planilha_rentabilidade)
        
        # Filtrar clientes que possuem dados de rentabilidade
        codigos_com_rentabilidade = set(df_rentabilidade['Código carteira smart'])
        df_clientes_com_rentabilidade = df_clientes[
            df_clientes['Código carteira smart'].isin(codigos_com_rentabilidade)
        ]
        
        clientes_por_nome = df_clientes_com_rentabilidade.groupby('Nome cliente')
        
        print("\n" + "=" * 80)
        print("CLIENTES DISPONÍVEIS PARA GERAÇÃO DE RELATÓRIOS")
        print("=" * 80)
        print(f"{'Nome do Cliente':<40} | {'Email':<30} | {'Carteiras'}")
        print("-" * 80)
        
        clientes_lista = []
        for nome, grupo in clientes_por_nome:
            email = grupo['Email cliente'].iloc[0]
            qtd_carteiras = len(grupo)
            print(f"{nome[:40]:<40} | {email[:30]:<30} | {qtd_carteiras}")
            clientes_lista.append(nome)
        
        print("-" * 80)
        print(f"Total: {len(clientes_lista)} cliente(s) com dados de rentabilidade disponíveis")
        print("=" * 80)
        
        logger.info(f"Listagem concluída: {len(clientes_lista)} clientes encontrados")
        return clientes_lista
        
    except Exception as e:
        error_msg = f"ERRO ao listar clientes: {str(e)}"
        logger.error(error_msg)
        print(error_msg)
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
            sucesso = gerar_relatorio_integrado(nome_ou_email_cliente=cliente, enviar_email=enviar)
            sys.exit(0 if sucesso else 1)
        else:
            print("Uso: python gerador.py [--listar | --cliente 'Nome Cliente' [--enviar]]")
            sys.exit(1)
    else:
        # Gerar relatórios para todos os clientes
        sucesso = gerar_relatorio_integrado()
        sys.exit(0 if sucesso else 1) 