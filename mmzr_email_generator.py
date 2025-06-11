"""
MMZR Family Office - Gerador de Relatórios de Performance

Este módulo implementa a classe MMZREmailGenerator responsável por processar dados
financeiros de planilhas Excel e gerar relatórios HTML personalizados para clientes.

Autor: MMZR Family Office
Versão: 2.0.0
Data: 2025-01-11
"""

import os
import logging
from typing import Dict, List, Optional, Any, Union
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class MMZREmailGenerator:
    """
    Gerador de emails HTML para MMZR Family Office.
    
    Esta classe processa dados financeiros de planilhas Excel e gera relatórios
    HTML personalizados com performance de carteiras, estratégias e ativos.
    
    Attributes:
        meses_pt (Dict[int, str]): Mapeamento de números dos meses para nomes em português
    """
    
    def __init__(self) -> None:
        """Inicializa o gerador de emails com configurações padrão."""
        self.meses_pt: Dict[int, str] = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        logger.info("MMZREmailGenerator inicializado com sucesso")
    
    def load_excel_data(self, filepath: str) -> Optional[pd.ExcelFile]:
        """
        Carrega dados de um arquivo Excel.
        
        Args:
            filepath (str): Caminho para o arquivo Excel
            
        Returns:
            Optional[pd.ExcelFile]: Objeto ExcelFile se bem-sucedido, None caso contrário
            
        Raises:
            FileNotFoundError: Se o arquivo não for encontrado
            ValueError: Se o arquivo não puder ser lido como Excel
        """
        try:
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"Arquivo não encontrado: {filepath}")
                
            excel_file = pd.ExcelFile(filepath)
            logger.info(f"Arquivo carregado: {filepath}")
            logger.info(f"Abas disponíveis: {excel_file.sheet_names}")
            
            return excel_file
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo {filepath}: {e}")
            return None
    
    def extract_performance_data(self, df: pd.DataFrame) -> List[Dict[str, Union[str, float]]]:
        """
        Extrai dados de performance do DataFrame (apenas Mês atual e No ano).
        
        Args:
            df (pd.DataFrame): DataFrame contendo os dados financeiros
            
        Returns:
            List[Dict[str, Union[str, float]]]: Lista com dados de performance
            
        Raises:
            ValueError: Se não encontrar dados de performance na planilha
        """
        performance_data: List[Dict[str, Union[str, float]]] = []
        
        try:
            # Procurar pela palavra "Performance" no DataFrame
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    cell_value = str(df.iloc[i, j])
                    if 'Performance' in cell_value:
                        # Encontrou a seção de performance
                        start_row = i + 2
                        
                        # Extrair dados das próximas linhas
                        for k in range(start_row, min(start_row + 5, len(df))):
                            row = df.iloc[k]
                            if pd.notna(row.iloc[0]):
                                periodo = str(row.iloc[0]).lower()
                                
                                # Filtrar apenas "Mês atual" e "No ano"
                                if "mês" in periodo or "mes" in periodo:
                                    mes_atual = self.meses_pt[datetime.now().month]
                                    periodo = f"{mes_atual}:"
                                elif "ano" in periodo:
                                    periodo = "No ano:"
                                else:
                                    continue
                                    
                                try:
                                    carteira = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                                    benchmark = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0.0
                                    diferenca = float(row.iloc[3]) if pd.notna(row.iloc[3]) and len(row) > 3 else carteira - benchmark
                                    
                                    performance_data.append({
                                        'periodo': periodo,
                                        'carteira': carteira,
                                        'benchmark': benchmark,
                                        'diferenca': diferenca
                                    })
                                except (ValueError, TypeError) as e:
                                    logger.warning(f"Erro ao converter valores numéricos: {e}")
                                    continue
                        
                        if performance_data:
                            logger.info(f"Extraídos {len(performance_data)} registros de performance")
                            return performance_data
            
            # Se não encontrou, lançar erro
            error_msg = "Não foi possível encontrar dados de 'Performance' na planilha"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        except Exception as e:
            logger.error(f"Erro ao extrair dados de performance: {e}")
            raise
    
    def extract_financial_return(self, df: pd.DataFrame) -> float:
        """
        Extrai dados de retorno financeiro do DataFrame.
        
        Args:
            df (pd.DataFrame): DataFrame contendo os dados financeiros
            
        Returns:
            float: Valor do retorno financeiro
            
        Raises:
            ValueError: Se não encontrar dados de retorno financeiro
        """
        try:
            # Procurar pelo termo "Retorno Financeiro"
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                    if ('Retorno Financeiro' in cell_value or 'Retorno' in cell_value) and 'Período' not in cell_value:
                        # Verificar células adjacentes
                        for di, dj in [(1, 0), (0, 1)]:  # Abaixo e à direita
                            ni, nj = i + di, j + dj
                            if ni < len(df) and nj < len(df.columns) and pd.notna(df.iloc[ni, nj]):
                                try:
                                    financial_return = float(df.iloc[ni, nj])
                                    logger.info(f"Retorno financeiro extraído: {financial_return}")
                                    return financial_return
                                except (ValueError, TypeError):
                                    continue
            
            error_msg = "Não foi possível encontrar 'Retorno Financeiro' na planilha"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        except Exception as e:
            logger.error(f"Erro ao extrair retorno financeiro: {e}")
            raise
    
    def extract_highlight_strategies(self, df: pd.DataFrame) -> List[str]:
        """
        Extrai estratégias de destaque (máximo 2).
        
        Args:
            df (pd.DataFrame): DataFrame contendo os dados financeiros
            
        Returns:
            List[str]: Lista com estratégias de destaque (máximo 2)
            
        Raises:
            ValueError: Se não encontrar estratégias de destaque
        """
        strategies: List[str] = []
        
        try:
            # Procurar por "Estratégias de Destaque" ou similar
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    if i < len(df) and j < len(df.columns):
                        cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                        if 'Estratégias de Destaque' in cell_value or 'Destaques' in cell_value:
                            # Extrair estratégias das linhas seguintes
                            start_row = i + 1
                            
                            for k in range(start_row, min(start_row + 5, len(df))):
                                if k < len(df) and len(strategies) < 2:  # Limitar a 2 estratégias
                                    row = df.iloc[k]
                                    for l in range(min(len(row), 3)):  # Limitar a 3 colunas para evitar dados não relacionados
                                        if pd.notna(row.iloc[l]) and str(row.iloc[l]).strip() != '' and len(strategies) < 2:
                                            strategy = str(row.iloc[l])
                                            if not any(s.lower() in strategy.lower() for s in ['estratégia', 'destaque', 'promotor', 'detrator']):
                                                strategies.append(strategy)
                                                if len(strategies) >= 2:  # Parar ao atingir 2 estratégias
                                                    break
                            
                            if strategies:
                                logger.info(f"Extraídas {len(strategies)} estratégias de destaque")
                                return strategies[:2]  # Garantir máximo 2 estratégias
            
            # Se não encontrou, lançar erro
            error_msg = "Não foi possível encontrar 'Estratégias de Destaque' na planilha"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        except Exception as e:
            logger.error(f"Erro ao extrair estratégias de destaque: {e}")
            raise
    
    def extract_promoter_assets(self, df: pd.DataFrame) -> List[str]:
        """
        Extrai ativos promotores (apenas os positivos, máximo 2).
        
        Args:
            df (pd.DataFrame): DataFrame contendo os dados financeiros
            
        Returns:
            List[str]: Lista com ativos promotores (máximo 2)
            
        Raises:
            ValueError: Se não encontrar ativos promotores
        """
        assets: List[str] = []
        
        try:
            # Procurar por "Ativos Promotores" ou similar
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    if i < len(df) and j < len(df.columns):
                        cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                        if 'Ativos Promotores' in cell_value or 'Promotores' in cell_value:
                            # Extrair ativos das linhas seguintes
                            start_row = i + 1
                            
                            for k in range(start_row, min(start_row + 10, len(df))):
                                if k < len(df) and len(assets) < 2:  # Limitar a 2 ativos
                                    row = df.iloc[k]
                                    for l in range(min(len(row), 5)):  # Verificar até 5 colunas
                                        if pd.notna(row.iloc[l]) and str(row.iloc[l]).strip() != '':
                                            asset = str(row.iloc[l])
                                            # Verificar se não contém palavras-chave
                                            if not any(s.lower() in asset.lower() for s in ['ativo', 'promotor', 'detrator', 'estratégia']):
                                                # Verificar se o ativo tem porcentagem positiva
                                                import re
                                                percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', asset)
                                                if percentage_match:
                                                    percentage_str = percentage_match.group(1).replace(',', '.')
                                                    try:
                                                        percentage = float(percentage_str)
                                                        if percentage > 0:  # Somente incluir se for positivo
                                                            assets.append(asset)
                                                            if len(assets) >= 2:  # Limitar a 2 ativos
                                                                break
                                                    except ValueError:
                                                        continue
                            
                            if assets:
                                logger.info(f"Extraídos {len(assets)} ativos promotores")
                                return assets[:2]  # Garantir máximo de 2 ativos
            
            # Se não encontrou, lançar erro
            error_msg = "Não foi possível encontrar 'Ativos Promotores' na planilha ou nenhum ativo com rendimento positivo foi encontrado"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        except Exception as e:
            logger.error(f"Erro ao extrair ativos promotores: {e}")
            raise
    
    def extract_detractor_assets(self, df: pd.DataFrame) -> List[str]:
        """
        Extrai ativos detratores (apenas os negativos, máximo 2).
        
        Args:
            df (pd.DataFrame): DataFrame contendo os dados financeiros
            
        Returns:
            List[str]: Lista com ativos detratores (máximo 2)
            
        Raises:
            ValueError: Se não encontrar ativos detratores
        """
        assets: List[str] = []
        
        try:
            # Procurar por "Ativos Detratores" ou similar
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    if i < len(df) and j < len(df.columns):
                        cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                        if 'Ativos Detratores' in cell_value or 'Detratores' in cell_value:
                            # Extrair ativos das linhas seguintes
                            start_row = i + 1
                            
                            for k in range(start_row, min(start_row + 10, len(df))):
                                if k < len(df) and len(assets) < 2:  # Limitar a 2 ativos
                                    row = df.iloc[k]
                                    for l in range(min(len(row), 5)):  # Verificar até 5 colunas
                                        if pd.notna(row.iloc[l]) and str(row.iloc[l]).strip() != '':
                                            asset = str(row.iloc[l])
                                            # Verificar se não contém palavras-chave
                                            if not any(s.lower() in asset.lower() for s in ['ativo', 'detrator', 'promotor', 'estratégia']):
                                                # Verificar se o ativo tem porcentagem negativa
                                                import re
                                                percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', asset)
                                                if percentage_match:
                                                    percentage_str = percentage_match.group(1).replace(',', '.')
                                                    try:
                                                        percentage = float(percentage_str)
                                                        if percentage < 0:  # Somente incluir se for negativo
                                                            assets.append(asset)
                                                            if len(assets) >= 2:  # Limitar a 2 ativos
                                                                break
                                                    except ValueError:
                                                        continue
                            
                            if assets:
                                logger.info(f"Extraídos {len(assets)} ativos detratores")
                                return assets[:2]  # Garantir máximo de 2 ativos
            
            # Se não encontrou, lançar erro
            error_msg = "Não foi possível encontrar 'Ativos Detratores' na planilha ou nenhum ativo com rendimento negativo foi encontrado"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        except Exception as e:
            logger.error(f"Erro ao extrair ativos detratores: {e}")
            raise
    
    def format_currency(self, value: float) -> str:
        """
        Formata valor como moeda brasileira.
        
        Args:
            value (float): Valor a ser formatado
            
        Returns:
            str: Valor formatado como moeda brasileira
        """
        if value >= 0:
            return f"R$ {value:,.2f}".replace(",", ".")
        else:
            return f"-R$ {abs(value):,.2f}".replace(",", ".")
    
    def format_percentage(self, value: float) -> str:
        """
        Formata valor como percentual.
        
        Args:
            value (float): Valor a ser formatado
            
        Returns:
            str: Valor formatado como percentual
        """
        if value > 0:
            return f"+{value:.2f}%"
        else:
            return f"{value:.2f}%"
    
    def generate_html_email(self, client_name: str, data_ref: datetime, portfolios_data: List[Dict[str, Any]]) -> str:
        """
        Gera o HTML completo do email.
        
        Args:
            client_name (str): Nome do cliente
            data_ref (datetime): Data de referência do relatório
            portfolios_data (List[Dict[str, Any]]): Dados das carteiras do cliente
            
        Returns:
            str: HTML completo do email
        """
        try:
            # Configurar mês/ano
            mes = self.meses_pt[data_ref.month]
            ano = data_ref.year
            
            logger.info(f"Gerando HTML para {client_name} - {mes}/{ano}")
            
            # HTML Header
            html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="light">
    <meta name="supported-color-schemes" content="light">
    <!--[if mso]>
    <style type="text/css">
    body, table, td {{font-family: Arial, Helvetica, sans-serif !important;}}
    </style>
    <![endif]-->
    <style>
    /* Estilos para forçar modo claro em dispositivos com tema escuro */
    :root {{
        color-scheme: light;
        supported-color-schemes: light;
    }}
    @media (prefers-color-scheme: dark) {{
        body,
        .body-wrapper {{
            background-color: #f4f4f4 !important;
        }}
        .content-wrapper {{
            background-color: #ffffff !important;
            color: #333333 !important;
        }}
        .header-bg {{
            background-color: #0D2035 !important;
        }}
        .header-text {{
            color: #ffffff !important;
        }}
        .section-bg {{
            background-color: #ffffff !important;
        }}
        .performance-header {{
            color: #0D2035 !important;
            border-bottom-color: #e0e0e0 !important;
        }}
        .data-table {{
            background-color: #ffffff !important;
        }}
        .table-header {{
            background-color: #f8f9fa !important;
            color: #0D2035 !important;
        }}
        .highlight-section {{
            background-color: #f8f9fa !important;
        }}
        .promoters-section {{
            background-color: #e8f5e9 !important;
        }}
        .detractors-section {{
            background-color: #ffebee !important;
        }}
        td, th, p, h1, h2, h3, h4, h5, h6, li {{
            color: inherit !important;
        }}
        .portfolio-header {{
            background-color: #0D2035 !important;
            color: #ffffff !important;
        }}
    }}
    </style>
</head>
<body class="body-wrapper" style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', sans-serif; line-height: 1.4; color: #333333; background-color: #f4f4f4;">
    <table role="presentation" style="width: 100%; border-collapse: collapse; border: 0; border-spacing: 0; background: #f4f4f4;">
        <tr>
            <td align="center" style="padding: 0;">
                <table role="presentation" class="content-wrapper" style="width: 100%; max-width: 800px; border-collapse: collapse; border: 0; border-spacing: 0; text-align: left; background: #ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                    <!-- Header -->
                    <tr>
                        <td style="padding: 0;">
                            <table role="presentation" class="header-bg" style="width: 100%; border-collapse: collapse; background: #0D2035;">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table role="presentation" style="width: 100%; border-collapse: collapse;">
                                            <tr>
                                                <td style="text-align: center; vertical-align: middle; width: 120px;">
                                                    <img src="documentos/img/logo-MMZR-azul.png" alt="MMZR Family Office" style="width: 120px; height: 100px; display: inline-block;">
                                                </td>
                                                <td style="text-align: left; vertical-align: middle; padding-left: 10px;">
                                                    <p class="header-text" style="margin: 0; font-size: 21px; color: #ffffff; opacity: 0.9; line-height: 1.2;">MMZR Family Office</p>
                                                    <p class="header-text" style="margin: 0; font-size: 14px; color: #ffffff; opacity: 0.9; line-height: 1.2;">Relatório Mensal de Performance - {mes} de {ano}</p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td class="section-bg" style="padding: 20px 20px; background-color: #ffffff;">
                            <h2 style="font-size: 15px; color: #0D2035; margin-bottom: 12px; margin-top: 0;">Olá {client_name},</h2>
                            
                            <p style="margin-top: 0; margin-bottom: 9px; ">Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.</p>"""
        
            # Adicionar cada carteira
            for portfolio in portfolios_data:
                html += self.generate_portfolio_section(portfolio)
        
            # Coletar todos os comentários das carteiras
            comentarios_todos = []
            for portfolio in portfolios_data:
                comentario = portfolio.get('comentarios', '')
                if comentario and comentario.strip():
                    comentarios_todos.append(comentario.strip())
            
            # Juntar comentários se houver múltiplos
            comentario_final = ' | '.join(comentarios_todos) if comentarios_todos else ""
        
            # Adicionar seções que faltaram
            html += self.generate_observacoes_section(comentario_final)
            html += self.generate_principais_indicadores_section()
            html += self.generate_botao_carta_section(mes, ano)
        
            # Footer
            html += f"""
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 12px 20px; text-align: center;">
                            <p style="margin: 0 0 3px 0; color: #666666; font-size: 11px;">MMZR Family Office | Gestão de Patrimônio</p>
                            <p style="margin: 0 0 3px 0; color: #666666; font-size: 11px;">Este é um email automático. Por favor, não responda.</p>
                            <p style="margin: 0; color: #666666; font-size: 11px;">© {ano} MMZR Family Office. Todos os direitos reservados.</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""
            
            return html
        
        except Exception as e:
            logger.error(f"Erro ao gerar HTML do email: {e}")
            return ""
    
    def generate_portfolio_section(self, portfolio: Dict[str, Any]) -> str:
        """
        Gera a seção HTML de uma carteira específica.
        
        Args:
            portfolio (Dict[str, Any]): Dados da carteira
            
        Returns:
            str: HTML da seção da carteira
        """
        name = portfolio.get('name', 'Carteira')
        portfolio_type = portfolio.get('type', 'Diversificada')
        comentarios = portfolio.get('comentarios', '')  # Comentários específicos da carteira
        data = portfolio.get('data', {})
        
        performance_data = data.get('performance', [])
        retorno_financeiro = data.get('retorno_financeiro', 0)
        estrategias_destaque = data.get('estrategias_destaque', [])
        ativos_promotores = data.get('ativos_promotores', [])
        ativos_detratores = data.get('ativos_detratores', [])
        
        html = f"""
                            <!-- Carteira: {name} -->
                            <table role="presentation" style="width: 100%; margin: 20px 0 0 0; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); background-color: #ffffff;">
                                <tr>
                                    <td class="header-bg portfolio-header" style="background-color: #0D2035; color: #ffffff; padding: 10px 15px;">
                                        <h3 style="margin: 0; font-size: 16px; font-weight: 500;">{name} <span style="font-weight: 300; font-size: 13px; margin-left: 8px; opacity: 0.8;">| {portfolio_type}</span></h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="section-bg" style="padding: 15px; background-color: #ffffff;">
                                        {self.generate_performance_table(performance_data, retorno_financeiro)}
                                        
                                        {self.generate_highlight_strategies_section(estrategias_destaque)}
                                        
                                        {self.generate_promoter_assets_section(ativos_promotores)}
                                        
                                        {self.generate_detractor_assets_section(ativos_detratores)}
                                    </td>
                                </tr>
                            </table>
"""
        return html
    
    def generate_performance_table(self, performance_data, retorno_financeiro=None):
        """Gera a tabela HTML de performance, incluindo retorno financeiro"""
        
        # Filtrar apenas os períodos necessários (Mês atual e No ano) sem duplicações
        filtered_data = []
        mes_adicionado = False
        ano_adicionado = False
        
        for item in performance_data:
            periodo = item['periodo'].lower() if isinstance(item['periodo'], str) else ""
            
            # Verificar se é mês atual
            if ":" in periodo and any(m.lower() in periodo for m in self.meses_pt.values()) and not mes_adicionado:
                filtered_data.append(item)
                mes_adicionado = True
            # Verificar se é ano atual
            elif "no ano" in periodo and not ano_adicionado:
                filtered_data.append(item)
                ano_adicionado = True
                
            # Se já temos os dois períodos, parar
            if mes_adicionado and ano_adicionado:
                break
        
        html = """
                                        <h4 class="performance-header" style="font-size: 18px; color: #0D2035; margin: 0 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">Performance</h4>
                                        <table role="presentation" class="data-table" style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 15px; background-color: #ffffff;">
                                            <thead>
                                                <tr>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6;">Período</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Carteira</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Benchmark</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Carteira vs. Benchmark</th>
                                                </tr>
                                            </thead>
                                            <tbody>
"""
        
        # Adicionar cada linha de performance
        for item in filtered_data:
            periodo = item['periodo']
            carteira = item['carteira']
            benchmark = item['benchmark']
            diferenca = item['diferenca']
            
            # Determinar cores com base nos valores
            carteira_color = "#28a745" if carteira > 0 else "#dc3545" if carteira < 0 else "#333333"
            diferenca_color = "#28a745" if diferenca > 0 else "#dc3545" if diferenca < 0 else "#333333"
            
            html += f"""
                                                <tr>
                                                    <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6; background-color: #ffffff;">{periodo}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {carteira_color}; font-weight: 500; background-color: #ffffff;">{self.format_percentage(carteira)}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; background-color: #ffffff;">{self.format_percentage(benchmark)}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {diferenca_color}; font-weight: 500; background-color: #ffffff;">{self.format_percentage(diferenca).replace('%', ' p.p.')}</td>
                                                </tr>
"""
        
        # Adicionar linha de retorno financeiro se disponível
        if retorno_financeiro is not None:
            color = "#28a745" if retorno_financeiro > 0 else "#dc3545" if retorno_financeiro < 0 else "#333333"
            html += f"""
                                                <tr>
                                                    <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6; font-weight: 500; background-color: #ffffff;">Retorno Financeiro:</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {color}; font-weight: 500; background-color: #ffffff;" colspan="3">{self.format_currency(retorno_financeiro)}</td>
                                                </tr>
"""
        
        html += """
                                            </tbody>
                                        </table>
"""
        return html
    
    def generate_financial_return_section(self, retorno_financeiro):
        """Gera a seção de retorno financeiro"""
        
        html = f"""
                                        <h4 class="performance-header" style="font-size: 18px; color: #0D2035; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">Retorno Financeiro</h4>
                                        <p style="font-size: 15px; margin: 8px 0 15px 0; padding: 10px; background-color: #f8f9fa; border-radius: 5px; text-align: center; font-weight: 500; color: #0D2035;">
                                            {self.format_currency(retorno_financeiro)}
                                        </p>
"""
        return html
    
    def generate_highlight_strategies_section(self, estrategias):
        """Gera a seção de estratégias de destaque"""
        
        html = """
                                        <h4 class="performance-header" style="font-size: 18px; color: #0D2035; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">Estratégias de Destaque</h4>
                                        <ul class="highlight-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #f8f9fa; border-radius: 5px; color: #333333;">
"""
        
        for estrategia in estrategias:
            html += f"""
                                            <li style="margin-bottom: 6px; font-size: 13px;">{estrategia}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def generate_promoter_assets_section(self, ativos):
        """Gera a seção de ativos promotores"""
        
        html = """
                                        <h4 class="performance-header" style="font-size: 18px; color: #0D2035; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">Ativos Promotores</h4>
                                        <ul class="promoters-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #e8f5e9; border-radius: 5px; color: #2e7d32;">
"""
        
        for ativo in ativos:
            # Adicionar o símbolo "+" antes da porcentagem se for um valor positivo
            import re
            ativo_formatado = ativo
            percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', ativo)
            if percentage_match:
                percentage_str = percentage_match.group(1).replace(',', '.')
                try:
                    percentage = float(percentage_str)
                    if percentage > 0 and not percentage_str.startswith('+'):
                        # Substituir a porcentagem sem o "+" por uma com o "+"
                        ativo_formatado = ativo.replace(f"({percentage_str}%)", f"(+{percentage_str}%)")
                except ValueError:
                    pass
                    
            html += f"""
                                            <li style="margin-bottom: 6px; font-size: 13px;">{ativo_formatado}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def generate_detractor_assets_section(self, ativos):
        """Gera a seção de ativos detratores"""
        
        html = """
                                        <h4 class="performance-header" style="font-size: 18px; color: #0D2035; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px;">Ativos Detratores</h4>
                                        <ul class="detractors-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #ffebee; border-radius: 5px; color: #c62828;">
"""
        
        for ativo in ativos:
            html += f"""
                                            <li style="margin-bottom: 6px; font-size: 13px;">{ativo}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def generate_observacoes_section(self, comentario_adicional: str = "") -> str:
        """Gera a seção de observações incluindo comentário adicional da planilha."""
        
        html = """
                            <!-- Observações finais -->
                            <table role="presentation" style="width: 100%; margin-top: 20px; border-collapse: collapse; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                                <tr>
                                    <td style="padding: 15px;">
                                        <p style="margin: 0 0 12px 0; color: #555555; font-size: 13px; line-height: 18px;">
                                            <strong style="font-weight: bold;">Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                                        </p>
                                        <p style="margin: 0; color: #555555; font-size: 12px; font-style: italic; line-height: 16px;">
                                            <strong style="font-weight: bold;">Obs.:</strong> Conforme solicitado, deixo o Felipe e Fernandito em cópia para também receberem as informações.
                                        </p>"""
        
        if comentario_adicional:
            html += f"""
                                        <p style="margin: 12px 0 0 0; color: #555555; font-size: 13px; line-height: 18px;">
                                            <strong style="font-weight: bold;">Comentário:</strong> {comentario_adicional}
                                        </p>"""
        
        html += """
                                    </td>
                                </tr>
                            </table>
"""
        return html
    
    def generate_principais_indicadores_section(self) -> str:
        """Gera a seção de principais indicadores."""
        
        html = """
                            <!-- Principais indicadores -->
                            <table role="presentation" style="width: 100%; margin-top: 15px; border-collapse: collapse; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                                <tr>
                                    <td style="padding: 12px;">
                                        <p style="margin: 0 0 8px 0; font-weight: bold; color: #333333; font-size: 13px; line-height: 16px;">Principais indicadores:</p>
                                        <p style="margin: 0; color: #555555; font-size: 11px; line-height: 15px;">
                                            Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                                            Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                                        </p>
                                    </td>
                                </tr>
                            </table>
"""
        return html
    
    def generate_botao_carta_section(self, mes: str, ano: int) -> str:
        """Gera a seção do botão da carta mensal."""
        
        mes_lowercase = mes.lower()
        carta_link = f"https://www.mmzrfo.com.br/post/carta-mensal-{mes_lowercase}-{ano}"
        
        html = f"""
                            <!-- Link para carta mensal como botão azul -->
                            <table role="presentation" style="width: 100%; margin-top: 25px; border-collapse: collapse;">
                                <tr>
                                    <td align="center" style="padding: 0;">
                                        <table role="presentation" style="border-collapse: collapse; background-color: #0D2035; border-radius: 4px;">
                                            <tr>
                                                <td style="padding: 12px 24px; text-align: center;">
                                                    <a href="{carta_link}" target="_blank" style="color: #ffffff; text-decoration: none; font-weight: bold; font-size: 14px; line-height: 18px;">Confira nossa carta completa: Carta {mes} {ano}</a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
"""
        return html
    
    def save_email_to_file(self, html_content: str, client_name: str, output_path: Optional[str] = None) -> str:
        """
        Salva o conteúdo HTML do e-mail em um arquivo.
        
        Args:
            html_content (str): Conteúdo HTML do email
            client_name (str): Nome do cliente
            output_path (Optional[str]): Caminho de saída personalizado
            
        Returns:
            str: Caminho do arquivo salvo
            
        Raises:
            IOError: Se não conseguir salvar o arquivo
        """
        try:
            # Remover caracteres inválidos para nome de arquivo
            safe_client_name = "".join([c if c.isalnum() or c in [' ', '_'] else '_' for c in client_name])
            safe_client_name = safe_client_name.replace(' ', '_')
            
            # Data atual para nome do arquivo
            date_str = datetime.now().strftime("%Y%m%d")
            
            # Caminho de saída
            if not output_path:
                filename = f"relatorio_mensal_{safe_client_name}_{date_str}.html"
                output_path = filename
            
            # Salvar o arquivo
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logger.info(f"Relatório salvo em: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Erro ao salvar arquivo: {e}")
            raise IOError(f"Não foi possível salvar o arquivo: {e}")


def process_and_generate_report(excel_path: str, client_config: Dict[str, Any]) -> Union[str, bool]:
    """
    Processa os dados e gera o relatório de e-mail.
    
    Args:
        excel_path (str): Caminho para o arquivo Excel
        client_config (Dict[str, Any]): Configuração do cliente
        
    Returns:
        Union[str, bool]: Caminho do arquivo gerado ou False se houver erro
    """
    try:
        # Criar o gerador
        generator = MMZREmailGenerator()
        
        # Carregar o Excel
        excel_file = generator.load_excel_data(excel_path)
        if not excel_file:
            logger.error("Erro ao carregar arquivo Excel.")
            return False
        
        # Inicializar dados do cliente
        client_name = client_config.get('name', 'Cliente')
        client_email = client_config.get('email', '')
        
        # Data de referência (hoje como padrão)
        data_ref = datetime.now()
        
        # Processar cada carteira
        portfolios_data = []
        
        for portfolio_config in client_config.get('portfolios', []):
            # Buscar a aba correspondente no Excel
            sheet_name = portfolio_config.get('sheet_name', '')
            
            if sheet_name and sheet_name in excel_file.sheet_names:
                # Ler os dados da aba
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Extrair todos os dados necessários
                portfolio_data = {
                    'name': portfolio_config.get('name', 'Carteira'),
                    'type': portfolio_config.get('type', 'Diversificada'),
                    'comentarios': portfolio_config.get('comentarios', ''),  # Adicionar suporte a comentários
                    'data': {
                        'performance': generator.extract_performance_data(df),
                        'retorno_financeiro': generator.extract_financial_return(df),
                        'estrategias_destaque': generator.extract_highlight_strategies(df),
                        'ativos_promotores': generator.extract_promoter_assets(df),
                        'ativos_detratores': generator.extract_detractor_assets(df)
                    }
                }
                
                portfolios_data.append(portfolio_data)
            else:
                error_msg = f"Aba '{sheet_name}' não encontrada no Excel. O relatório não pode ser gerado."
                logger.error(error_msg)
                raise ValueError(error_msg)
        
        # Gerar o HTML do e-mail
        html_content = generator.generate_html_email(client_name, data_ref, portfolios_data)
        
        # Salvar o e-mail em um arquivo
        output_file = generator.save_email_to_file(html_content, client_name)
        
        logger.info(f"Relatório gerado com sucesso para {client_name}!")
        return output_file
    
    except Exception as e:
        logger.error(f"Erro ao gerar relatório: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return False


# Exemplo de uso
if __name__ == "__main__":
    # Configuração do cliente
    client = {
        'name': 'João Silva',
        'email': 'joao.silva@example.com',
        'portfolios': [
            {
                'name': 'Carteira Moderada',
                'type': 'Renda Variável + Renda Fixa',
                'sheet_name': 'Base Consolidada',
                'benchmark_name': 'IPCA+5%'
            },
            {
                'name': 'Carteira Conservadora',
                'type': 'Renda Fixa',
                'sheet_name': 'Base Clientes',
                'benchmark_name': 'CDI'
            }
        ]
    }
    
    # Processar e gerar relatório
    process_and_generate_report('documentos/dados/Planilha Inteli.xlsm', client) 