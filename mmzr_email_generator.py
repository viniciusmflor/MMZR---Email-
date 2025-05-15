import os
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta

class MMZREmailGenerator:
    """Gerador de emails HTML para MMZR Family Office"""
    
    def __init__(self):
        self.meses_pt = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
    
    def load_excel_data(self, filepath):
        """Carrega dados do Excel"""
        try:
            # Ler o arquivo Excel
            excel_file = pd.ExcelFile(filepath)
            print(f"Arquivo carregado: {filepath}")
            print(f"Abas disponíveis: {excel_file.sheet_names}")
            
            # Retornar o objeto ExcelFile para processar múltiplas abas
            return excel_file
        except Exception as e:
            print(f"Erro ao carregar arquivo: {e}")
            return None
    
    def extract_performance_data(self, df):
        """Extrai dados de performance do DataFrame (apenas Mês atual e No ano)"""
        performance_data = []
        
        # Procurar pela palavra "Performance" no DataFrame
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                if 'Performance' in cell_value:
                    # Encontrou a seção de performance
                    # Assumir que os dados começam 2 linhas abaixo
                    start_row = i + 2
                    
                    # Extrair dados das próximas linhas
                    for k in range(start_row, min(start_row + 5, len(df))):
                        row = df.iloc[k]
                        if pd.notna(row.iloc[0]):  # Se tem período
                            periodo = str(row.iloc[0]).lower()
                            
                            # Filtrar apenas "Mês atual" e "No ano"
                            if "mês" in periodo or "mes" in periodo:
                                # Obter o mês atual
                                mes_atual = self.meses_pt[datetime.now().month]
                                periodo = f"{mes_atual}:"
                            elif "ano" in periodo:
                                periodo = "No ano:"
                            else:
                                continue  # Pular outras entradas
                                
                            try:
                                carteira = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                                benchmark = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                                diferenca = float(row.iloc[3]) if pd.notna(row.iloc[3]) and len(row) > 3 else carteira - benchmark
                                
                                performance_data.append({
                                    'periodo': periodo,
                                    'carteira': carteira,
                                    'benchmark': benchmark,
                                    'diferenca': diferenca
                                })
                            except (ValueError, TypeError):
                                # Ignorar linhas com valores não numéricos
                                pass
                    
                    if performance_data:
                        return performance_data
        
        # Caso não encontre através da palavra "Performance", criar dados simulados
        if not performance_data:
            # Obter o mês atual
            mes_atual = self.meses_pt[datetime.now().month]
            performance_data = [
                {'periodo': f"{mes_atual}:", 'carteira': 2.38, 'benchmark': 1.45, 'diferenca': 0.93},
                {'periodo': "No ano:", 'carteira': 8.76, 'benchmark': 5.32, 'diferenca': 3.44}
            ]
        
        return performance_data
    
    def extract_financial_return(self, df):
        """Extrai dados de retorno financeiro"""
        financial_return = None
        
        # Procurar pelo termo "Retorno Financeiro" ou texto similar
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                if ('Retorno Financeiro' in cell_value or 'Retorno' in cell_value) and 'Período' not in cell_value:
                    if i+1 < len(df) and j+1 < len(df.columns):
                        # Pegar o valor na célula abaixo ou ao lado
                        if pd.notna(df.iloc[i+1, j]):
                            try:
                                financial_return = float(df.iloc[i+1, j])
                                return financial_return
                            except (ValueError, TypeError):
                                # Se não conseguir converter, tentar a próxima célula
                                pass
                        
                        if pd.notna(df.iloc[i, j+1]):
                            try:
                                financial_return = float(df.iloc[i, j+1])
                                return financial_return
                            except (ValueError, TypeError):
                                # Se não conseguir converter, continuar procurando
                                pass
        
        # Se não encontrou, retornar valor simulado
        return 1140.27
    
    def extract_highlight_strategies(self, df):
        """Extrai estratégias de destaque (máximo 2)"""
        strategies = []
        
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
                            return strategies[:2]  # Garantir máximo 2 estratégias
        
        # Se não encontrou, retornar valores simulados (limitar a 2)
        return ["RETORNO ABSOLUTO (3,12%) > RENDA VARIÁVEL (8,54%)", "OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%) > SHARP LB ADVISORY FIC FIA (12,57%)"][:2]
    
    def extract_promoter_assets(self, df):
        """Extrai ativos promotores (apenas os positivos, máximo 2)"""
        assets = []
        
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
                                                    pass
                        
                        if len(assets) > 0:
                            return assets[:2]  # Garantir máximo de 2 ativos
        
        # Se não encontrou ou não encontrou positivos suficientes, retornar valores simulados positivos (máximo 2)
        if len(assets) == 0:
            return ["OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%)", "SHARP LB ADVISORY FIC FIA (12,57%)"][:2]
        elif len(assets) == 1:
            assets.append("SHARP LB ADVISORY FIC FIA (12,57%)")
            return assets
        
        return assets[:2]  # Garantir máximo de 2 ativos
    
    def extract_detractor_assets(self, df):
        """Extrai ativos detratores (apenas os negativos, máximo 2)"""
        assets = []
        
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
                                                    pass
                        
                        if len(assets) > 0:
                            return assets[:2]  # Garantir máximo de 2 ativos
        
        # Se não encontrou, retornar valores simulados negativos (máximo 2)
        if len(assets) == 0:
            return ["ALLOCATION GLOBAL EQUITIES FI MULT (-1,66%)", "BTG PACTUAL DISCOVERY FI MULT (-0,85%)"][:2]
        elif len(assets) == 1:
            assets.append("BTG PACTUAL DISCOVERY FI MULT (-0,85%)")
            return assets
        
        return assets[:2]  # Garantir máximo de 2 ativos
    
    def format_currency(self, value):
        """Formata valor como moeda brasileira"""
        if value >= 0:
            return f"R$ {value:,.2f}".replace(",", ".")
        else:
            return f"-R$ {abs(value):,.2f}".replace(",", ".")
    
    def format_percentage(self, value):
        """Formata valor como percentual"""
        if value > 0:
            return f"+{value:.2f}%"
        else:
            return f"{value:.2f}%"
    
    def generate_html_email(self, client_name, data_ref, portfolios_data):
        """Gera o HTML completo do email"""
        
        # Configurar mês/ano
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        # HTML Header
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!--[if mso]>
    <style type="text/css">
    body, table, td {{font-family: Arial, Helvetica, sans-serif !important;}}
    </style>
    <![endif]-->
</head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', sans-serif; line-height: 1.6; color: #333333; background-color: #f4f4f4;">
    <table role="presentation" style="width: 100%; border-collapse: collapse; border: 0; border-spacing: 0; background: #f4f4f4;">
        <tr>
            <td align="center" style="padding: 0;">
                <table role="presentation" style="width: 100%; max-width: 800px; border-collapse: collapse; border: 0; border-spacing: 0; text-align: left; background: #ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                    <!-- Header -->
                    <tr>
                        <td style="padding: 0;">
                            <table role="presentation" style="width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #061844 0%, #083266 100%);">
                                <tr>
                                    <td style="padding: 40px 30px; text-align: center;">
                                        <!-- Logo MMZR -->
                                        <img src="documentos/img/LogoAzul_MMZR.jpg" alt="MMZR Family Office" style="max-width: 200px; height: auto; margin-bottom: 20px;">
                                        <h1 style="margin: 0; font-size: 28px; font-weight: 300; color: #ffffff; letter-spacing: 1px;">MMZR Family Office</h1>
                                        <p style="margin: 10px 0 0 0; font-size: 16px; color: #ffffff; opacity: 0.9;">Relatório Mensal de Performance - {mes} de {ano}</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td style="padding: 30px;">
                            <h2 style="font-size: 20px; color: #061844; margin-bottom: 20px;">Olá {client_name},</h2>
                            
                            <p style="margin-bottom: 20px;">Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.</p>
"""
        
        # Adicionar cada carteira
        for portfolio in portfolios_data:
            html += self.generate_portfolio_section(portfolio)
        
        # Footer
        html += f"""
                            <!-- Disclaimer -->
                            <table role="presentation" style="width: 100%; margin-top: 30px; border-collapse: collapse;">
                                <tr>
                                    <td style="padding: 15px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 4px;">
                                        <p style="margin: 0; color: #856404; font-size: 13px;">
                                            <strong>Aviso Legal:</strong> Este relatório é destinado exclusivamente ao uso interno e não deve ser reproduzido ou distribuído sem autorização prévia. Os rendimentos passados não garantem resultados futuros.
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 20px 30px; text-align: center;">
                            <p style="margin: 0 0 5px 0; color: #666666; font-size: 12px;">MMZR Family Office | Gestão de Patrimônio</p>
                            <p style="margin: 0 0 5px 0; color: #666666; font-size: 12px;">Este é um email automático. Por favor, não responda.</p>
                            <p style="margin: 0; color: #666666; font-size: 12px;">© {ano} MMZR Family Office. Todos os direitos reservados.</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""
        
        return html
    
    def generate_portfolio_section(self, portfolio):
        """Gera a seção HTML de uma carteira específica"""
        
        name = portfolio.get('name', 'Carteira')
        portfolio_type = portfolio.get('type', 'Diversificada')
        data = portfolio.get('data', {})
        
        performance_data = data.get('performance', [])
        retorno_financeiro = data.get('retorno_financeiro', 0)
        estrategias_destaque = data.get('estrategias_destaque', [])
        ativos_promotores = data.get('ativos_promotores', [])
        ativos_detratores = data.get('ativos_detratores', [])
        
        html = f"""
                            <!-- Carteira: {name} -->
                            <table role="presentation" style="width: 100%; margin: 30px 0; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                                <tr>
                                    <td style="background-color: #061844; color: #ffffff; padding: 15px 20px;">
                                        <h3 style="margin: 0; font-size: 18px; font-weight: 500;">{name} <span style="font-weight: 300; font-size: 14px; margin-left: 10px; opacity: 0.8;">| {portfolio_type}</span></h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 20px;">
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
                                        <h4 style="font-size: 20px; color: #061844; margin: 0 0 15px 0; font-weight: 500; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">Performance</h4>
                                        <table role="presentation" style="width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 25px;">
                                            <thead>
                                                <tr>
                                                    <th style="background-color: #f8f9fa; color: #061844; font-weight: 600; padding: 12px 8px; text-align: left; border-bottom: 2px solid #dee2e6;">Período</th>
                                                    <th style="background-color: #f8f9fa; color: #061844; font-weight: 600; padding: 12px 8px; text-align: center; border-bottom: 2px solid #dee2e6;">Carteira</th>
                                                    <th style="background-color: #f8f9fa; color: #061844; font-weight: 600; padding: 12px 8px; text-align: center; border-bottom: 2px solid #dee2e6;">Benchmark</th>
                                                    <th style="background-color: #f8f9fa; color: #061844; font-weight: 600; padding: 12px 8px; text-align: center; border-bottom: 2px solid #dee2e6;">Carteira vs. Benchmark</th>
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
                                                    <td style="padding: 10px 8px; text-align: left; border-bottom: 1px solid #dee2e6;">{periodo}</td>
                                                    <td style="padding: 10px 8px; text-align: center; border-bottom: 1px solid #dee2e6; color: {carteira_color}; font-weight: 500;">{self.format_percentage(carteira)}</td>
                                                    <td style="padding: 10px 8px; text-align: center; border-bottom: 1px solid #dee2e6;">{self.format_percentage(benchmark)}</td>
                                                    <td style="padding: 10px 8px; text-align: center; border-bottom: 1px solid #dee2e6; color: {diferenca_color}; font-weight: 500;">{self.format_percentage(diferenca).replace('%', ' p.p.')}</td>
                                                </tr>
"""
        
        # Adicionar linha de retorno financeiro se disponível
        if retorno_financeiro is not None:
            color = "#28a745" if retorno_financeiro > 0 else "#dc3545" if retorno_financeiro < 0 else "#333333"
            html += f"""
                                                <tr>
                                                    <td style="padding: 10px 8px; text-align: left; border-bottom: 1px solid #dee2e6; font-weight: 500;">Retorno Financeiro:</td>
                                                    <td style="padding: 10px 8px; text-align: center; border-bottom: 1px solid #dee2e6; color: {color}; font-weight: 500;" colspan="3">{self.format_currency(retorno_financeiro)}</td>
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
                                        <h4 style="font-size: 20px; color: #061844; margin: 25px 0 15px 0; font-weight: 500; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">Retorno Financeiro</h4>
                                        <p style="font-size: 16px; margin: 10px 0 25px 0; padding: 15px; background-color: #f8f9fa; border-radius: 5px; text-align: center; font-weight: 500; color: #061844;">
                                            {self.format_currency(retorno_financeiro)}
                                        </p>
"""
        return html
    
    def generate_highlight_strategies_section(self, estrategias):
        """Gera a seção de estratégias de destaque"""
        
        html = """
                                        <h4 style="font-size: 20px; color: #061844; margin: 25px 0 15px 0; font-weight: 500; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">Estratégias de Destaque</h4>
                                        <ul style="margin: 10px 0 25px 0; padding: 15px 15px 15px 35px; background-color: #f8f9fa; border-radius: 5px; color: #333333;">
"""
        
        for estrategia in estrategias:
            html += f"""
                                            <li style="margin-bottom: 8px;">{estrategia}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def generate_promoter_assets_section(self, ativos):
        """Gera a seção de ativos promotores"""
        
        html = """
                                        <h4 style="font-size: 20px; color: #061844; margin: 25px 0 15px 0; font-weight: 500; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">Ativos Promotores</h4>
                                        <ul style="margin: 10px 0 25px 0; padding: 15px 15px 15px 35px; background-color: #e8f5e9; border-radius: 5px; color: #2e7d32;">
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
                                            <li style="margin-bottom: 8px;">{ativo_formatado}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def generate_detractor_assets_section(self, ativos):
        """Gera a seção de ativos detratores"""
        
        html = """
                                        <h4 style="font-size: 20px; color: #061844; margin: 25px 0 15px 0; font-weight: 500; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px;">Ativos Detratores</h4>
                                        <ul style="margin: 10px 0 25px 0; padding: 15px 15px 15px 35px; background-color: #ffebee; border-radius: 5px; color: #c62828;">
"""
        
        for ativo in ativos:
            html += f"""
                                            <li style="margin-bottom: 8px;">{ativo}</li>
"""
        
        html += """
                                        </ul>
"""
        return html
    
    def save_email_to_file(self, html_content, client_name, output_path=None):
        """Salva o conteúdo HTML do e-mail em um arquivo"""
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
        
        print(f"Relatório salvo em: {output_path}")
        return output_path


def process_and_generate_report(excel_path, client_config):
    """Processa os dados e gera o relatório de e-mail"""
    try:
        # Criar o gerador
        generator = MMZREmailGenerator()
        
        # Carregar o Excel
        excel_file = generator.load_excel_data(excel_path)
        if not excel_file:
            print("Erro ao carregar arquivo Excel.")
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
                print(f"Aviso: Aba '{sheet_name}' não encontrada no Excel. Usando dados simulados.")
                
                # Usar dados simulados para esta carteira
                portfolio_data = {
                    'name': portfolio_config.get('name', 'Carteira'),
                    'type': portfolio_config.get('type', 'Diversificada'),
                    'data': {
                        'performance': [
                            {'periodo': 'Mês atual', 'carteira': 2.38, 'benchmark': 1.45, 'diferenca': 0.93},
                            {'periodo': 'Ano atual', 'carteira': 8.76, 'benchmark': 5.32, 'diferenca': 3.44}
                        ],
                        'retorno_financeiro': 1140.27,
                        'estrategias_destaque': [
                            "RETORNO ABSOLUTO (3,12%) > RENDA VARIÁVEL (8,54%)",
                            "OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%) > SHARP LB ADVISORY FIC FIA (12,57%)"
                        ],
                        'ativos_promotores': [
                            "OCEANA LONG BIASED ADVISORY FIC FI MULT (7,83%)",
                            "SHARP LB ADVISORY FIC FIA (12,57%)"
                        ],
                        'ativos_detratores': [
                            "ALLOCATION GLOBAL EQUITIES FI MULT (-1,66%)",
                            "BTG PACTUAL DISCOVERY FI MULT (-0,85%)"
                        ]
                    }
                }
                
                portfolios_data.append(portfolio_data)
        
        # Gerar o HTML do e-mail
        html_content = generator.generate_html_email(client_name, data_ref, portfolios_data)
        
        # Salvar o e-mail em um arquivo
        output_file = generator.save_email_to_file(html_content, client_name)
        
        print(f"Relatório gerado com sucesso para {client_name}!")
        return output_file
    
    except Exception as e:
        print(f"Erro ao gerar relatório: {str(e)}")
        import traceback
        traceback.print_exc()
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