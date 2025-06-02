import os
import pandas as pd
from datetime import date, datetime, timedelta
import re

class MMZREmailGenerator:
    """Gerador de emails HTML para MMZR Family Office"""
    
    def __init__(self, compact_mode=True):
        """
        Inicializa o gerador
        
        Args:
            compact_mode (bool): Ativa modo compacto (padrão: True)
        """
        self.meses_pt = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        self.compact_mode = compact_mode
    
    def generate_email_subject(self, data_ref=None):
        """
        Gera o assunto automático do email no padrão:
        'MMZR Family Office | Desempenho [Mês] de [Ano]'
        
        Args:
            data_ref (datetime): Data de referência (padrão: hoje)
        
        Returns:
            str: Assunto formatado
        """
        if data_ref is None:
            data_ref = datetime.now()
        
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        return f"MMZR Family Office | Desempenho {mes} de {ano}"
    
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
        
        # Se não encontrou, lançar erro
        error_msg = "Erro: Não foi possível encontrar dados de 'Performance' na planilha."
        print(error_msg)
        raise ValueError(error_msg)
    
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
        
        # Se não encontrou, lançar erro
        error_msg = "Erro: Não foi possível encontrar 'Retorno Financeiro' na planilha."
        print(error_msg)
        raise ValueError(error_msg)
    
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
        
        # Se não encontrou, lançar erro
        error_msg = "Erro: Não foi possível encontrar 'Estratégias de Destaque' na planilha."
        print(error_msg)
        raise ValueError(error_msg)
    
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
        
        # Se não encontrou, lançar erro
        error_msg = "Erro: Não foi possível encontrar 'Ativos Promotores' na planilha ou nenhum ativo com rendimento positivo foi encontrado."
        print(error_msg)
        raise ValueError(error_msg)
    
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
        
        # Se não encontrou, lançar erro
        error_msg = "Erro: Não foi possível encontrar 'Ativos Detratores' na planilha ou nenhum ativo com rendimento negativo foi encontrado."
        print(error_msg)
        raise ValueError(error_msg)
    
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
    
    def get_banker_info(self, client_name, base_path=None):
        """Obtém o nome do banker responsável pelo cliente a partir da aba Base Consolidada"""
        import os
        import pandas as pd
        
        if base_path is None:
            from mmzr_compatibilidade import MMZRCompatibilidade
            base_path, _ = MMZRCompatibilidade.get_planilhas_path()
        
        try:
            # Carregar a aba Base Consolidada da planilha
            excel = pd.ExcelFile(base_path)
            if 'Base Consolidada' in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name='Base Consolidada')
                
                # Limpar espaços em branco nos nomes dos clientes
                df['NomeCompletoCliente'] = df['NomeCompletoCliente'].str.strip() if 'NomeCompletoCliente' in df.columns else None
                df['NomeCliente'] = df['NomeCliente'].str.strip() if 'NomeCliente' in df.columns else None
                
                # Buscar o cliente pelo nome completo ou pelo primeiro nome
                cliente_row = df[(df['NomeCompletoCliente'] == client_name) | (df['NomeCliente'] == client_name)]
                
                if len(cliente_row) > 0:
                    banker = cliente_row['Banker'].iloc[0] if 'Banker' in df.columns and pd.notna(cliente_row['Banker'].iloc[0]) else "Banker"
                    banker_pronome = cliente_row['NomePronomeBanker'].iloc[0] if 'NomePronomeBanker' in df.columns and pd.notna(cliente_row['NomePronomeBanker'].iloc[0]) else banker
                    return banker, banker_pronome
            
            # Se não encontrar, retornar valores padrão
            return "Banker", "o Banker"
        except Exception as e:
            print(f"Erro ao buscar informações do banker: {str(e)}")
            return "Banker", "o Banker"
    
    def get_email_css(self):
        """Retorna o CSS otimizado do email"""
        return """
    <style>
    /* Reset e compatibilidade */
    * {
        box-sizing: border-box;
        -webkit-box-sizing: border-box;
        -moz-box-sizing: border-box;
    }
    
    :root {
        color-scheme: only light !important;
        supported-color-schemes: only light !important;
        forced-color-adjust: none !important;
    }

    html, body {
        background-color: #ffffff !important;
        color: #333333 !important;
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    
    img, a img {
        border: 0;
        height: auto;
        outline: none;
        text-decoration: none;
        max-width: 100%;
        display: block;
    }
    
    /* Layout principal */
    .email-container, .content-wrapper {
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        background-color: #ffffff !important;
    }
    
    /* Tabelas de dados */
    .data-table {
        font-size: 12px !important;
        line-height: 1.3 !important;
        width: 100% !important;
        border-collapse: collapse;
    }
    
    .data-table th, .data-table td {
        font-size: 12px !important;
        padding: 5px 4px !important;
        line-height: 1.2 !important;
        border-bottom: 1px solid #dee2e6 !important;
    }
    
    .data-table th {
        background-color: #f8f9fa !important;
        color: #0D2035 !important;
        font-weight: bold;
    }
    
    .data-table td {
        background-color: #ffffff !important;
    }
    
    /* Headers e seções */
    .portfolio-header h3 {
        font-size: 14px !important;
        margin: 0 !important;
        line-height: 1.2 !important;
        color: #0D2035 !important;
    }
    
    .portfolio-header span {
        font-size: 12px !important;
        line-height: 1.2 !important;
        color: #666666 !important;
    }
    
    /* Listas de ativos */
    .highlight-section, .promoters-section, .detractors-section {
        margin: 6px 0 8px 0 !important;
        padding: 6px 6px 6px 20px !important;
    }
    
    .highlight-section li, .promoters-section li, .detractors-section li {
        font-size: 11px !important;
        margin-bottom: 3px !important;
        line-height: 1.3 !important;
    }
    
    /* Mobile responsivo */
    @media screen and (max-width: 600px) {
        .data-table, .data-table th, .data-table td {
            font-size: 10px !important;
            padding: 3px 2px !important;
            line-height: 1.1 !important;
        }
        
        .data-table th, .data-table td {
            white-space: nowrap;
        }
    }
    </style>"""
    
    def generate_html_email(self, client_name, data_ref, portfolios_data):
        """Gera o HTML completo do email sem bordas cinzas"""
        
        # Configurar mês/ano
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        # Caminho da imagem logo (usar caminho relativo)
        logo_path = os.path.join("recursos_email", "logo-MMZR-azul.png")
        
        # Gerar link para a carta mensal (mês atual)
        carta_mes = self.meses_pt[datetime.now().month].lower()
        carta_link = f"https://www.mmzrfo.com.br/post/carta-mensal-{carta_mes}-{datetime.now().year}"
        
        # Obter informações do banker
        banker, banker_pronome = self.get_banker_info(client_name)
        
        # Criar o texto da observação baseado no banker
        if banker == 'Banker 4':
            obs_text = "<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe em cópia para também receber as informações."
        else:
            obs_text = f"<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e {banker_pronome} em cópia para também receberem as informações."
        
        # HTML completo sem bordas cinzas
        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="color-scheme" content="only light">
    <meta name="supported-color-schemes" content="only light">
    <meta name="theme-color" content="#ffffff">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="format-detection" content="telephone=no">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="x-apple-disable-message-reformatting">
    <meta name="og:title" property="og:title" content="Relatório MMZR Family Office">
    <!--[if mso]>
    <style type="text/css">
        body, table, td, p, a, li, blockquote {{font-family: Arial, Helvetica, sans-serif !important;}}
        .mso-hide {{display: none !important;}}
        table {{border-collapse: collapse !important;}}
    </style>
    <![endif]-->
    <style>
    body {{
        background-color: #ffffff !important;
        color: #333333 !important;
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        min-width: 100% !important;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', sans-serif;
    }}
    
    .email-container {{
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        background-color: #ffffff !important;
    }}
    
    .content-wrapper {{
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        background-color: #ffffff !important;
    }}
    
    .data-table {{
        font-size: 12px !important;
        width: 100% !important;
        border-collapse: collapse;
        background-color: #ffffff !important;
    }}
    
    .data-table th {{
        font-size: 12px !important;
        padding: 5px 4px !important;
        background-color: #f8f9fa !important;
        color: #0D2035 !important;
        border-bottom: 1px solid #dee2e6 !important;
        font-weight: 600;
    }}
    
    .data-table td {{
        font-size: 12px !important;
        padding: 5px 4px !important;
        background-color: #ffffff !important;
        border-bottom: 1px solid #dee2e6 !important;
        color: #333333 !important;
    }}
    
    .portfolio-header {{
        background-color: #0D2035 !important;
        color: #ffffff !important;
        padding: 10px 12px;
    }}
    
    .portfolio-header h3 {{
        margin: 0;
        font-size: 16px;
        color: #ffffff !important;
        font-weight: 500;
    }}
    
    .portfolio-header span {{
        font-size: 14px;
        color: #ffffff !important;
        font-weight: 300;
        opacity: 0.9;
    }}
    
    .performance-header {{
        font-size: 14px;
        color: #0D2035 !important;
        margin: 0 0 10px 0;
        padding-bottom: 5px;
        border-bottom: 1px solid #e0e0e0;
        font-weight: 500;
    }}
    
    .compact-section {{
        margin: 15px 0;
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        overflow: hidden;
        background-color: #ffffff !important;
    }}
    
    .section-content {{
        padding: 15px;
        background-color: #ffffff !important;
    }}
    
    @media screen and (max-width: 600px) {{
        .data-table th,
        .data-table td {{
            font-size: 10px !important;
            padding: 3px 2px !important;
        }}
        
        .portfolio-header h3 {{
            font-size: 14px !important;
        }}
        
        .portfolio-header span {{
            font-size: 12px !important;
            display: block;
            margin-top: 2px;
        }}
        
        .performance-header {{
            font-size: 12px !important;
        }}
        
        .section-content {{
            padding: 10px !important;
        }}
    }}
    
    @media screen and (max-width: 480px) {{
        .data-table th,
        .data-table td {{
            font-size: 9px !important;
            padding: 2px 1px !important;
        }}
        
        .portfolio-header h3 {{
            font-size: 12px !important;
        }}
        
        .performance-header {{
            font-size: 11px !important;
        }}
    }}
    </style>
</head>
<body>
    <div class="email-container">
        <table class="content-wrapper" cellpadding="0" cellspacing="0" style="width: 100%; background-color: #ffffff !important; border-collapse: collapse;">
            <!-- Header -->
            <tr>
                <td style="background-color: #0D2035 !important; color: #ffffff !important; padding: 12px; text-align: center;">
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <td style="text-align: left; vertical-align: middle; width: 60px;">
                                <img src="{logo_path}" alt="MMZR Family Office" width="50" height="40" style="display: block; border: 0;">
                            </td>
                            <td style="text-align: left; vertical-align: middle; padding-left: 10px;">
                                <h1 style="margin: 0; font-size: 18px; color: #ffffff !important; font-weight: 500;">MMZR Family Office</h1>
                                <p style="margin: 0; font-size: 14px; color: #ffffff !important;">Relatório Mensal - {mes} {ano}</p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- Content -->
            <tr>
                <td class="section-content" style="padding: 15px; background-color: #ffffff !important;">
                    <h2 style="font-size: 16px; color: #0D2035 !important; margin-bottom: 10px; margin-top: 0; font-weight: 500;">Olá {client_name},</h2>
                    <p style="margin-top: 0; margin-bottom: 15px; color: #333333 !important; font-size: 14px; line-height: 1.4;">
                        Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.
                    </p>"""
        
        # Adicionar cada carteira
        for portfolio in portfolios_data:
            html += self.generate_portfolio_section(portfolio)
        
        # Adicionar link para a carta mensal
        html += f"""
                    <!-- Link para a carta mensal -->
                    <div style="margin-top: 20px; text-align: center;">
                        <a href="{carta_link}" style="display: inline-block; background-color: #0D2035; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: 500; font-size: 14px;">Confira nossa carta completa: Carta {mes} {ano}</a>
                    </div>
                    
                    <!-- Disclaimer -->
                    <div style="margin-top: 20px; padding: 12px; background-color: #f8f9fa; border-radius: 5px;">
                        <p style="margin: 0 0 8px 0; color: #555555 !important; font-size: 11px; font-style: italic; line-height: 1.3;">
                            <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                        </p>
                        <p style="margin: 0; color: #555555 !important; font-size: 11px; font-style: italic; line-height: 1.3;">
                            {obs_text}
                        </p>
                    </div>

                    <!-- Principais indicadores -->
                    <div style="margin-top: 15px; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <p style="margin: 0 0 5px 0; font-weight: bold; color: #333333 !important; font-size: 12px;">Principais indicadores:</p>
                        <p style="margin: 0; color: #555555 !important; font-size: 10px; font-style: italic; line-height: 1.3;">
                            Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                            Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                        </p>
                    </div>
                </td>
            </tr>
            
            <!-- Footer -->
            <tr>
                <td style="background-color: #f8f9fa !important; padding: 12px; text-align: center;">
                    <p style="margin: 0 0 4px 0; color: #666666 !important; font-size: 11px;">MMZR Family Office | Gestão de Patrimônio</p>
                    <p style="margin: 0; color: #666666 !important; font-size: 11px;">© {ano} MMZR Family Office. Todos os direitos reservados.</p>
                </td>
            </tr>
        </table>
    </div>
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
                    <div class="compact-section" style="margin: 8px 0; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden; background-color: #ffffff !important;">
                        <div class="portfolio-header" style="background-color: #0D2035 !important; color: #ffffff !important; padding: 6px 10px;">
                            <h3 style="margin: 0; font-size: 14px; color: #ffffff !important; font-weight: 500;">{name} <span style="font-size: 12px; color: #ffffff !important; font-weight: 300; opacity: 0.9; margin-left: 4px;">| {portfolio_type}</span></h3>
                        </div>
                        <div class="section-content" style="padding: 8px; background-color: #ffffff !important;">
                            {self.generate_performance_table(performance_data, retorno_financeiro)}
                            
                            {self.generate_highlight_strategies(estrategias_destaque)}
                            
                            {self.generate_promoter_assets(ativos_promotores)}
                            
                            {self.generate_detractor_assets(ativos_detratores)}
                        </div>
                    </div>"""
        
        return html
    
    def generate_performance_table(self, performance_data, retorno_financeiro=None):
        """Gera a tabela HTML de performance otimizada"""
        
        # Filtrar apenas os períodos necessários (Mês atual e No ano)
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
        
        html = f"""
                            <h4 class="performance-header" style="font-size: 12px; color: #0D2035 !important; margin: 0 0 5px 0; padding-bottom: 3px; border-bottom: 1px solid #e0e0e0; font-weight: 500;">Performance</h4>
                            <div style="overflow-x: auto; -webkit-overflow-scrolling: touch; width: 100%;">
                                <table class="data-table" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 6px; background-color: #ffffff !important;">
                                    <thead>
                                        <tr>
                                            <th style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 3px 2px; text-align: left; border-bottom: 1px solid #dee2e6 !important; font-size: 10px;">Período</th>
                                            <th style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; font-size: 10px;">Carteira</th>
                                            <th style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; font-size: 10px;">Benchmark</th>
                                            <th style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; font-size: 10px;">Carteira vs. Benchmark</th>
                                        </tr>
                                    </thead>
                                    <tbody>"""
        
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
                                            <td style="padding: 3px 2px; text-align: left; border-bottom: 1px solid #dee2e6 !important; background-color: #ffffff !important; color: #333333 !important; font-size: 10px;">{periodo}</td>
                                            <td style="padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {carteira_color}; font-weight: 500; background-color: #ffffff !important; font-size: 10px;">{self.format_percentage(carteira)}</td>
                                            <td style="padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; background-color: #ffffff !important; color: #333333 !important; font-size: 10px;">{self.format_percentage(benchmark)}</td>
                                            <td style="padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {diferenca_color}; font-weight: 500; background-color: #ffffff !important; font-size: 10px;">{self.format_percentage(diferenca).replace('%', ' p.p.')}</td>
                                        </tr>"""
        
        # Adicionar linha de retorno financeiro se disponível
        if retorno_financeiro is not None:
            color = "#28a745" if retorno_financeiro > 0 else "#dc3545" if retorno_financeiro < 0 else "#333333"
            html += f"""
                                        <tr>
                                            <td style="padding: 3px 2px; text-align: left; border-bottom: 1px solid #dee2e6 !important; font-weight: 500; background-color: #ffffff !important; color: #333333 !important; font-size: 10px;">Retorno Financeiro:</td>
                                            <td style="padding: 3px 2px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {color}; font-weight: 500; background-color: #ffffff !important; font-size: 10px;" colspan="3">{self.format_currency(retorno_financeiro)}</td>
                                        </tr>"""
        
        html += """
                                    </tbody>
                                </table>
                            </div>"""
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
    
    def generate_highlight_strategies(self, estrategias):
        """Gera a seção de estratégias de destaque"""
        
        if not estrategias:
            return ""
        
        html = """
                            <h4 style="font-size: 12px; color: #0D2035 !important; margin: 8px 0 4px 0; padding-bottom: 3px; border-bottom: 1px solid #e0e0e0; font-weight: 500;">Estratégias de Destaque</h4>
                            <ul style="margin: 0 0 8px 0; padding: 4px 4px 4px 16px; background-color: #f8f9fa !important; border-radius: 3px; color: #333333 !important;">"""
        
        for estrategia in estrategias:
            html += f"""
                                <li style="margin-bottom: 2px; font-size: 10px; color: #333333 !important; line-height: 1.2;">{estrategia}</li>"""
        
        html += """
                            </ul>"""
        return html
    
    def generate_promoter_assets(self, ativos):
        """Gera a seção de ativos promotores"""
        
        if not ativos:
            return ""
        
        html = """
                            <h4 style="font-size: 12px; color: #0D2035 !important; margin: 8px 0 4px 0; padding-bottom: 3px; border-bottom: 1px solid #e0e0e0; font-weight: 500;">Ativos Promotores</h4>
                            <ul style="margin: 0 0 8px 0; padding: 4px 4px 4px 16px; background-color: #e8f5e9 !important; border-radius: 3px; color: #2e7d32 !important;">"""
        
        for ativo in ativos:
            # Adicionar o símbolo "+" antes da porcentagem se for um valor positivo
            ativo_formatado = ativo
            percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', ativo)
            if percentage_match:
                percentage_str = percentage_match.group(1).replace(',', '.')
                try:
                    percentage = float(percentage_str)
                    if percentage > 0 and not percentage_str.startswith('+'):
                        ativo_formatado = ativo.replace(f"({percentage_str}%)", f"(+{percentage_str}%)")
                except ValueError:
                    pass
                    
            html += f"""
                                <li style="margin-bottom: 2px; font-size: 10px; color: #2e7d32 !important; line-height: 1.2;">{ativo_formatado}</li>"""
        
        html += """
                            </ul>"""
        return html
    
    def generate_detractor_assets(self, ativos):
        """Gera a seção de ativos detratores"""
        
        if not ativos:
            return ""
        
        html = """
                            <h4 style="font-size: 12px; color: #0D2035 !important; margin: 8px 0 4px 0; padding-bottom: 3px; border-bottom: 1px solid #e0e0e0; font-weight: 500;">Ativos Detratores</h4>
                            <ul style="margin: 0 0 8px 0; padding: 4px 4px 4px 16px; background-color: #ffebee !important; border-radius: 3px; color: #c62828 !important;">"""
        
        for ativo in ativos:
            html += f"""
                                <li style="margin-bottom: 2px; font-size: 10px; color: #c62828 !important; line-height: 1.2;">{ativo}</li>"""
        
        html += """
                            </ul>"""
        return html
    
    def save_email_to_file(self, html_content, client_name, output_path=None):
        """Salva o conteúdo HTML do e-mail em um arquivo e copia recursos necessários"""
        # Remover caracteres inválidos para nome de arquivo
        safe_client_name = "".join([c if c.isalnum() or c in [' ', '_'] else '_' for c in client_name])
        safe_client_name = safe_client_name.replace(' ', '_')
        
        # Data atual para nome do arquivo
        date_str = datetime.now().strftime("%Y%m%d")
        
        # Caminho de saída
        if not output_path:
            filename = f"relatorio_mensal_{safe_client_name}_{date_str}.html"
            output_path = filename
        
        # Diretório para recursos
        resources_dir = "recursos_email"
        if not os.path.exists(resources_dir):
            os.makedirs(resources_dir)
        
        # Copiar o logo para o diretório de recursos
        logo_src = os.path.join("documentos", "img", "logo-MMZR-azul.png")
        logo_dest = os.path.join(resources_dir, "logo-MMZR-azul.png")
        
        # Se o logo de origem existe, copiar para o destino
        if os.path.exists(logo_src):
            import shutil
            shutil.copy2(logo_src, logo_dest)
            print(f"Logo copiado para {logo_dest}")
            
            # Atualizar o caminho da imagem no HTML para usar um caminho relativo
            # em vez de absoluto para melhor compatibilidade com clientes de email
            html_content = html_content.replace(os.path.abspath(logo_src), logo_dest)
        
        # Verificar se ainda existem caminhos absolutos no HTML e substituí-los
        import re
        # Substituir caminhos absolutos em src de imagens
        absolute_paths = re.findall(r'src="(/[^"]+logo-MMZR-azul\.png)"', html_content)
        for path in absolute_paths:
            html_content = html_content.replace(path, logo_dest)
            
        # Substituir caminhos absolutos no VML para o Outlook
        vml_paths = re.findall(r'v:fill type="frame" src="([^"]+logo-MMZR-azul\.png)"', html_content)
        for path in vml_paths:
            html_content = html_content.replace(f'v:fill type="frame" src="{path}"', f'v:fill type="frame" src="{logo_dest}"')
        
        # Salvar o arquivo
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Relatório salvo em: {output_path}")
        return output_path
