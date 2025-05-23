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
    
    def generate_html_email(self, client_name, data_ref, portfolios_data):
        """Gera o HTML completo do email"""
        
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
            # Se o banker é o Banker 4 (Felipe), usar texto singular sem duplicação
            obs_text = "<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe em cópia para também receber as informações."
        else:
            # Se o banker não é o Banker 4, usar texto plural com os dois nomes
            obs_text = f"<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e {banker_pronome} em cópia para também receberem as informações."
        
        # HTML Header
        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="only light">
    <meta name="supported-color-schemes" content="only light">
    <meta name="theme-color" content="#ffffff">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="format-detection" content="telephone=no">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="x-apple-disable-message-reformatting">
    <meta name="og:title" property="og:title" content="Relatório MMZR Family Office">
    <!-- Gmail mobile fix -->
    <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
    <!--[if mso]>
    <noscript>
    <xml>
        <o:OfficeDocumentSettings>
            <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
    </xml>
    </noscript>
    <![endif]-->
    <!--[if mso]>
    <style type="text/css">
        body, table, td, p, a, li, blockquote {{font-family: Arial, Helvetica, sans-serif !important;}}
        .mso-hide {{display: none !important;}}
        table {{border-collapse: collapse !important;}}
        .mso-text-color {{mso-style-textfill-fill-color: #333333 !important;}}
        .mso-text-bg {{mso-style-textfill-fill-bgcolor: #ffffff !important;}}
    </style>
    <![endif]-->
    <style>
    /* Estilos para forçar modo claro em dispositivos com tema escuro */
    :root {{
        color-scheme: only light !important;
        supported-color-schemes: only light !important;
        forced-color-adjust: none !important;
    }}

    html, body {{
        background-color: #f4f4f4 !important;
        color: #333333 !important;
    }}
    
    /* Regras para normalizar elementos */
    img, a img {{
        border: 0;
        height: auto;
        outline: none;
        text-decoration: none;
        max-width: 100%;
    }}
    
    /* Fix para Gmail iOS */
    @supports (-webkit-overflow-scrolling: touch) {{
        .header-logo-container {{
            max-height: 60px !important;
        }}
        .header-logo {{
            max-height: 60px !important;
        }}
    }}
    
    /* Regra principal para forçar cores no tema escuro */
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
        
        /* Para Gmail em dispositivos Android */
        u + .body .body-wrapper {{
            background-color: #f4f4f4 !important;
        }}
        
        /* Para aplicativos de email do iOS */
        [data-ogsc] .body-wrapper,
        [data-ogsb] .body-wrapper {{
            background-color: #f4f4f4 !important;
        }}
        
        [data-ogsc] .content-wrapper,
        [data-ogsb] .content-wrapper {{
            background-color: #ffffff !important;
        }}
    }}
    
    /* Suportes específicos para diferentes clientes de email */
    /* Yahoo Mail */
    @media yahoo {{
        .content-wrapper {{
            background-color: #ffffff !important;
        }}
        .header-bg {{
            background-color: #0D2035 !important;
        }}
    }}
    
    /* Outlook.com */
    [class~=x_body] {{
        background-color: #f4f4f4 !important;
    }}
    
    /* Samsung Email */
    #MessageViewBody, #MessageWebViewDiv {{
        background-color: #f4f4f4 !important;
    }}
    
    /* Regras de responsividade para todos os dispositivos */
    @media screen and (max-width: 600px) {{
        .content-wrapper {{
            width: 100% !important;
            max-width: 100% !important;
        }}
        
        td {{
            padding: 8px !important;
        }}
        
        .mobile-full-width {{
            width: 100% !important;
        }}
        
        .mobile-text-center {{
            text-align: center !important;
        }}
        
        .mobile-smaller-text {{
            font-size: 14px !important;
        }}
        
        /* Ajustes específicos para o header */
        .header-logo-container {{
            width: 60px !important;
            height: 50px !important;
            max-width: 60px !important;
        }}
        
        .header-logo {{
            width: 60px !important;
            height: auto !important;
            max-height: 50px !important;
            min-width: 50px !important;
        }}
        
        .header-text-main {{
            font-size: 16px !important;
        }}
        
        .header-text-sub {{
            font-size: 12px !important;
        }}
    }}
    
    /* Correções específicas para Android */
    @media screen and (max-width: 480px) {{
        u + .body .header-bg {{
            padding: 6px !important;
        }}
        
        u + .body .header-logo-container {{
            width: 55px !important;
            height: 46px !important;
            min-width: 50px !important;
        }}
    }}
    
    /* Correções específicas para iOS */
    @media screen and (max-device-width: 480px) {{
        .iOS-header {{
            font-size: 16px !important;
        }}
        
        .iOS-logo {{
            min-width: 55px !important;
            max-width: 60px !important;
        }}
    }}
    
    /* Regras para telas maiores */
    @media screen and (min-width: 601px) {{
        .header-bg {{
            padding: 8px 10px !important;
            min-height: 0 !important;
            height: auto !important;
        }}
        
        .content-wrapper {{
            max-width: 600px !important;
        }}
        
        .header-logo-container {{
            width: 70px !important;
            height: 58px !important;
            max-height: 58px !important;
        }}
        
        .header-text-main {{
            font-size: 17px !important;
        }}
        
        .header-text-sub {{
            font-size: 13px !important;
        }}
    }}
    </style>
</head>
<body class="body-wrapper" style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', sans-serif; line-height: 1.4; color: #333333 !important; background-color: #f4f4f4 !important; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
    <!--[if mso | IE]>
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#f4f4f4">
    <tr>
    <td align="center">
    <![endif]-->
    <div style="background-color: #f4f4f4 !important; width: 100%;">
        <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; border: 0; border-spacing: 0; background-color: #f4f4f4 !important; margin: 0; padding: 0;">
            <tr>
                <td align="center" style="padding: 0;">
                    <table role="presentation" class="content-wrapper" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 600px; border-collapse: collapse; border: 0; border-spacing: 0; text-align: left; background-color: #ffffff !important; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin: 0 auto;">
                        <!-- Header -->
                        <tr>
                            <td style="padding: 0;">
                                <table role="presentation" class="header-bg" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; background-color: #0D2035 !important;">
                                    <tr>
                                        <td style="padding: 6px 10px;">
                                            <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                                                <tr>
                                                    <td style="text-align: center; vertical-align: middle; width: 70px;" class="mobile-full-width">
                                                        <!-- Logo com fallback text -->
                                                        <div class="header-logo-container iOS-logo" style="display: inline-block; width: 70px; height: 58px; background-color: #0D2035; overflow: hidden; max-width: 70px;">
                                                            <!--[if mso]>
                                                            <v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:70px;height:58px;">
                                                                <v:fill type="frame" src="{logo_path}" color="#0D2035" />
                                                                <v:textbox style="mso-fit-shape-to-text:true" inset="0,0,0,0">
                                                                    <center style="font-family:Arial,sans-serif;font-size:18px;color:#ffffff;font-weight:bold;">MMZR</center>
                                                                </v:textbox>
                                                            </v:rect>
                                                            <![endif]-->
                                                            <!--[if !mso]><!-->
                                                            <img class="header-logo" src="{logo_path}" alt="MMZR Family Office" width="70" height="58" style="display: inline-block; border: 0; max-width: 100%; width: auto; height: auto; max-height: 58px;">
                                                            <!--<![endif]-->
                                                        </div>
                                                    </td>
                                                    <td style="text-align: left; vertical-align: middle; padding-left: 8px;" class="mobile-text-center">
                                                        <p class="header-text header-text-main iOS-header" style="margin: 0; font-size: 17px; color: #ffffff !important; line-height: 1.2;">MMZR Family Office</p>
                                                        <p class="header-text header-text-sub mobile-smaller-text" style="margin: 0; font-size: 13px; color: #ffffff !important; line-height: 1.2;">Relatório Mensal de Performance - {mes} de {ano}</p>
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
                            <td class="section-bg" style="padding: 20px; background-color: #ffffff !important;">
                                <h2 style="font-size: 15px; color: #0D2035 !important; margin-bottom: 12px; margin-top: 0;">Olá {client_name},</h2>
                                
                                <p style="margin-top: 0; margin-bottom: 9px; color: #333333 !important;">Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.</p>"""
        
        # Adicionar cada carteira
        for portfolio in portfolios_data:
            html += self.generate_portfolio_section(portfolio)
        
        # Adicionar link para a carta mensal
        html += f"""
                                <!-- Link para a carta mensal -->
                                <div style="margin-top: 25px; text-align: center;">
                                    <a href="{carta_link}" style="display: inline-block; background-color: #0D2035; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: 500; font-size: 14px;">Confira nossa carta completa: Carta Mensal - {mes} {ano}</a>
                                </div>
"""
        
        # Disclaimer e observações
        html += f"""
                                <!-- Disclaimer -->
                                <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; margin-top: 25px; border-collapse: collapse;">
                                    <tr>
                                        <td style="padding: 10px; border-radius: 4px;">
                                            <p style="margin: 0 0 10px 0; color: #555555 !important; font-size: 12px; font-style: italic;">
                                                <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                                            </p>
                                            <p style="margin: 0; color: #555555 !important; font-size: 12px; font-style: italic;">
                                                {obs_text}
                                            </p>
                                        </td>
                                    </tr>
                                </table>

                                <!-- Principais indicadores -->
                                <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; margin-top: 15px; border-collapse: collapse;">
                                    <tr>
                                        <td style="padding: 0;">
                                            <p style="margin: 0 0 5px 0; font-weight: bold; color: #333333 !important; font-size: 13px;">Principais indicadores:</p>
                                            <p style="margin: 0; color: #555555 !important; font-size: 12px; font-style: italic;">
                                                Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                                                Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                                            </p>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        
                        <!-- Footer -->
                        <tr>
                            <td style="background-color: #f8f9fa !important; padding: 12px 20px; text-align: center;">
                                <p style="margin: 0 0 3px 0; color: #666666 !important; font-size: 11px;">MMZR Family Office | Gestão de Patrimônio</p>
                                <p style="margin: 0 0 3px 0; color: #666666 !important; font-size: 11px;">Este é um email automático. Por favor, não responda.</p>
                                <p style="margin: 0; color: #666666 !important; font-size: 11px;">© {ano} MMZR Family Office. Todos os direitos reservados.</p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    <!--[if mso | IE]>
    </td>
    </tr>
    </table>
    <![endif]-->
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
                                <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; margin: 20px 0 0 0; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); background-color: #ffffff !important;">
                                    <tr>
                                        <td class="header-bg portfolio-header" style="background-color: #0D2035 !important; color: #ffffff !important; padding: 8px 15px;">
                                            <h3 style="margin: 0; font-size: 16px; font-weight: 500; color: #ffffff !important;">{name} <span style="font-weight: 300; font-size: 13px; margin-left: 8px; opacity: 0.8; color: #ffffff !important;">| {portfolio_type}</span></h3>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="section-bg" style="padding: 15px; background-color: #ffffff !important; color: #333333 !important;">
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
                                            <h4 class="performance-header" style="font-size: 18px; color: #0D2035 !important; margin: 0 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0 !important; padding-bottom: 8px;">Performance</h4>
                                            <div style="overflow-x: auto; -webkit-overflow-scrolling: touch;">
                                            <table role="presentation" class="data-table" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 15px; background-color: #ffffff !important;">
                                                <thead>
                                                    <tr>
                                                        <th class="table-header" style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6 !important;">Período</th>
                                                        <th class="table-header" style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important;">Carteira</th>
                                                        <th class="table-header" style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important;">Benchmark</th>
                                                        <th class="table-header" style="background-color: #f8f9fa !important; color: #0D2035 !important; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important;">Carteira vs. Benchmark</th>
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
            carteira_color = "#28a745 !important" if carteira > 0 else "#dc3545 !important" if carteira < 0 else "#333333 !important"
            diferenca_color = "#28a745 !important" if diferenca > 0 else "#dc3545 !important" if diferenca < 0 else "#333333 !important"
            
            html += f"""
                                                    <tr>
                                                        <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6 !important; background-color: #ffffff !important; color: #333333 !important;">{periodo}</td>
                                                        <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {carteira_color}; font-weight: 500; background-color: #ffffff !important;">{self.format_percentage(carteira)}</td>
                                                        <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important; background-color: #ffffff !important; color: #333333 !important;">{self.format_percentage(benchmark)}</td>
                                                        <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {diferenca_color}; font-weight: 500; background-color: #ffffff !important;">{self.format_percentage(diferenca).replace('%', ' p.p.')}</td>
                                                    </tr>
"""
        
        # Adicionar linha de retorno financeiro se disponível
        if retorno_financeiro is not None:
            color = "#28a745 !important" if retorno_financeiro > 0 else "#dc3545 !important" if retorno_financeiro < 0 else "#333333 !important"
            html += f"""
                                                    <tr>
                                                        <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6 !important; font-weight: 500; background-color: #ffffff !important; color: #333333 !important;">Retorno Financeiro:</td>
                                                        <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6 !important; color: {color}; font-weight: 500; background-color: #ffffff !important;" colspan="3">{self.format_currency(retorno_financeiro)}</td>
                                                    </tr>
"""
        
        html += """
                                                </tbody>
                                            </table>
                                            </div>
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
                                            <h4 class="performance-header" style="font-size: 18px; color: #0D2035 !important; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0 !important; padding-bottom: 8px;">Estratégias de Destaque</h4>
                                            <ul class="highlight-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #f8f9fa !important; border-radius: 5px; color: #333333 !important;">
"""
        
        for estrategia in estrategias:
            html += f"""
                                                <li style="margin-bottom: 6px; font-size: 13px; color: #333333 !important;">{estrategia}</li>
"""
        
        html += """
                                            </ul>
"""
        return html
    
    def generate_promoter_assets_section(self, ativos):
        """Gera a seção de ativos promotores"""
        
        html = """
                                            <h4 class="performance-header" style="font-size: 18px; color: #0D2035 !important; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0 !important; padding-bottom: 8px;">Ativos Promotores</h4>
                                            <ul class="promoters-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #e8f5e9 !important; border-radius: 5px; color: #2e7d32 !important;">
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
                                                <li style="margin-bottom: 6px; font-size: 13px; color: #2e7d32 !important;">{ativo_formatado}</li>
"""
        
        html += """
                                            </ul>
"""
        return html
    
    def generate_detractor_assets_section(self, ativos):
        """Gera a seção de ativos detratores"""
        
        html = """
                                            <h4 class="performance-header" style="font-size: 18px; color: #0D2035 !important; margin: 20px 0 12px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0 !important; padding-bottom: 8px;">Ativos Detratores</h4>
                                            <ul class="detractors-section" style="margin: 8px 0 15px 0; padding: 10px 10px 10px 30px; background-color: #ffebee !important; border-radius: 5px; color: #c62828 !important;">
"""
        
        for ativo in ativos:
            html += f"""
                                                <li style="margin-bottom: 6px; font-size: 13px; color: #c62828 !important;">{ativo}</li>
"""
        
        html += """
                                            </ul>
"""
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
        
        # Garantir que o diretório de recursos existe
        resources_dir = "recursos_email"
        if not os.path.exists(resources_dir):
            os.makedirs(resources_dir)
                
        # Garantir que a imagem do logo foi copiada antes de gerar o HTML
        logo_src = os.path.join("documentos", "img", "logo-MMZR-azul.png")
        logo_dest = os.path.join(resources_dir, "logo-MMZR-azul.png")
        
        if os.path.exists(logo_src):
            import shutil
            shutil.copy2(logo_src, logo_dest)
            print(f"Logo copiado para {logo_dest}")
        
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
                error_msg = f"Erro: Aba '{sheet_name}' não encontrada no Excel. O relatório não pode ser gerado."
                print(error_msg)
                raise ValueError(error_msg)
        
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