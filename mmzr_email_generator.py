import os
import pandas as pd
from datetime import datetime
import re
import base64

class MMZREmailGenerator:
    """Gerador de emails HTML para MMZR Family Office"""
    
    def __init__(self):
        """Inicializa o gerador"""
        self.meses_pt = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
    
    def generate_email_subject(self, data_ref=None):
        """Gera o assunto do email"""
        if data_ref is None:
            data_ref = datetime.now()
        
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        return f"MMZR Family Office | Desempenho {mes} de {ano}"
    
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
        """Obtém informações do banker responsável pelo cliente"""
        if base_path is None:
            from mmzr_compatibilidade import MMZRCompatibilidade
            base_path, _ = MMZRCompatibilidade.get_planilhas_path()
        
        try:
            excel = pd.ExcelFile(base_path)
            if 'Base Consolidada' in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name='Base Consolidada')
                
                df['NomeCompletoCliente'] = df['NomeCompletoCliente'].str.strip() if 'NomeCompletoCliente' in df.columns else None
                df['NomeCliente'] = df['NomeCliente'].str.strip() if 'NomeCliente' in df.columns else None
                
                cliente_row = df[(df['NomeCompletoCliente'] == client_name) | (df['NomeCliente'] == client_name)]
                
                if len(cliente_row) > 0:
                    banker = cliente_row['Banker'].iloc[0] if 'Banker' in df.columns and pd.notna(cliente_row['Banker'].iloc[0]) else "Banker"
                    banker_pronome = cliente_row['NomePronomeBanker'].iloc[0] if 'NomePronomeBanker' in df.columns and pd.notna(cliente_row['NomePronomeBanker'].iloc[0]) else banker
                    return banker, banker_pronome
            
            return "Banker", "o Banker"
        except Exception:
            return "Banker", "o Banker"
    
    def get_logo_base64(self):
        """Converte a logo em base64"""
        # Lista de caminhos possíveis para a logo
        possible_paths = [
            os.path.join("recursos_email", "logo-MMZR-azul.png"),
            os.path.join("documentos", "img", "logo-MMZR-azul.png"),
            "logo-MMZR-azul.png"
        ]
        
        for logo_path in possible_paths:
            try:
                if os.path.exists(logo_path):
                    with open(logo_path, "rb") as image_file:
                        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                        return f"data:image/png;base64,{encoded_string}"
            except Exception as e:
                print(f"Erro ao tentar carregar logo de {logo_path}: {e}")
                continue
        
        print("AVISO: Logo não encontrada em nenhum dos caminhos esperados")
        return None
    
    def generate_html_email(self, client_name, data_ref, portfolios_data):
        """Gera o HTML completo do email compatível com Outlook"""
        
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        logo_base64 = self.get_logo_base64()
        logo_img = f'<img src="{logo_base64}" alt="MMZR Family Office" width="80" height="64" style="display: block; border: 0; max-width: 80px;">' if logo_base64 else '<span style="color: #ffffff; font-weight: bold;">MMZR</span>'
        
        carta_mes = self.meses_pt[datetime.now().month].lower()
        carta_link = f"https://www.mmzrfo.com.br/post/carta-mensal-{carta_mes}-{datetime.now().year}"
        
        banker, banker_pronome = self.get_banker_info(client_name)
        
        if banker == 'Banker 4':
            obs_text = "<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe em cópia para também receber as informações."
        else:
            obs_text = f"<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e {banker_pronome} em cópia para também receberem as informações."
        
        # Coletar comentários das carteiras
        comentarios_carteiras = []
        for portfolio in portfolios_data:
            comentario = portfolio.get('data', {}).get('comentarios', None)
            if comentario:
                comentarios_carteiras.append(f"<strong>Comentário:</strong> {comentario}")
        
        comentarios_html = ""
        if comentarios_carteiras:
            comentarios_html = "<br>".join(comentarios_carteiras)
            comentarios_html = f"<p style=\"margin: 8px 0 0 0; color: #555555; font-size: 12px; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;\">{comentarios_html}</p>"
        
        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <!--[if mso]>
    <noscript>
        <xml>
            <o:OfficeDocumentSettings>
                <o:AllowPNG/>
                <o:PixelsPerInch>96</o:PixelsPerInch>
            </o:OfficeDocumentSettings>
        </xml>
    </noscript>
    <![endif]-->
    <style type="text/css">
        /* Design tokens para espaçamentos padronizados */
        /* xs: 4px, sm: 8px, md: 12px, lg: 16px, xl: 20px, xxl: 24px */
        
        /* Outlook-specific resets */
        table {{ border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}
        img {{ border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }}
        p {{ margin: 0; }}
        
        /* Font fallbacks for Outlook */
        .fallback-font {{ font-family: Arial, Helvetica, sans-serif; }}
        
        /* Padronização de espaçamentos */
        .content-spacing {{ padding: 16px; }}
        .section-spacing {{ margin: 20px 0; }}
        .header-spacing {{ margin: 8px 0 6px 0; }}
        .list-spacing {{ padding-left: 16px; margin: 0; }}
        .cell-spacing {{ padding: 8px; }}
        
        /* Outlook button fix */
        .button-link {{ 
            display: inline-block; 
            text-decoration: none; 
            background-color: #0D2035; 
            color: #ffffff; 
            padding: 12px 24px; 
            border-radius: 4px; 
            font-weight: bold; 
            font-size: 14px; 
            font-family: Arial, Helvetica, sans-serif;
            border: none;
        }}
        
        /* Remove box-shadow for Outlook */
        <!--[if mso]>
        .no-shadow {{ box-shadow: none !important; }}
        <![endif]-->
    </style>
</head>
<body style="margin: 0; padding: 0; background-color: #f5f5f5; color: #333333; font-family: Arial, Helvetica, sans-serif; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
    <!-- Container table with fixed width -->
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse; background-color: #f5f5f5; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
        <tr>
            <td style="padding: 20px 10px;">
                <!-- Main email content table with max width -->
                <table cellpadding="0" cellspacing="0" border="0" width="680" style="max-width: 680px; margin: 0; border-collapse: collapse; background-color: #ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.1); mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                    <!-- Header -->
                    <tr>
                        <td style="background-color: #0D2035; padding: 16px; text-align: center;">
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <tr>
                                    <td style="text-align: left; vertical-align: middle; width: 90px;">
                                        {logo_img}
                                    </td>
                                    <td style="text-align: left; vertical-align: middle; padding-left: 16px;">
                                        <h1 style="margin: 0; font-size: 20px; color: #ffffff; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">MMZR Family Office</h1>
                                        <p style="margin: 4px 0 0 0; font-size: 16px; color: #ffffff; font-family: Arial, Helvetica, sans-serif;">Relatório Mensal - {mes} {ano}</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td style="padding: 16px; background-color: #ffffff;">
                            <p style="margin: 0 0 12px 0; font-size: 14px; color: #333333; font-family: Arial, Helvetica, sans-serif;">
                                Olá {client_name},
                            </p>
                            
                            <p style="margin: 0 0 16px 0; font-size: 14px; color: #333333; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">
                                Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.
                            </p>"""
        
        # Adicionar seções das carteiras
        for portfolio in portfolios_data:
            html += self.generate_portfolio_section_outlook_compatible(portfolio)
        
        # Observações finais
        html += f"""
                 <!-- Observações finais -->
                 <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 16px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                     <tr>
                         <td style="padding: 12px; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                             <p style="margin: 0 0 8px 0; color: #555555; font-size: 12px; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">
                                 <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                             </p>
                             <p style="margin: 0; color: #555555; font-size: 11px; font-style: italic; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">
                                 {obs_text}
                             </p>
                             {comentarios_html}
                         </td>
                     </tr>
                 </table>

                 <!-- Principais indicadores -->
                 <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 12px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                     <tr>
                         <td style="padding: 12px; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                             <p style="margin: 0 0 4px 0; font-weight: bold; color: #333333; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Principais indicadores:</p>
                             <p style="margin: 0; color: #555555; font-size: 10px; font-style: italic; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">
                                 Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                                 Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                             </p>
                         </td>
                     </tr>
                 </table>
                 
                 <!-- Link para carta mensal -->
                 <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 16px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                     <tr>
                         <td style="text-align: left;">
                             <table cellpadding="0" cellspacing="0" border="0" style="margin: 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                 <tr>
                                     <td style="background-color: #0D2035; border: 1px solid #0D2035; padding: 12px 24px; text-align: center;">
                                         <a href="{carta_link}" target="_blank" style="color: #ffffff; text-decoration: none; font-weight: bold; font-size: 14px; font-family: Arial, Helvetica, sans-serif; display: block;">Confira nossa carta completa: Carta {mes} {ano}</a>
                                     </td>
                                 </tr>
                             </table>
                         </td>
                     </tr>
                 </table>
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 12px; text-align: center; border-top: 1px solid #e9ecef;">
                            <p style="margin: 0; font-size: 11px; color: #666666; font-family: Arial, Helvetica, sans-serif;">
                                MMZR Family Office | Gestão de Patrimônios
                            </p>
                            <p style="margin: 4px 0 0 0; font-size: 10px; color: #888888; font-family: Arial, Helvetica, sans-serif;">
                                © 2025 MMZR Family Office. Todos os direitos reservados.
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""
        
        return html
    
    def generate_portfolio_section_outlook_compatible(self, portfolio):
        """Gera a seção de carteira compatível com Outlook"""
        name = portfolio['name']
        portfolio_type = portfolio['type']
        data = portfolio['data']
        
        # Gerar performance table
        performance_table = self.generate_performance_table_outlook_compatible(
            data['performance'], 
            data.get('retorno_financeiro', 0)
        )
        
        # Gerar estratégias de destaque
        estrategias_html = self.generate_highlight_strategies_outlook_compatible(
            data.get('estrategias_destaque', [])
        )
        
        # Gerar ativos promotores
        promotores_html = self.generate_promoter_assets_outlook_compatible(
            data.get('ativos_promotores', [])
        )
        
        # Gerar ativos detratores
        detratores_html = self.generate_detractor_assets_outlook_compatible(
            data.get('ativos_detratores', [])
        )
        
        return f"""
                <!-- Carteira: {name} -->
                <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 20px 0; border-collapse: collapse; background-color: #ffffff; border: 1px solid #e0e0e0; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                    <!-- Header da carteira -->
                    <tr>
                        <td style="background-color: #0D2035; color: #ffffff; padding: 8px;">
                            <h3 style="margin: 0; font-size: 16px; color: #ffffff; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">{name}</h3>
                            <span style="font-size: 14px; color: #ffffff; font-family: Arial, Helvetica, sans-serif; opacity: 0.9;">{portfolio_type}</span>
                        </td>
                    </tr>
                    <!-- Conteúdo da carteira -->
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff;">
                            {performance_table}
                            {estrategias_html}
                            {promotores_html}
                            {detratores_html}
                        </td>
                    </tr>
                </table>"""
    
    def generate_performance_table_outlook_compatible(self, performance_data, retorno_financeiro=None):
        """Gera tabela de performance compatível com Outlook"""
        
        rows_html = ""
        for item in performance_data:
            periodo = item['periodo']
            carteira_val = item['carteira']
            benchmark_val = item['benchmark'] 
            diferenca_val = item['diferenca']
            
            # Formatação de valores
            carteira_formatted = self.format_percentage(carteira_val)
            benchmark_formatted = self.format_percentage(benchmark_val)
            diferenca_formatted = f"{diferenca_val:.2f} p.p."
            
            # Cores condicionais
            carteira_color = "#28a745" if carteira_val >= 0 else "#dc3545"
            diferenca_color = "#28a745" if diferenca_val >= 0 else "#dc3545"
            if abs(diferenca_val) < 0.01:
                diferenca_color = "#333333"
                
            rows_html += f"""
                                <tr>
                                    <td style="padding: 6px 4px; text-align: left; border: 1px solid #dee2e6; background-color: #ffffff; color: #333333; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">{periodo}</td>
                                    <td style="padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; color: {carteira_color}; font-weight: bold; background-color: #ffffff; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">{carteira_formatted}</td>
                                    <td style="padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; background-color: #ffffff; color: #333333; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">{benchmark_formatted}</td>
                                    <td style="padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; color: {diferenca_color}; font-weight: bold; background-color: #ffffff; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">{diferenca_formatted}</td>
                                </tr>"""
        
        # Linha de retorno financeiro
        if retorno_financeiro is not None:
            retorno_formatted = self.format_currency(retorno_financeiro)
            retorno_color = "#28a745" if retorno_financeiro >= 0 else "#dc3545"
            
            rows_html += f"""
                                <tr>
                                    <td style="padding: 6px 4px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; background-color: #ffffff; color: #333333; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Retorno Financeiro:</td>
                                    <td style="padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; color: {retorno_color}; font-weight: bold; background-color: #ffffff; font-size: 12px; font-family: Arial, Helvetica, sans-serif;" colspan="3">{retorno_formatted}</td>
                                </tr>"""
        
        return f"""
                            <h4 style="font-size: 14px; color: #0D2035; margin: 0 0 6px 0; padding-bottom: 4px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">Performance</h4>
                            <table cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 8px; background-color: #ffffff; border: 1px solid #dee2e6; font-family: Arial, Helvetica, sans-serif; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <thead>
                                    <tr>
                                        <th style="background-color: #f8f9fa; color: #0D2035; font-weight: bold; padding: 6px 4px; text-align: left; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Período</th>
                                        <th style="background-color: #f8f9fa; color: #0D2035; font-weight: bold; padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Carteira</th>
                                        <th style="background-color: #f8f9fa; color: #0D2035; font-weight: bold; padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Benchmark</th>
                                        <th style="background-color: #f8f9fa; color: #0D2035; font-weight: bold; padding: 6px 4px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif;">Carteira vs. Benchmark</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {rows_html}
                                </tbody>
                            </table>"""
    
    def generate_highlight_strategies_outlook_compatible(self, estrategias):
        """Gera estratégias de destaque compatível com Outlook"""
        if not estrategias:
            return ""
            
        items_html = ""
        for estrategia in estrategias:
            items_html += f"""
                                        <li style="margin-bottom: 2px; font-size: 12px; color: #333333; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">{estrategia}</li>"""
        
        return f"""
                            <h4 style="font-size: 14px; color: #0D2035; margin: 8px 0 6px 0; padding-bottom: 4px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">Estratégias de Destaque</h4>
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 0 0 8px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <tr>
                                    <td style="padding: 8px; background-color: #f0f8ff; border-left: 4px solid #0D2035;">
                                        <ul style="margin: 0; padding-left: 16px; list-style-type: disc;">
                                            {items_html}
                                        </ul>
                                    </td>
                                </tr>
                            </table>"""
    
    def generate_promoter_assets_outlook_compatible(self, ativos):
        """Gera ativos promotores compatível com Outlook"""
        if not ativos:
            return ""
            
        items_html = ""
        for ativo in ativos:
            items_html += f"""
                                        <li style="margin-bottom: 2px; font-size: 12px; color: #2e7d32; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">{ativo}</li>"""
        
        return f"""
                            <h4 style="font-size: 14px; color: #0D2035; margin: 8px 0 6px 0; padding-bottom: 4px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">Ativos Promotores</h4>
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 0 0 8px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <tr>
                                    <td style="padding: 8px; background-color: #f0fff0; border-left: 4px solid #28a745;">
                                        <ul style="margin: 0; padding-left: 16px; list-style-type: disc;">
                                            {items_html}
                                        </ul>
                                    </td>
                                </tr>
                            </table>"""
    
    def generate_detractor_assets_outlook_compatible(self, ativos):
        """Gera ativos detratores compatível com Outlook"""
        if not ativos:
            return ""
            
        items_html = ""
        for ativo in ativos:
            items_html += f"""
                                        <li style="margin-bottom: 2px; font-size: 12px; color: #c62828; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">{ativo}</li>"""
        
        return f"""
                            <h4 style="font-size: 14px; color: #0D2035; margin: 8px 0 6px 0; padding-bottom: 4px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">Ativos Detratores</h4>
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 0 0 8px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <tr>
                                    <td style="padding: 8px; background-color: #fff5f5; border-left: 4px solid #dc3545;">
                                        <ul style="margin: 0; padding-left: 16px; list-style-type: disc;">
                                            {items_html}
                                        </ul>
                                    </td>
                                </tr>
                            </table>"""
    
    def save_email_to_file(self, html_content, client_name, output_path=None):
        """Salva o conteúdo HTML do e-mail em um arquivo"""
        # Remover caracteres inválidos para nome de arquivo
        safe_client_name = "".join([c if c.isalnum() or c in [' ', '_'] else '_' for c in client_name])
        safe_client_name = safe_client_name.replace(' ', '_')
        
        date_str = datetime.now().strftime("%Y%m%d")
        
        if not output_path:
            filename = f"relatorio_mensal_{safe_client_name}_{date_str}.html"
            output_path = filename
        
        # Criar diretório para recursos
        resources_dir = "recursos_email"
        if not os.path.exists(resources_dir):
            os.makedirs(resources_dir)
        
        # Copiar o logo para o diretório de recursos
        logo_src = os.path.join("documentos", "img", "logo-MMZR-azul.png")
        logo_dest = os.path.join(resources_dir, "logo-MMZR-azul.png")
        
        if os.path.exists(logo_src):
            import shutil
            shutil.copy2(logo_src, logo_dest)
        
        # Salvar o arquivo
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path 