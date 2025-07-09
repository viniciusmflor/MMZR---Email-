"""
MMZR Family Office - Gerador de Relatórios de Performance
Versão Final - Sistema de geração de relatórios HTML para clientes

Autor: MMZR Family Office
Versão: 1.0.1 - Correção de compatibilidade com Microsoft Outlook
"""

import os
import logging
import base64
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Any, Union

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class MMZREmailGenerator:
    """
    Gerador de emails HTML para MMZR Family Office.
    Versão final otimizada com funcionalidades essenciais.
    """
    
    def __init__(self) -> None:
        """Inicializa o gerador com configurações essenciais."""
        self.meses_pt = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        self.logo_base64 = self._load_logo()
        logger.info("MMZREmailGenerator inicializado")
    
    def _load_logo(self) -> str:
        """Carrega e converte logo para base64."""
        logo_paths = [
            "documentos/img/logo-MMZR-azul.png",
            "documentos/img/LogoAzul_MMZR.jpg"
        ]
        
        for path in logo_paths:
            if os.path.exists(path):
                try:
                    with open(path, "rb") as f:
                        base64_string = base64.b64encode(f.read()).decode('utf-8')
                        mime_type = 'image/png' if path.endswith('.png') else 'image/jpeg'
                        return f"data:{mime_type};base64,{base64_string}"
                except Exception as e:
                    logger.warning(f"Erro ao carregar logo {path}: {e}")
        
        logger.warning("Logo não encontrada")
        return ""
    
    def load_excel_data(self, filepath: str) -> Optional[pd.ExcelFile]:
        """Carrega dados do arquivo Excel."""
        try:
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"Arquivo não encontrado: {filepath}")
            return pd.ExcelFile(filepath)
        except Exception as e:
            logger.error(f"Erro ao carregar {filepath}: {e}")
            return None
    
    def extract_banker_info(self, excel_path: str, client_banker: str = None) -> Dict[str, str]:
        """Extrai informações dos bankers da planilha."""
        default_bankers = {'banker_padrao': 'Felipe', 'outro_banker': 'Fernandito'}
        
        try:
            excel_file = pd.ExcelFile(excel_path)
            if "Base Consolidada" not in excel_file.sheet_names:
                return default_bankers
            
            df = pd.read_excel(excel_file, sheet_name="Base Consolidada")
            mapeamento = df[['Banker', 'NomePronomeBanker']].drop_duplicates()
            
            # Banker padrão (Felipe como fallback)
            banker4 = mapeamento[mapeamento['Banker'] == 'Banker 4']
            banker_padrao = 'Felipe' if banker4.empty or banker4.iloc[0]['NomePronomeBanker'] == 'Banker 4' else banker4.iloc[0]['NomePronomeBanker']
            
            # Outro banker (baseado no cliente ou Fernandito como padrão)
            outro_banker = 'Fernandito'
            if client_banker and client_banker != 'Banker 4':
                client_info = mapeamento[mapeamento['Banker'] == client_banker]
                if not client_info.empty and client_info.iloc[0]['NomePronomeBanker'] != client_banker:
                    outro_banker = client_info.iloc[0]['NomePronomeBanker']
            
            result = {'banker_padrao': banker_padrao, 'outro_banker': outro_banker}
            logger.info(f"Bankers: {banker_padrao} e {outro_banker}")
            return result
            
        except Exception as e:
            logger.error(f"Erro ao extrair bankers: {e}")
            return default_bankers
    
    def extract_performance_data(self, df: pd.DataFrame) -> List[Dict[str, Union[str, float]]]:
        """Extrai dados de performance (Mês atual e No ano)."""
        for i in range(len(df)):
            for j in range(len(df.columns)):
                if 'Performance' in str(df.iloc[i, j]):
                    performance_data = []
                    start_row = i + 2
                    
                    for k in range(start_row, min(start_row + 5, len(df))):
                        row = df.iloc[k]
                        if pd.notna(row.iloc[0]):
                            periodo = str(row.iloc[0]).lower()
                            
                            if "mês" in periodo or "mes" in periodo:
                                periodo = f"{self.meses_pt[datetime.now().month]}:"
                            elif "ano" in periodo:
                                periodo = "No ano:"
                            else:
                                continue
                            
                            try:
                                performance_data.append({
                                    'periodo': periodo,
                                    'carteira': float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0,
                                    'benchmark': float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0.0,
                                    'diferenca': float(row.iloc[3]) if pd.notna(row.iloc[3]) and len(row) > 3 else 0.0
                                })
                            except (ValueError, TypeError):
                                continue
                    
                    if performance_data:
                        return performance_data
        
        raise ValueError("Dados de performance não encontrados")
    
    def extract_financial_return(self, df: pd.DataFrame) -> float:
        """Extrai retorno financeiro."""
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                if 'Retorno Financeiro' in cell or ('Retorno' in cell and 'Período' not in cell):
                    for di, dj in [(1, 0), (0, 1)]:
                        ni, nj = i + di, j + dj
                        if ni < len(df) and nj < len(df.columns) and pd.notna(df.iloc[ni, nj]):
                            try:
                                return float(df.iloc[ni, nj])
                            except (ValueError, TypeError):
                                continue
        
        raise ValueError("Retorno financeiro não encontrado")
    
    def extract_highlight_strategies(self, df: pd.DataFrame) -> List[str]:
        """Extrai estratégias de destaque (máximo 2)."""
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                if 'Estratégias de Destaque' in cell or 'Destaques' in cell:
                    strategies = []
                    start_row = i + 1
                    
                    for k in range(start_row, min(start_row + 5, len(df))):
                        if len(strategies) >= 2:
                            break
                        row = df.iloc[k]
                        for l in range(min(len(row), 3)):
                            if pd.notna(row.iloc[l]) and str(row.iloc[l]).strip():
                                strategy = str(row.iloc[l])
                                if not any(s.lower() in strategy.lower() for s in ['estratégia', 'destaque', 'promotor', 'detrator']):
                                    strategies.append(strategy)
                                    break
                    
                    return strategies[:2] if strategies else ["Sem estratégias de destaque"]
        
        raise ValueError("Estratégias de destaque não encontradas")
    
    def extract_assets(self, df: pd.DataFrame, asset_type: str) -> List[str]:
        """Extrai ativos promotores ou detratores."""
        import re
        search_terms = {'promotor': 'Ativos Promotores', 'detrator': 'Ativos Detratores'}
        target_sign = 1 if asset_type == 'promotor' else -1
        
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                if search_terms[asset_type] in cell or asset_type.title() + 'es' in cell:
                    assets = []
                    start_row = i + 1
                    
                    for k in range(start_row, min(start_row + 10, len(df))):
                        if len(assets) >= 2:
                            break
                        row = df.iloc[k]
                        for l in range(min(len(row), 5)):
                            if pd.notna(row.iloc[l]) and str(row.iloc[l]).strip():
                                asset = str(row.iloc[l])
                                if not any(s.lower() in asset.lower() for s in ['ativo', 'promotor', 'detrator', 'estratégia']):
                                    percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', asset)
                                    if percentage_match:
                                        try:
                                            percentage = float(percentage_match.group(1).replace(',', '.'))
                                            if (target_sign > 0 and percentage > 0) or (target_sign < 0 and percentage < 0):
                                                if asset_type == 'promotor' and not asset.startswith('(+'):
                                                    asset = asset.replace(f"({percentage_str}%)", f"(+{percentage_str}%)")
                                                assets.append(asset)
                                                break
                                        except ValueError:
                                            continue
                    
                    return assets[:2] if assets else [f"Sem ativos {asset_type}es"]
        
        raise ValueError(f"Ativos {asset_type}es não encontrados")
    
    def format_currency(self, value: float) -> str:
        """Formata valor como moeda brasileira."""
        if value >= 0:
            return f"R$ {value:,.2f}".replace(",", ".")
        return f"-R$ {abs(value):,.2f}".replace(",", ".")
    
    def format_percentage(self, value: float) -> str:
        """Formata valor como percentual."""
        return f"+{value:.2f}%" if value > 0 else f"{value:.2f}%"
    
    def generate_html_email(self, client_name: str, data_ref: datetime, portfolios_data: List[Dict[str, Any]], bankers_info: Dict[str, str] = None) -> str:
        """Gera o HTML completo do email."""
        mes = self.meses_pt[data_ref.month]
        ano = data_ref.year
        
        if not bankers_info:
            bankers_info = {'banker_padrao': 'Felipe', 'outro_banker': 'Fernandito'}
        
        # Coletar comentários
        comentarios = []
        for p in portfolios_data:
            comentario = p.get('comentarios', '')
            if comentario and str(comentario).strip():
                comentarios.append(str(comentario).strip())
        comentario_final = ' | '.join(comentarios) if comentarios else ""
        
        html = self._generate_html_header(mes, ano)
        html += self._generate_html_body_start(client_name, data_ref)
        
        for portfolio in portfolios_data:
            html += self._generate_portfolio_section(portfolio)
        
        html += self._generate_observacoes_section(comentario_final, bankers_info)
        html += self._generate_principais_indicadores_section()
        html += self._generate_carta_section(mes, ano)
        html += self._generate_html_footer(ano)
        
        return html
    
    def _generate_html_header(self, mes: str, ano: int) -> str:
        """Gera o cabeçalho HTML."""
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="light">
    <!--[if mso]>
    <xml>
        <o:OfficeDocumentSettings>
            <o:AllowPNG/>
            <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
    </xml>
    <![endif]-->
    <style>
    /* CSS específico para Outlook */
    .mmzr-logo {{ width: 120px; height: 60px; max-width: 120px; max-height: 60px; }}
    img {{ border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }}
    table {{ border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}
    :root {{ color-scheme: light; supported-color-schemes: light; }}
    @media (prefers-color-scheme: dark) {{
        body, .body-wrapper {{ background-color: #f4f4f4 !important; }}
        .content-wrapper {{ background-color: #ffffff !important; color: #333333 !important; }}
        .header-bg {{ background-color: #0D2035 !important; }}
        .header-text {{ color: #ffffff !important; }}
        .section-bg {{ background-color: #ffffff !important; }}
        .performance-header {{ color: #0D2035 !important; border-bottom-color: #e0e0e0 !important; }}
        .data-table {{ background-color: #ffffff !important; }}
        .table-header {{ background-color: #f8f9fa !important; color: #0D2035 !important; }}
        .highlight-section {{ background-color: #f8f9fa !important; }}
        .promoters-section {{ background-color: #e8f5e9 !important; }}
        .detractors-section {{ background-color: #ffebee !important; }}
        td, th, p, h1, h2, h3, h4, h5, h6, li {{ color: inherit !important; }}
        .portfolio-header {{ background-color: #0D2035 !important; color: #ffffff !important; }}
    }}
    </style>
</head>
<body class="body-wrapper" style="margin: 0; padding: 0; font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif; line-height: 1.4; color: #333333; background-color: #f4f4f4;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse; background: #f4f4f4; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
        <tr>
            <td align="center" style="padding: 0;">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" class="content-wrapper" style="width: 100%; max-width: 800px; border-collapse: collapse; text-align: left; background: #ffffff; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                    <tr>
                        <td style="padding: 0;">
                            <table role="presentation" cellpadding="0" cellspacing="0" border="0" class="header-bg" style="width: 100%; border-collapse: collapse; background: #0D2035; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                <tr>
                                    <td style="padding: 12px 16px;">
                                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                            <tr>
                                                <td style="width: 120px; vertical-align: top; text-align: center;">
                                                    {f'<img src="{self.logo_base64}" alt="MMZR Family Office" width="120" height="60" style="width: 120px; height: 60px; max-width: 120px; display: block; border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic;">' if self.logo_base64 else '<div style="width: 120px; height: 60px; display: block; background-color: #ffffff; border: 2px solid #ffffff; border-radius: 6px; color: #0D2035; font-weight: bold; font-size: 10px; text-align: center; line-height: 1.1; padding-top: 12px;">MMZR<br>Family<br>Office</div>'}
                                                </td>
                                                <td style="vertical-align: top; padding-left: 16px;">
                                                    <h1 class="header-text" style="margin: 0 0 4px 0; font-size: 18px; font-weight: bold; color: #ffffff; line-height: 1.2; font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;">MMZR Family Office</h1>
                                                    <p class="header-text" style="margin: 0; font-size: 13px; color: #ffffff; opacity: 0.85; line-height: 1.3; font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;">Relatório Mensal de Performance<br>{mes} de {ano}</p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>"""
    
    def _generate_html_body_start(self, client_name: str, data_ref: datetime) -> str:
        """Gera o início do corpo do HTML."""
        return f"""
                    <tr>
                        <td class="section-bg" style="padding: 12px 16px; background-color: #ffffff;">
                            <h2 style="font-size: 15px; color: #0D2035; margin-bottom: 8px; margin-top: 0;">Olá {client_name},</h2>
                            <p style="margin-top: 0; margin-bottom: 6px;">Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>{data_ref.strftime('%d/%m/%Y')}</strong>.</p>"""
    
    def _generate_portfolio_section(self, portfolio: Dict[str, Any]) -> str:
        """Gera a seção de uma carteira."""
        name = portfolio.get('name', 'Carteira')
        portfolio_type = portfolio.get('type', 'Diversificada')
        data = portfolio.get('data', {})
        
        return f"""
                            <table role="presentation" style="width: 100%; margin: 12px 0 0 0; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); background-color: #ffffff;">
                                <tr>
                                    <td class="header-bg portfolio-header" style="background-color: #0D2035; color: #ffffff; padding: 6px 12px;">
                                        <h3 style="margin: 0; font-size: 16px; font-weight: 500;">{name} <span style="font-weight: 300; font-size: 13px; margin-left: 8px; opacity: 0.8;">| {portfolio_type}</span></h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="section-bg" style="padding: 10px; background-color: #ffffff;">
                                        {self._generate_performance_table(data.get('performance', []), data.get('retorno_financeiro', 0))}
                                        {self._generate_strategies_section(data.get('estrategias_destaque', []))}
                                        {self._generate_assets_section('Promotores', data.get('ativos_promotores', []), '#e8f5e9', '#2e7d32')}
                                        {self._generate_assets_section('Detratores', data.get('ativos_detratores', []), '#ffebee', '#c62828')}
                                    </td>
                                </tr>
                            </table>"""
    
    def _generate_performance_table(self, performance_data: List[Dict], retorno_financeiro: float) -> str:
        """Gera tabela de performance."""
        # Filtrar dados únicos
        filtered_data = []
        mes_added, ano_added = False, False
        
        for item in performance_data:
            periodo = item['periodo'].lower() if isinstance(item['periodo'], str) else ""
            if ":" in periodo and any(m.lower() in periodo for m in self.meses_pt.values()) and not mes_added:
                filtered_data.append(item)
                mes_added = True
            elif "no ano" in periodo and not ano_added:
                filtered_data.append(item)
                ano_added = True
            if mes_added and ano_added:
                break
        
        html = """
                                        <h4 class="performance-header" style="font-size: 16px; color: #0D2035; margin: 0 0 8px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 6px;">Performance</h4>
                                        <table role="presentation" class="data-table" style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 10px; background-color: #ffffff;">
                                            <thead>
                                                <tr>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6;">Período</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Carteira</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Benchmark</th>
                                                    <th class="table-header" style="background-color: #f8f9fa; color: #0D2035; font-weight: 600; padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6;">Carteira vs. Benchmark</th>
                                                </tr>
                                            </thead>
                                            <tbody>"""
        
        for item in filtered_data:
            carteira_color = "#28a745" if item['carteira'] > 0 else "#dc3545" if item['carteira'] < 0 else "#333333"
            diferenca_color = "#28a745" if item['diferenca'] > 0 else "#dc3545" if item['diferenca'] < 0 else "#333333"
            
            html += f"""
                                                <tr>
                                                    <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6; background-color: #ffffff;">{item['periodo']}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {carteira_color}; font-weight: 500; background-color: #ffffff;">{self.format_percentage(item['carteira'])}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; background-color: #ffffff;">{self.format_percentage(item['benchmark'])}</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {diferenca_color}; font-weight: 500; background-color: #ffffff;">{self.format_percentage(item['diferenca']).replace('%', ' p.p.')}</td>
                                                </tr>"""
        
        if retorno_financeiro is not None:
            color = "#28a745" if retorno_financeiro > 0 else "#dc3545" if retorno_financeiro < 0 else "#333333"
            html += f"""
                                                <tr>
                                                    <td style="padding: 8px 6px; text-align: left; border-bottom: 1px solid #dee2e6; font-weight: 500; background-color: #ffffff;">Retorno Financeiro:</td>
                                                    <td style="padding: 8px 6px; text-align: center; border-bottom: 1px solid #dee2e6; color: {color}; font-weight: 500; background-color: #ffffff;" colspan="3">{self.format_currency(retorno_financeiro)}</td>
                                                </tr>"""
        
        return html + """
                                            </tbody>
                                        </table>"""
    
    def _generate_strategies_section(self, strategies: List[str]) -> str:
        """Gera seção de estratégias."""
        html = """
                                        <h4 class="performance-header" style="font-size: 16px; color: #0D2035; margin: 12px 0 8px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 6px;">Estratégias de Destaque</h4>
                                        <ul class="highlight-section" style="margin: 6px 0 10px 0; padding: 8px 8px 8px 24px; background-color: #f8f9fa; border-radius: 5px; color: #333333;">"""
        
        for strategy in strategies:
            html += f"""
                                            <li style="margin-bottom: 6px; font-size: 13px;">{strategy}</li>"""
        
        return html + """
                                        </ul>"""
    
    def _generate_assets_section(self, title: str, assets: List[str], bg_color: str, text_color: str) -> str:
        """Gera seção de ativos (promotores ou detratores)."""
        import re
        
        html = f"""
                                        <h4 class="performance-header" style="font-size: 16px; color: #0D2035; margin: 12px 0 8px 0; font-weight: 500; border-bottom: 1px solid #e0e0e0; padding-bottom: 6px;">Ativos {title}</h4>
                                        <ul class="{title.lower()}-section" style="margin: 6px 0 10px 0; padding: 8px 8px 8px 24px; background-color: {bg_color}; border-radius: 5px; color: {text_color};">"""
        
        for asset in assets:
            # Adicionar "+" para promotores se necessário
            if title == 'Promotores':
                percentage_match = re.search(r'\(([-+]?\d+[.,]?\d*)%\)', asset)
                if percentage_match:
                    percentage_str = percentage_match.group(1).replace(',', '.')
                    try:
                        percentage = float(percentage_str)
                        if percentage > 0 and not percentage_str.startswith('+'):
                            asset = asset.replace(f"({percentage_str}%)", f"(+{percentage_str}%)")
                    except ValueError:
                        pass
            
            html += f"""
                                            <li style="margin-bottom: 6px; font-size: 13px;">{asset}</li>"""
        
        return html + """
                                        </ul>"""
    
    def _generate_observacoes_section(self, comentario: str, bankers_info: Dict[str, str]) -> str:
        """Gera seção de observações."""
        banker_padrao = bankers_info.get('banker_padrao', 'Felipe')
        outro_banker = bankers_info.get('outro_banker', 'Fernandito')
        
        html = f"""
                            <table role="presentation" style="width: 100%; margin-top: 12px; border-collapse: collapse; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                                <tr>
                                    <td style="padding: 10px;">
                                        <p style="margin: 0 0 12px 0; color: #555555; font-size: 13px; line-height: 18px;">
                                            <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                                        </p>
                                        <p style="margin: 0; color: #555555; font-size: 12px; font-style: italic; line-height: 16px;">
                                            <strong>Obs.:</strong> Conforme solicitado, deixo o {banker_padrao} e {outro_banker} em cópia para também receberem as informações.
                                        </p>"""
        
        if comentario:
            html += f"""
                                        <p style="margin: 12px 0 0 0; color: #555555; font-size: 13px; line-height: 18px;">
                                            <strong>Comentário:</strong> {comentario}
                                        </p>"""
        
        return html + """
                                    </td>
                                </tr>
                            </table>"""
    
    def _generate_principais_indicadores_section(self) -> str:
        """Gera seção de principais indicadores."""
        return """
                            <table role="presentation" style="width: 100%; margin-top: 8px; border-collapse: collapse; background-color: #f8f9fa; border: 1px solid #e9ecef;">
                                <tr>
                                    <td style="padding: 8px;">
                                        <p style="margin: 0 0 8px 0; font-weight: bold; color: #333333; font-size: 13px; line-height: 16px;">Principais indicadores:</p>
                                        <p style="margin: 0; color: #555555; font-size: 11px; line-height: 15px;">
                                            Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                                            Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                                        </p>
                                    </td>
                                </tr>
                            </table>"""
    
    def _generate_carta_section(self, mes: str, ano: int) -> str:
        """Gera seção da carta mensal."""
        carta_link = f"https://www.mmzrfo.com.br/post/carta-mensal-{mes.lower()}-{ano}"
        return f"""
                            <table role="presentation" style="width: 100%; margin-top: 12px; border-collapse: collapse;">
                                <tr>
                                    <td align="center" style="padding: 0;">
                                        <table role="presentation" style="border-collapse: collapse; background-color: #0D2035; border-radius: 4px;">
                                            <tr>
                                                <td style="padding: 8px 16px; text-align: center;">
                                                    <a href="{carta_link}" target="_blank" style="color: #ffffff; text-decoration: none; font-weight: bold; font-size: 14px; line-height: 18px;">Confira nossa carta completa: Carta {mes} {ano}</a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>"""
    
    def _generate_html_footer(self, ano: int) -> str:
        """Gera rodapé HTML."""
        return f"""
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #f8f9fa; padding: 8px 16px; text-align: center;">
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
    
    def generate_email_subject(self, data_ref: datetime) -> str:
        """Gera assunto do email."""
        mes_nome = self.meses_pt[data_ref.month]
        return f"MMZR Family Office - Relatório de Performance - {mes_nome}/{data_ref.year}"
    
    def save_email_to_file(self, html_content: str, client_name: str, output_path: Optional[str] = None) -> str:
        """Salva email em arquivo HTML."""
        try:
            safe_name = "".join([c if c.isalnum() or c in [' ', '_'] else '_' for c in client_name]).replace(' ', '_')
            date_str = datetime.now().strftime("%Y%m%d")
            
            if not output_path:
                output_path = f"relatorio_mensal_{safe_name}_{date_str}.html"
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logger.info(f"Relatório salvo: {output_path}")
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
        client_banker = client_config.get('banker', None)  # Banker do cliente se disponível
        
        # Data de referência (hoje como padrão)
        data_ref = datetime.now()
        
        # Extrair informações dos bankers
        bankers_info = generator.extract_banker_info(excel_path, client_banker)
        
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
                        'ativos_promotores': generator.extract_assets(df, 'promotor'),
                        'ativos_detratores': generator.extract_assets(df, 'detrator')
                    }
                }
                
                portfolios_data.append(portfolio_data)
            else:
                error_msg = f"Aba '{sheet_name}' não encontrada no Excel. O relatório não pode ser gerado."
                logger.error(error_msg)
                raise ValueError(error_msg)
        
        # Gerar o HTML do e-mail com informações dinâmicas dos bankers
        html_content = generator.generate_html_email(client_name, data_ref, portfolios_data, bankers_info)
        
        # Salvar o e-mail em um arquivo
        output_file = generator.save_email_to_file(html_content, client_name)
        
        logger.info(f"Relatório gerado com sucesso para {client_name}!")
        logger.info(f"Bankers: {bankers_info['banker_padrao']} e {bankers_info['outro_banker']}")
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