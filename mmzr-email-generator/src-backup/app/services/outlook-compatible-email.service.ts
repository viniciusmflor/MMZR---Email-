import { Injectable } from '@angular/core';

export interface PortfolioData {
  name: string;
  type: string;
  data: {
    performance: PerformanceItem[];
    retorno_financeiro?: number;
    estrategias_destaque: string[];
    ativos_promotores: string[];
    ativos_detratores: string[];
  };
}

export interface PerformanceItem {
  periodo: string;
  carteira: number;
  benchmark: number;
  diferenca: number;
}

export interface EmailConfiguration {
  clientName: string;
  dataRef: Date;
  portfolios: PortfolioData[];
  companyName?: string;
  logoBase64?: string;
  customFooter?: string;
}

@Injectable({
  providedIn: 'root'
})
export class OutlookCompatibleEmailService {
  private readonly mesesPortugues = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
  };

  private readonly corPrimaria = '#0D2035';
  private readonly corSuccesso = '#28a745';
  private readonly corPerigo = '#dc3545';
  private readonly corTexto = '#333333';
  private readonly corFundo = '#ffffff';

  /**
   * Gera o HTML do email otimizado para Outlook e outros clientes de email
   */
  generateOutlookCompatibleEmail(config: EmailConfiguration): string {
    const mesFormatado = this.mesesPortugues[config.dataRef.getMonth() + 1];
    const anoFormatado = config.dataRef.getFullYear();
    const dataFormatada = this.formatarData(config.dataRef);

    return `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <meta name="color-scheme" content="only light">
    <meta name="supported-color-schemes" content="only light">
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
    <!--[if mso]>
    <style type="text/css">
        body, table, td, p, a, li, blockquote {
            font-family: Arial, sans-serif !important;
        }
        table {
            border-collapse: collapse !important;
        }
        .outlook-group-fix {
            width: 100% !important;
        }
    </style>
    <![endif]-->
    <title>MMZR Family Office - Relatório ${mesFormatado} ${anoFormatado}</title>
</head>
<body style="margin: 0; padding: 0; background-color: ${this.corFundo}; font-family: Arial, Helvetica, sans-serif; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; font-size: 14px; line-height: 1.4; color: ${this.corTexto};">
    <!--[if mso | IE]>
    <table align="center" border="0" cellpadding="0" cellspacing="0" class="outlook-group-fix" role="presentation" style="width:100%;">
    <tr>
    <td>
    <![endif]-->
    
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; max-width: 600px; margin: 0 auto; background-color: ${this.corFundo};">
        ${this.gerarCabecalho(config.logoBase64, mesFormatado, anoFormatado)}
        
        <tr>
            <td style="padding: 20px; background-color: ${this.corFundo};">
                ${this.gerarSaudacao(config.clientName, dataFormatada)}
                
                ${config.portfolios.map(portfolio => this.gerarSecaoPortfolio(portfolio)).join('')}
                
                ${this.gerarObservacoes()}
                
                ${this.gerarIndicadores()}
                
                ${this.gerarBotaoCarta(mesFormatado, anoFormatado)}
            </td>
        </tr>
        
        ${this.gerarRodape(anoFormatado, config.customFooter)}
    </table>
    
    <!--[if mso | IE]>
    </td>
    </tr>
    </table>
    <![endif]-->
</body>
</html>`;
  }

  private gerarCabecalho(logoBase64?: string, mes?: string, ano?: number): string {
    const logoHtml = logoBase64 
      ? `<img src="${logoBase64}" alt="MMZR Family Office" width="50" height="40" style="display: block; border: 0; max-width: 50px; height: auto;"/>`
      : `<div style="width: 50px; height: 40px; background-color: ${this.corFundo}; color: ${this.corPrimaria}; font-weight: bold; font-size: 12px; text-align: center; line-height: 40px;">MMZR</div>`;

    return `
        <tr>
            <td style="background-color: ${this.corPrimaria}; padding: 15px;">
                <!--[if mso | IE]>
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                <td width="60" style="vertical-align: middle;">
                <![endif]-->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                    <tr>
                        <td width="60" style="vertical-align: middle; text-align: left;">
                            ${logoHtml}
                        </td>
                        <td style="vertical-align: middle; padding-left: 15px;">
                            <h1 style="margin: 0; font-size: 20px; color: ${this.corFundo}; font-weight: bold; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">MMZR Family Office</h1>
                            <p style="margin: 5px 0 0 0; font-size: 16px; color: ${this.corFundo}; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Relatório Mensal - ${mes} ${ano}</p>
                        </td>
                    </tr>
                </table>
                <!--[if mso | IE]>
                </td>
                </tr>
                </table>
                <![endif]-->
            </td>
        </tr>`;
  }

  private gerarSaudacao(clientName: string, dataFormatada: string): string {
    return `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                    <tr>
                        <td style="padding: 0 0 15px 0;">
                            <p style="margin: 0; font-size: 14px; color: ${this.corTexto}; font-family: Arial, Helvetica, sans-serif; line-height: 1.4;">
                                Olá ${clientName},
                            </p>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 0 0 20px 0;">
                            <p style="margin: 0; font-size: 14px; color: ${this.corTexto}; line-height: 1.5; font-family: Arial, Helvetica, sans-serif;">
                                Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>${dataFormatada}</strong>.
                            </p>
                        </td>
                    </tr>
                </table>`;
  }

  private gerarSecaoPortfolio(portfolio: PortfolioData): string {
    return `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin: 25px 0; border: 1px solid #e0e0e0; border-radius: 5px; overflow: hidden; background-color: ${this.corFundo};">
                    <tr>
                        <td style="background-color: ${this.corPrimaria}; padding: 12px;">
                            <h3 style="margin: 0; font-size: 16px; color: ${this.corFundo}; font-weight: bold; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${portfolio.name}</h3>
                            <span style="font-size: 14px; color: ${this.corFundo}; font-family: Arial, Helvetica, sans-serif; opacity: 0.9; line-height: 1.2;">${portfolio.type}</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 15px; background-color: ${this.corFundo};">
                            ${this.gerarTabelaPerformance(portfolio.data.performance, portfolio.data.retorno_financeiro)}
                            ${this.gerarListaItens('Estratégias de Destaque', portfolio.data.estrategias_destaque, '#f0f8ff', this.corPrimaria)}
                            ${this.gerarListaItens('Ativos Promotores', portfolio.data.ativos_promotores, '#f0fff0', this.corSuccesso)}
                            ${this.gerarListaItens('Ativos Detratores', portfolio.data.ativos_detratores, '#fff5f5', this.corPerigo)}
                        </td>
                    </tr>
                </table>`;
  }

  private gerarTabelaPerformance(performance: PerformanceItem[], retornoFinanceiro?: number): string {
    const linhasPerformance = performance.map(item => {
      const corCarteira = item.carteira > 0 ? this.corSuccesso : item.carteira < 0 ? this.corPerigo : this.corTexto;
      const corDiferenca = item.diferenca > 0 ? this.corSuccesso : item.diferenca < 0 ? this.corPerigo : this.corTexto;

      return `
                                <tr>
                                    <td style="padding: 8px 6px; text-align: left; border: 1px solid #dee2e6; background-color: ${this.corFundo}; color: ${this.corTexto}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${item.periodo}</td>
                                    <td style="padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; color: ${corCarteira}; font-weight: bold; background-color: ${this.corFundo}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${this.formatarPorcentagem(item.carteira)}</td>
                                    <td style="padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; background-color: ${this.corFundo}; color: ${this.corTexto}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${this.formatarPorcentagem(item.benchmark)}</td>
                                    <td style="padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; color: ${corDiferenca}; font-weight: bold; background-color: ${this.corFundo}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${this.formatarPorcentagem(item.diferenca).replace('%', ' p.p.')}</td>
                                </tr>`;
    }).join('');

    const linhaRetorno = retornoFinanceiro !== undefined ? `
                                <tr>
                                    <td style="padding: 8px 6px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; background-color: ${this.corFundo}; color: ${this.corTexto}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Retorno Financeiro:</td>
                                    <td colspan="3" style="padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; color: ${retornoFinanceiro >= 0 ? this.corSuccesso : this.corPerigo}; font-weight: bold; background-color: ${this.corFundo}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${this.formatarMoeda(retornoFinanceiro)}</td>
                                </tr>` : '';

    return `
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-bottom: 15px;">
                                <tr>
                                    <td>
                                        <h4 style="font-size: 14px; color: ${this.corPrimaria}; margin: 0 0 10px 0; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Performance</h4>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; border: 1px solid #dee2e6;">
                                            <thead>
                                                <tr>
                                                    <th style="background-color: #f8f9fa; color: ${this.corPrimaria}; font-weight: bold; padding: 8px 6px; text-align: left; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Período</th>
                                                    <th style="background-color: #f8f9fa; color: ${this.corPrimaria}; font-weight: bold; padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Carteira</th>
                                                    <th style="background-color: #f8f9fa; color: ${this.corPrimaria}; font-weight: bold; padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Benchmark</th>
                                                    <th style="background-color: #f8f9fa; color: ${this.corPrimaria}; font-weight: bold; padding: 8px 6px; text-align: center; border: 1px solid #dee2e6; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Carteira vs. Benchmark</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                ${linhasPerformance}
                                                ${linhaRetorno}
                                            </tbody>
                                        </table>
                                    </td>
                                </tr>
                            </table>`;
  }

  private gerarListaItens(titulo: string, itens: string[], corFundo: string, corBorda: string): string {
    if (!itens || itens.length === 0) {
      return '';
    }

    const itensHtml = itens.map(item => 
      `<li style="margin-bottom: 3px; font-size: 12px; color: ${titulo.includes('Detratores') ? '#c62828' : titulo.includes('Promotores') ? '#2e7d32' : this.corTexto}; line-height: 1.3; font-family: Arial, Helvetica, sans-serif;">${item}</li>`
    ).join('');

    return `
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin: 15px 0 8px 0;">
                                <tr>
                                    <td>
                                        <h4 style="font-size: 14px; color: ${this.corPrimaria}; margin: 0 0 8px 0; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${titulo}</h4>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 6px; background-color: ${corFundo}; border-left: 4px solid ${corBorda}; border-radius: 3px;">
                                        <ul style="margin: 0; padding-left: 20px; list-style-type: disc;">
                                            ${itensHtml}
                                        </ul>
                                    </td>
                                </tr>
                            </table>`;
  }

  private gerarObservacoes(): string {
    return `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 20px;">
                    <tr>
                        <td style="padding: 15px; background-color: #f8f9fa; border-radius: 5px; border: 1px solid #e9ecef;">
                            <p style="margin: 0 0 10px 0; color: #555555; font-size: 12px; line-height: 1.4; font-family: Arial, Helvetica, sans-serif;">
                                <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                            </p>
                            <p style="margin: 0; color: #555555; font-size: 11px; font-style: italic; line-height: 1.3; font-family: Arial, Helvetica, sans-serif;">
                                <strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e Fernandito em cópia para também receberem as informações.
                            </p>
                        </td>
                    </tr>
                </table>`;
  }

  private gerarIndicadores(): string {
    return `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 15px;">
                    <tr>
                        <td style="padding: 10px; background-color: #f8f9fa; border-radius: 5px; border: 1px solid #e9ecef;">
                            <p style="margin: 0 0 5px 0; font-weight: bold; color: ${this.corTexto}; font-size: 12px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">Principais indicadores:</p>
                            <p style="margin: 0; color: #555555; font-size: 10px; font-style: italic; line-height: 1.3; font-family: Arial, Helvetica, sans-serif;">
                                Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                                Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                            </p>
                        </td>
                    </tr>
                </table>`;
  }

  private gerarBotaoCarta(mes: string, ano: number): string {
    const mesLowercase = mes.toLowerCase();
    const cartaLink = `https://www.mmzrfo.com.br/post/carta-mensal-${mesLowercase}-${ano}`;

    return `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin-top: 20px;">
                    <tr>
                        <td style="text-align: center;">
                            <!--[if mso]>
                            <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="${cartaLink}" style="height:44px;v-text-anchor:middle;width:280px;" arcsize="9%" stroke="f" fillcolor="${this.corPrimaria}">
                            <w:anchorlock/>
                            <center style="color:${this.corFundo};font-family:Arial,sans-serif;font-size:14px;font-weight:bold;">Confira nossa carta completa: Carta ${mes} ${ano}</center>
                            </v:roundrect>
                            <![endif]-->
                            <!--[if !mso]><!-->
                            <a href="${cartaLink}" target="_blank" style="display: inline-block; background-color: ${this.corPrimaria}; color: ${this.corFundo}; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px; font-family: Arial, Helvetica, sans-serif; text-align: center; border: none; line-height: 1.2;">Confira nossa carta completa: Carta ${mes} ${ano}</a>
                            <!--<![endif]-->
                        </td>
                    </tr>
                </table>`;
  }

  private gerarRodape(ano: number, customFooter?: string): string {
    const footerText = customFooter || 'MMZR Family Office | Gestão de Patrimônio';

    return `
        <tr>
            <td style="background-color: #f8f9fa; padding: 15px; text-align: center; border-top: 1px solid #e9ecef;">
                <p style="margin: 0 0 5px 0; color: #666666; font-size: 11px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">${footerText}</p>
                <p style="margin: 0; color: #666666; font-size: 11px; font-family: Arial, Helvetica, sans-serif; line-height: 1.2;">© ${ano} MMZR Family Office. Todos os direitos reservados.</p>
            </td>
        </tr>`;
  }

  /**
   * Formata uma porcentagem com sinal
   */
  private formatarPorcentagem(valor: number): string {
    const sinal = valor > 0 ? '+' : '';
    return `${sinal}${valor.toFixed(2)}%`;
  }

  /**
   * Formatar moeda brasileira
   */
  private formatarMoeda(valor: number): string {
    const sinal = valor >= 0 ? 'R$ ' : '-R$ ';
    const valorAbsoluto = Math.abs(valor);
    return `${sinal}${valorAbsoluto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  }

  /**
   * Formatar data brasileira
   */
  private formatarData(data: Date): string {
    return data.toLocaleDateString('pt-BR');
  }

  /**
   * Gera um assunto padrão para o email
   */
  generateEmailSubject(dataRef: Date): string {
    const mes = this.mesesPortugues[dataRef.getMonth() + 1];
    const ano = dataRef.getFullYear();
    return `MMZR Family Office | Desempenho ${mes} de ${ano}`;
  }

  /**
   * Valida se os dados do portfólio estão completos
   */
  validatePortfolioData(portfolio: PortfolioData): boolean {
    return !!(
      portfolio.name &&
      portfolio.type &&
      portfolio.data &&
      Array.isArray(portfolio.data.performance) &&
      Array.isArray(portfolio.data.estrategias_destaque) &&
      Array.isArray(portfolio.data.ativos_promotores) &&
      Array.isArray(portfolio.data.ativos_detratores)
    );
  }

  /**
   * Converte imagem para base64
   */
  async convertImageToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = error => reject(error);
      reader.readAsDataURL(file);
    });
  }
} 