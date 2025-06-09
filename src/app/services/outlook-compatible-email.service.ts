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

    return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="only light">
    <meta name="supported-color-schemes" content="only light">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <!--[if mso]>
    <style type="text/css">
        body, table, td, p, a, li, blockquote {font-family: Arial, Helvetica, sans-serif !important;}
        table {border-collapse: collapse !important;}
        img {border: 0 !important;}
    </style>
    <![endif]-->
</head>
<body style="margin: 0 !important; padding: 0 !important; background-color: ${this.corFundo} !important; color: ${this.corTexto} !important; font-family: Arial, Helvetica, sans-serif !important; -webkit-text-size-adjust: 100% !important; -ms-text-size-adjust: 100% !important;">
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse !important; background-color: ${this.corFundo} !important; mso-table-lspace: 0pt !important; mso-table-rspace: 0pt !important;">
        ${this.gerarCabecalho(config.logoBase64, mesFormatado, anoFormatado)}
        
        <tr>
            <td style="padding: 20px !important; background-color: ${this.corFundo} !important;">
                ${this.gerarSaudacao(config.clientName, dataFormatada)}
                
                ${config.portfolios.map(portfolio => this.gerarSecaoPortfolio(portfolio)).join('')}
                
                ${this.gerarObservacoes()}
                
                ${this.gerarIndicadores()}
                
                ${this.gerarBotaoCarta(mesFormatado, anoFormatado)}
            </td>
        </tr>
        
        ${this.gerarRodape(anoFormatado, config.customFooter)}
    </table>
</body>
</html>`;
  }

  private gerarCabecalho(logoBase64?: string, mes?: string, ano?: number): string {
    const logoHtml = logoBase64 
      ? `<img src="${logoBase64}" alt="MMZR Family Office" width="80" height="64" style="display: block !important; border: 0 !important; max-width: 80px !important; height: auto !important;">`
      : `<div style="width: 50px; height: 40px; background-color: ${this.corFundo}; color: ${this.corPrimaria}; font-weight: bold; font-size: 12px; text-align: center; line-height: 40px;">MMZR</div>`;

    return `
        <tr>
            <td style="background-color: ${this.corPrimaria} !important; padding: 15px !important; text-align: center !important;">
                <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse !important;">
                    <tr>
                        <td style="text-align: left !important; vertical-align: middle !important; width: 90px !important;">
                            ${logoHtml}
                        </td>
                        <td style="text-align: left !important; vertical-align: middle !important; padding-left: 15px !important;">
                            <h1 style="margin: 0 !important; font-size: 20px !important; color: ${this.corFundo} !important; font-weight: bold !important; font-family: Arial, Helvetica, sans-serif !important;">MMZR Family Office</h1>
                            <p style="margin: 5px 0 0 0 !important; font-size: 16px !important; color: ${this.corFundo} !important; font-family: Arial, Helvetica, sans-serif !important;">Relatório Mensal - ${mes} ${ano}</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>`;
  }

  private gerarSaudacao(clientName: string, dataFormatada: string): string {
    return `
                <p style="margin: 0 0 15px 0 !important; font-size: 14px !important; color: ${this.corTexto} !important; font-family: Arial, Helvetica, sans-serif !important;">
                    Olá ${clientName},
                </p>
                
                <p style="margin: 0 0 20px 0 !important; font-size: 14px !important; color: ${this.corTexto} !important; line-height: 1.5 !important; font-family: Arial, Helvetica, sans-serif !important;">
                    Segue o relatório mensal com o desempenho de suas carteiras referente a <strong>${dataFormatada}</strong>.
                </p>`;
  }

  private gerarSecaoPortfolio(portfolio: PortfolioData): string {
    return `
                <!-- Carteira: ${portfolio.name} -->
                <div style="margin: 25px 0 !important; border: 1px solid #e0e0e0 !important; border-radius: 5px !important; overflow: hidden !important; background-color: ${this.corFundo} !important; box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;">
                    <div style="background-color: ${this.corPrimaria} !important; color: ${this.corFundo} !important; padding: 12px !important;">
                        <h3 style="margin: 0 !important; font-size: 16px !important; color: ${this.corFundo} !important; font-weight: bold !important; font-family: Arial, Helvetica, sans-serif !important;">${portfolio.name}</h3>
                        <span style="font-size: 14px !important; color: ${this.corFundo} !important; font-family: Arial, Helvetica, sans-serif !important; opacity: 0.9 !important;">${portfolio.type}</span>
                    </div>
                    <div style="padding: 15px !important; background-color: ${this.corFundo} !important;">
                        
                        ${this.gerarTabelaPerformance(portfolio.data.performance, portfolio.data.retorno_financeiro)}
                        
                        ${this.gerarListaItens('Estratégias de Destaque', portfolio.data.estrategias_destaque, '#f0f8ff', this.corPrimaria)}
                        
                        ${this.gerarListaItens('Ativos Promotores', portfolio.data.ativos_promotores, '#f0fff0', this.corSuccesso)}
                        
                        ${this.gerarListaItens('Ativos Detratores', portfolio.data.ativos_detratores, '#fff5f5', this.corPerigo)}
                    </div>
                </div>`;
  }

  private gerarTabelaPerformance(performance: PerformanceItem[], retornoFinanceiro?: number): string {
    const linhasPerformance = performance.map(item => {
      const corCarteira = item.carteira > 0 ? this.corSuccesso : item.carteira < 0 ? this.corPerigo : this.corTexto;
      const corDiferenca = item.diferenca > 0 ? this.corSuccesso : item.diferenca < 0 ? this.corPerigo : this.corTexto;

      return `
                                <tr>
                                    <td style="padding: 8px 6px !important; text-align: left !important; border: 1px solid #dee2e6 !important; background-color: ${this.corFundo} !important; color: ${this.corTexto} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">${item.periodo}</td>
                                    <td style="padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; color: ${corCarteira} !important; font-weight: bold !important; background-color: ${this.corFundo} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">${this.formatarPorcentagem(item.carteira)}</td>
                                    <td style="padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; background-color: ${this.corFundo} !important; color: ${this.corTexto} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">${this.formatarPorcentagem(item.benchmark)}</td>
                                    <td style="padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; color: ${corDiferenca} !important; font-weight: bold !important; background-color: ${this.corFundo} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">${this.formatarPorcentagem(item.diferenca).replace('%', ' p.p.')}</td>
                                </tr>`;
    }).join('');

    const linhaRetorno = retornoFinanceiro !== undefined ? `
                                <tr>
                                    <td style="padding: 8px 6px !important; text-align: left !important; border: 1px solid #dee2e6 !important; font-weight: bold !important; background-color: ${this.corFundo} !important; color: ${this.corTexto} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Retorno Financeiro:</td>
                                    <td style="padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; color: ${retornoFinanceiro >= 0 ? this.corSuccesso : this.corPerigo} !important; font-weight: bold !important; background-color: ${this.corFundo} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;" colspan="3">${this.formatarMoeda(retornoFinanceiro)}</td>
                                </tr>` : '';

    return `
                        <h4 style="font-size: 14px !important; color: ${this.corPrimaria} !important; margin: 0 0 10px 0 !important; padding-bottom: 5px !important; border-bottom: 1px solid #e0e0e0 !important; font-weight: bold !important; font-family: Arial, Helvetica, sans-serif !important;">Performance</h4>
                        <table cellpadding="0" cellspacing="0" border="0" style="width: 100% !important; border-collapse: collapse !important; font-size: 12px !important; margin-bottom: 15px !important; background-color: ${this.corFundo} !important; border: 1px solid #dee2e6 !important; font-family: Arial, Helvetica, sans-serif !important;">
                            <thead>
                                <tr>
                                    <th style="background-color: #f8f9fa !important; color: ${this.corPrimaria} !important; font-weight: bold !important; padding: 8px 6px !important; text-align: left !important; border: 1px solid #dee2e6 !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Período</th>
                                    <th style="background-color: #f8f9fa !important; color: ${this.corPrimaria} !important; font-weight: bold !important; padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Carteira</th>
                                    <th style="background-color: #f8f9fa !important; color: ${this.corPrimaria} !important; font-weight: bold !important; padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Benchmark</th>
                                    <th style="background-color: #f8f9fa !important; color: ${this.corPrimaria} !important; font-weight: bold !important; padding: 8px 6px !important; text-align: center !important; border: 1px solid #dee2e6 !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Carteira vs. Benchmark</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${linhasPerformance}
                                ${linhaRetorno}
                            </tbody>
                        </table>
                        `;
  }

  private gerarListaItens(titulo: string, itens: string[], corFundo: string, corBorda: string): string {
    if (!itens || itens.length === 0) {
      return '';
    }

    const corTextoItem = titulo.includes('Detratores') ? '#c62828' : titulo.includes('Promotores') ? '#2e7d32' : this.corTexto;
    
    const itensHtml = itens.map(item => 
      `<li style="margin-bottom: 1px; font-size: 12px; color: ${corTextoItem}; line-height: 1.3; font-family: Arial, Helvetica, sans-serif;">${item}</li>`
    ).join('');

    return `
                        
                        <h4 style="font-size: 14px; color: ${this.corPrimaria}; margin: 15px 0 8px 0; padding-bottom: 5px; border-bottom: 1px solid #e0e0e0; font-weight: bold; font-family: Arial, Helvetica, sans-serif;">${titulo}</h4>
                        <div style="margin: 0 0 12px 0; padding: 6px; background-color: ${corFundo}; border-left: 4px solid ${corBorda}; border-radius: 3px;">
                            <ul style="margin: 0; padding-left: 12px; list-style-type: disc;">
                                ${itensHtml}
                            </ul>
                        </div>`;
  }

  private gerarObservacoes(): string {
    return `
                
                <!-- Observações finais -->
                <div style="margin-top: 20px !important; padding: 15px !important; background-color: #f8f9fa !important; border-radius: 5px !important; border: 1px solid #e9ecef !important;">
                    <p style="margin: 0 0 10px 0 !important; color: #555555 !important; font-size: 12px !important; line-height: 1.4 !important; font-family: Arial, Helvetica, sans-serif !important;">
                        <strong>Obs.:</strong> Eventuais ajustes retroativos do IPCA, após a divulgação oficial do indicador, podem impactar marginalmente a rentabilidade do portfólio no mês anterior.
                    </p>
                    <p style="margin: 0 !important; color: #555555 !important; font-size: 11px !important; font-style: italic !important; line-height: 1.3 !important; font-family: Arial, Helvetica, sans-serif !important;">
                        <strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e Fernandito em cópia para também receberem as informações.
                    </p>
                </div>`;
  }

  private gerarIndicadores(): string {
    return `

                <!-- Principais indicadores -->
                <div style="margin-top: 15px !important; padding: 10px !important; background-color: #f8f9fa !important; border-radius: 5px !important; border: 1px solid #e9ecef !important;">
                    <p style="margin: 0 0 5px 0 !important; font-weight: bold !important; color: ${this.corTexto} !important; font-size: 12px !important; font-family: Arial, Helvetica, sans-serif !important;">Principais indicadores:</p>
                    <p style="margin: 0 !important; color: #555555 !important; font-size: 10px !important; font-style: italic !important; line-height: 1.3 !important; font-family: Arial, Helvetica, sans-serif !important;">
                        Locais: CDI: +1,06%, Ibovespa: +3,69%, Prefixados (IRF-M): +2,99%, Ativos IPCA (IMA-B): +2,09%, Imobiliários (IFIX): +3,01%, Dólar (Ptax): -1,42%, Multimercados (IHFA): +3,85%<br>
                        Internacionais: MSCI AC: +0,77%, S&P 500 -0,76%, Euro Stoxx 600 -1,21%, MSCI China -4,55%, MSCI EM +1,04%, Ouro +5,29%, Petróleo BRENT -14,97%, Minério de ferro -2,68% e Bitcoin (IBIT) +14,31%
                    </p>
                </div>`;
  }

  private gerarBotaoCarta(mes: string, ano: number): string {
    const mesLowercase = mes.toLowerCase();
    const cartaLink = `https://www.mmzrfo.com.br/post/carta-mensal-${mesLowercase}-${ano}`;

    return `
                
                <!-- Link para carta mensal como botão azul -->
                <div style="margin-top: 20px !important; text-align: center !important;">
                    <a href="${cartaLink}" target="_blank" style="display: inline-block !important; background-color: ${this.corPrimaria} !important; color: ${this.corFundo} !important; padding: 12px 24px !important; text-decoration: none !important; border-radius: 4px !important; font-weight: bold !important; font-size: 14px !important; font-family: Arial, Helvetica, sans-serif !important; text-align: center !important; border: none !important; -webkit-text-size-adjust: none !important;">Confira nossa carta completa: Carta ${mes} ${ano}</a>
                </div>`;
  }

  private gerarRodape(ano: number, customFooter?: string): string {
    const footerText = customFooter || 'MMZR Family Office | Gestão de Patrimônio';

    return `
        
        <!-- Footer -->
        <tr>
            <td style="background-color: #f8f9fa !important; padding: 15px !important; text-align: center !important; border-top: 1px solid #e9ecef !important;">
                <p style="margin: 0 0 5px 0 !important; color: #666666 !important; font-size: 11px !important; font-family: Arial, Helvetica, sans-serif !important;">${footerText}</p>
                <p style="margin: 0 !important; color: #666666 !important; font-size: 11px !important; font-family: Arial, Helvetica, sans-serif !important;">© ${ano} MMZR Family Office. Todos os direitos reservados.</p>
            </td>
        </tr>`;
  }

  /**
   * Formata uma porcentagem com sinal
   */
  private formatarPorcentagem(valor: number): string {
    const sinal = valor > 0 ? '+' : '';
    return `\${sinal}\${valor.toFixed(2)}%`;
  }

  /**
   * Formatar moeda brasileira
   */
  private formatarMoeda(valor: number): string {
    const sinal = valor >= 0 ? 'R$ ' : '-R$ ';
    const valorAbsoluto = Math.abs(valor);
    return `\${sinal}\${valorAbsoluto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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
    return `MMZR Family Office | Desempenho \${mes} de \${ano}`;
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