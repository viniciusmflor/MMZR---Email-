import { Component, signal, computed, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { 
  OutlookCompatibleEmailService,
  EmailConfiguration,
  PortfolioData,
  PerformanceItem
} from '../../services/outlook-compatible-email.service';

@Component({
  selector: 'app-email-generator',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="email-generator-container">
      <div class="form-section">
        <h2>Gerador de Email Compatível com Outlook</h2>
        
        <div class="form-group">
          <label for="clientName">Nome do Cliente:</label>
          <input 
            id="clientName"
            type="text" 
            [(ngModel)]="clientName"
            placeholder="Digite o nome do cliente"
            class="form-control">
        </div>

        <div class="form-group">
          <label for="reportDate">Data do Relatório:</label>
          <input 
            id="reportDate"
            type="date" 
            [(ngModel)]="reportDate"
            class="form-control">
        </div>

        <div class="form-group">
          <label for="logoUpload">Logo da Empresa (Base64):</label>
          <input 
            id="logoUpload"
            type="file" 
            accept="image/*"
            (change)="onLogoUpload($event)"
            class="form-control">
          @if (logoBase64()) {
            <div class="logo-preview">
              <img [src]="logoBase64()" alt="Logo Preview" style="max-width: 100px; height: auto;">
            </div>
          }
        </div>

        <div class="portfolios-section">
          <h3>Portfólios</h3>
          @for (portfolio of portfolios(); track portfolio.name) {
            <div class="portfolio-form">
              <h4>{{ portfolio.name || 'Novo Portfólio' }}</h4>
              
              <div class="form-row">
                <div class="form-group">
                  <label>Nome do Portfólio:</label>
                  <input 
                    type="text" 
                    [(ngModel)]="portfolio.name"
                    placeholder="Ex: Offshore"
                    class="form-control">
                </div>
                
                <div class="form-group">
                  <label>Tipo:</label>
                  <input 
                    type="text" 
                    [(ngModel)]="portfolio.type"
                    placeholder="Ex: Diversificada"
                    class="form-control">
                </div>
              </div>

              <div class="form-group">
                <label>Retorno Financeiro:</label>
                <input 
                  type="number" 
                  step="0.01"
                  [(ngModel)]="portfolio.data.retorno_financeiro"
                  placeholder="Ex: -17026.39"
                  class="form-control">
              </div>

              <div class="performance-section">
                <h5>Performance</h5>
                @for (perf of portfolio.data.performance; track $index) {
                  <div class="performance-row">
                    <input 
                      type="text" 
                      [(ngModel)]="perf.periodo"
                      placeholder="Período (ex: Junho:)"
                      class="form-control">
                    <input 
                      type="number" 
                      step="0.01"
                      [(ngModel)]="perf.carteira"
                      placeholder="Carteira %"
                      class="form-control">
                    <input 
                      type="number" 
                      step="0.01"
                      [(ngModel)]="perf.benchmark"
                      placeholder="Benchmark %"
                      class="form-control">
                    <input 
                      type="number" 
                      step="0.01"
                      [(ngModel)]="perf.diferenca"
                      placeholder="Diferença %"
                      class="form-control">
                    <button type="button" (click)="removePerformanceItem(portfolio, $index)" class="btn-remove">
                      Remover
                    </button>
                  </div>
                }
                <button type="button" (click)="addPerformanceItem(portfolio)" class="btn-add">
                  Adicionar Performance
                </button>
              </div>

              <div class="strategies-section">
                <h5>Estratégias de Destaque</h5>
                @for (strategy of portfolio.data.estrategias_destaque; track $index) {
                  <div class="strategy-row">
                    <input 
                      type="text" 
                      [(ngModel)]="portfolio.data.estrategias_destaque[$index]"
                      placeholder="Ex: FIXED INCOME (-0.05%)"
                      class="form-control">
                    <button type="button" (click)="removeStrategy(portfolio, $index)" class="btn-remove">
                      Remover
                    </button>
                  </div>
                }
                <button type="button" (click)="addStrategy(portfolio)" class="btn-add">
                  Adicionar Estratégia
                </button>
              </div>

              <div class="assets-section">
                <h5>Ativos Promotores</h5>
                @for (asset of portfolio.data.ativos_promotores; track $index) {
                  <div class="asset-row">
                    <input 
                      type="text" 
                      [(ngModel)]="portfolio.data.ativos_promotores[$index]"
                      placeholder="Ex: JUPITER GLOBAL EQUITY (+1.16%)"
                      class="form-control">
                    <button type="button" (click)="removePromoterAsset(portfolio, $index)" class="btn-remove">
                      Remover
                    </button>
                  </div>
                }
                <button type="button" (click)="addPromoterAsset(portfolio)" class="btn-add">
                  Adicionar Ativo Promotor
                </button>
              </div>

              <div class="assets-section">
                <h5>Ativos Detratores</h5>
                @for (asset of portfolio.data.ativos_detratores; track $index) {
                  <div class="asset-row">
                    <input 
                      type="text" 
                      [(ngModel)]="portfolio.data.ativos_detratores[$index]"
                      placeholder="Ex: BLACKROCK WORLD TECHNOLOGY (-11.23%)"
                      class="form-control">
                    <button type="button" (click)="removeDetractorAsset(portfolio, $index)" class="btn-remove">
                      Remover
                    </button>
                  </div>
                }
                <button type="button" (click)="addDetractorAsset(portfolio)" class="btn-add">
                  Adicionar Ativo Detrator
                </button>
              </div>

              <button type="button" (click)="removePortfolio($index)" class="btn-remove-portfolio">
                Remover Portfólio
              </button>
            </div>
          }
          
          <button type="button" (click)="addPortfolio()" class="btn-add-portfolio">
            Adicionar Portfólio
          </button>
        </div>

        <div class="actions">
          <button type="button" (click)="generateEmail()" class="btn-generate">
            Gerar Email
          </button>
          
          @if (generatedHtml()) {
            <button type="button" (click)="copyToClipboard()" class="btn-copy">
              Copiar HTML
            </button>
            
            <button type="button" (click)="downloadHtml()" class="btn-download">
              Download HTML
            </button>
          }
        </div>
      </div>

      @if (generatedHtml()) {
        <div class="preview-section">
          <h3>Preview do Email</h3>
          <div class="email-preview" [innerHTML]="generatedHtml()"></div>
        </div>
      }
    </div>
  `,
  styles: [`
    .email-generator-container {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 2rem;
      padding: 2rem;
      max-width: 1400px;
      margin: 0 auto;
    }

    .form-section {
      background: #f8f9fa;
      padding: 2rem;
      border-radius: 8px;
      overflow-y: auto;
      max-height: 90vh;
    }

    .preview-section {
      background: #ffffff;
      padding: 1rem;
      border-radius: 8px;
      border: 1px solid #dee2e6;
      overflow-y: auto;
      max-height: 90vh;
    }

    .form-group {
      margin-bottom: 1rem;
    }

    .form-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1rem;
    }

    label {
      display: block;
      margin-bottom: 0.5rem;
      font-weight: 600;
      color: #0D2035;
    }

    .form-control {
      width: 100%;
      padding: 0.5rem;
      border: 1px solid #ced4da;
      border-radius: 4px;
      font-size: 0.875rem;
    }

    .form-control:focus {
      outline: none;
      border-color: #0D2035;
      box-shadow: 0 0 0 2px rgba(13, 32, 53, 0.1);
    }

    .portfolio-form {
      background: #ffffff;
      padding: 1.5rem;
      border-radius: 8px;
      margin-bottom: 1.5rem;
      border: 1px solid #dee2e6;
    }

    .performance-row,
    .strategy-row,
    .asset-row {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr 1fr auto;
      gap: 0.5rem;
      align-items: center;
      margin-bottom: 0.5rem;
    }

    .strategy-row,
    .asset-row {
      grid-template-columns: 1fr auto;
    }

    .btn-add,
    .btn-remove,
    .btn-generate,
    .btn-copy,
    .btn-download,
    .btn-add-portfolio,
    .btn-remove-portfolio {
      padding: 0.5rem 1rem;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 0.875rem;
      font-weight: 500;
      transition: background-color 0.2s;
    }

    .btn-add,
    .btn-generate,
    .btn-add-portfolio {
      background-color: #28a745;
      color: white;
    }

    .btn-add:hover,
    .btn-generate:hover,
    .btn-add-portfolio:hover {
      background-color: #218838;
    }

    .btn-remove,
    .btn-remove-portfolio {
      background-color: #dc3545;
      color: white;
    }

    .btn-remove:hover,
    .btn-remove-portfolio:hover {
      background-color: #c82333;
    }

    .btn-copy,
    .btn-download {
      background-color: #0D2035;
      color: white;
    }

    .btn-copy:hover,
    .btn-download:hover {
      background-color: #1a3a5c;
    }

    .actions {
      display: flex;
      gap: 1rem;
      margin-top: 2rem;
    }

    .logo-preview {
      margin-top: 0.5rem;
      padding: 0.5rem;
      border: 1px solid #dee2e6;
      border-radius: 4px;
      background: white;
    }

    .email-preview {
      border: 1px solid #dee2e6;
      border-radius: 4px;
      min-height: 400px;
      background: white;
    }

    h2, h3, h4, h5 {
      color: #0D2035;
      margin-bottom: 1rem;
    }

    @media (max-width: 1200px) {
      .email-generator-container {
        grid-template-columns: 1fr;
      }
    }
  `]
})
export class EmailGeneratorComponent {
  private emailService = inject(OutlookCompatibleEmailService);

  clientName = signal('Vinicius Maciel');
  reportDate = signal(new Date().toISOString().split('T')[0]);
  logoBase64 = signal<string>('');
  portfolios = signal<PortfolioData[]>([this.createDefaultPortfolio()]);
  generatedHtml = signal<string>('');

  private createDefaultPortfolio(): PortfolioData {
    return {
      name: 'Offshore',
      type: 'Offshore',
      data: {
        performance: [
          { periodo: 'Junho:', carteira: -1.76, benchmark: 0.37, diferenca: -4.76 },
          { periodo: 'No ano:', carteira: -0.20, benchmark: 1.19, diferenca: -0.17 }
        ],
        retorno_financeiro: -17026.39,
        estrategias_destaque: ['FIXED INCOME (-0.05%)', 'CASH (23.53%)'],
        ativos_promotores: [
          'JUPITER GLOBAL EQUITY ABSOLUTE RETURN (+1.16%)',
          'VANGUARD EMERGING MARKETS ETF (VFEA) (+1.43%)'
        ],
        ativos_detratores: [
          'BLACKROCK WORLD TECHNOLOGY FUND A2 (-11.23%)',
          'URNJ - SPROTT JUNIOR URANIUM MINERS ETF (-8.94%)'
        ]
      }
    };
  }

  async onLogoUpload(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    
    if (file) {
      try {
        const base64 = await this.emailService.convertImageToBase64(file);
        this.logoBase64.set(base64);
      } catch (error) {
        console.error('Erro ao converter imagem:', error);
        alert('Erro ao processar a imagem. Tente novamente.');
      }
    }
  }

  addPortfolio(): void {
    this.portfolios.update(portfolios => [...portfolios, this.createDefaultPortfolio()]);
  }

  removePortfolio(index: number): void {
    this.portfolios.update(portfolios => portfolios.filter((_, i) => i !== index));
  }

  addPerformanceItem(portfolio: PortfolioData): void {
    portfolio.data.performance.push({
      periodo: '',
      carteira: 0,
      benchmark: 0,
      diferenca: 0
    });
    this.portfolios.update(portfolios => [...portfolios]);
  }

  removePerformanceItem(portfolio: PortfolioData, index: number): void {
    portfolio.data.performance.splice(index, 1);
    this.portfolios.update(portfolios => [...portfolios]);
  }

  addStrategy(portfolio: PortfolioData): void {
    portfolio.data.estrategias_destaque.push('');
    this.portfolios.update(portfolios => [...portfolios]);
  }

  removeStrategy(portfolio: PortfolioData, index: number): void {
    portfolio.data.estrategias_destaque.splice(index, 1);
    this.portfolios.update(portfolios => [...portfolios]);
  }

  addPromoterAsset(portfolio: PortfolioData): void {
    portfolio.data.ativos_promotores.push('');
    this.portfolios.update(portfolios => [...portfolios]);
  }

  removePromoterAsset(portfolio: PortfolioData, index: number): void {
    portfolio.data.ativos_promotores.splice(index, 1);
    this.portfolios.update(portfolios => [...portfolios]);
  }

  addDetractorAsset(portfolio: PortfolioData): void {
    portfolio.data.ativos_detratores.push('');
    this.portfolios.update(portfolios => [...portfolios]);
  }

  removeDetractorAsset(portfolio: PortfolioData, index: number): void {
    portfolio.data.ativos_detratores.splice(index, 1);
    this.portfolios.update(portfolios => [...portfolios]);
  }

  generateEmail(): void {
    const config: EmailConfiguration = {
      clientName: this.clientName(),
      dataRef: new Date(this.reportDate()),
      portfolios: this.portfolios(),
      logoBase64: this.logoBase64() || undefined
    };

    // Validar dados antes de gerar
    const invalidPortfolios = config.portfolios.filter(p => !this.emailService.validatePortfolioData(p));
    if (invalidPortfolios.length > 0) {
      alert('Alguns portfólios possuem dados incompletos. Verifique todos os campos obrigatórios.');
      return;
    }

    const html = this.emailService.generateOutlookCompatibleEmail(config);
    this.generatedHtml.set(html);
  }

  async copyToClipboard(): Promise<void> {
    try {
      await navigator.clipboard.writeText(this.generatedHtml());
      alert('HTML copiado para a área de transferência!');
    } catch (error) {
      console.error('Erro ao copiar:', error);
      // Fallback para navegadores mais antigos
      const textArea = document.createElement('textarea');
      textArea.value = this.generatedHtml();
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand('copy');
      document.body.removeChild(textArea);
      alert('HTML copiado para a área de transferência!');
    }
  }

  downloadHtml(): void {
    const blob = new Blob([this.generatedHtml()], { type: 'text/html' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    const date = new Date(this.reportDate());
    const fileName = `relatorio_mensal_${this.clientName()}_${date.getFullYear()}${(date.getMonth() + 1).toString().padStart(2, '0')}${date.getDate().toString().padStart(2, '0')}.html`;
    
    link.href = url;
    link.download = fileName;
    link.click();
    
    window.URL.revokeObjectURL(url);
  }
} 