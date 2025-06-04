import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
import { EmailGeneratorComponent } from './components/email-generator/email-generator.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet, EmailGeneratorComponent],
  template: `
    <div class="app-container">
      <header class="app-header">
        <div class="header-content">
          <h1>MMZR Email Generator</h1>
          <p>Gerador de emails HTML compatível com Outlook e outros clientes de email</p>
        </div>
      </header>
      
      <main class="app-main">
        <app-email-generator />
      </main>
      
      <footer class="app-footer">
        <p>&copy; 2025 MMZR Family Office. Desenvolvido com Angular e TypeScript.</p>
      </footer>
    </div>
  `,
  styles: [`
    .app-container {
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }

    .app-header {
      background: linear-gradient(135deg, #0D2035 0%, #1a3a5c 100%);
      color: white;
      padding: 2rem 0;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
    }

    .header-content {
      max-width: 1200px;
      margin: 0 auto;
      padding: 0 2rem;
      text-align: center;
    }

    .header-content h1 {
      margin: 0 0 0.5rem 0;
      font-size: 2.5rem;
      font-weight: 700;
      letter-spacing: -0.025em;
    }

    .header-content p {
      margin: 0;
      font-size: 1.125rem;
      opacity: 0.9;
      font-weight: 300;
    }

    .app-main {
      flex: 1;
      background: #f8f9fa;
      min-height: calc(100vh - 200px);
    }

    .app-footer {
      background: #343a40;
      color: white;
      text-align: center;
      padding: 1rem;
      margin-top: auto;
    }

    .app-footer p {
      margin: 0;
      font-size: 0.875rem;
      opacity: 0.8;
    }

    @media (max-width: 768px) {
      .header-content h1 {
        font-size: 2rem;
      }

      .header-content p {
        font-size: 1rem;
      }

      .header-content {
        padding: 0 1rem;
      }
    }
  `]
})
export class AppComponent {
  title = 'MMZR Email Generator';
} 