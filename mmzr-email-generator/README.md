# MMZR Email Generator - Compatível com Outlook

Uma aplicação Angular/TypeScript especializada em gerar emails HTML totalmente compatíveis com Microsoft Outlook e outros clientes de email.

## 🚀 Características Principais

### ✅ Compatibilidade Total com Outlook
- **Estilos inline**: Todos os estilos são aplicados diretamente nos elementos HTML
- **Estrutura de tabelas**: Layout baseado em tabelas para máxima compatibilidade
- **Conditional Comments**: Suporte específico para versões do Outlook
- **VML (Vector Markup Language)**: Botões e elementos gráficos compatíveis com Outlook
- **Fallbacks**: Alternativas para funcionalidades não suportadas

### 🎯 Funcionalidades
- Geração de relatórios mensais de performance
- Interface visual para criação de emails
- Preview em tempo real
- Exportação em HTML otimizado
- Upload e conversão de logos para Base64
- Validação de dados antes da geração

### 🛠️ Tecnologias Utilizadas
- **Angular 17+** com Standalone Components
- **TypeScript** com strict type checking
- **SCSS** para estilização
- **Signals** para gerenciamento de estado reativo
- **Inject function** para injeção de dependências

## 📋 Pré-requisitos

- Node.js (versão 18 ou superior)
- npm ou yarn
- Angular CLI (opcional, mas recomendado)

## 🔧 Instalação

1. **Clone o repositório:**
```bash
git clone <url-do-repositorio>
cd mmzr-email-generator
```

2. **Instale as dependências:**
```bash
npm install
```

3. **Execute a aplicação:**
```bash
npm start
```

4. **Acesse no navegador:**
```
http://localhost:4200
```

## 📖 Como Usar

### 1. Configuração Básica
- **Nome do Cliente**: Digite o nome que aparecerá na saudação
- **Data do Relatório**: Selecione a data de referência
- **Logo**: Faça upload da logo da empresa (será convertida para Base64)

### 2. Configuração de Portfólios
- **Adicionar Portfólio**: Clique em "Adicionar Portfólio" para criar novos
- **Nome e Tipo**: Defina o nome e tipo de cada portfólio
- **Performance**: Adicione dados de performance (período, carteira, benchmark, diferença)
- **Retorno Financeiro**: Informe o valor de retorno em reais

### 3. Ativos e Estratégias
- **Estratégias de Destaque**: Liste as principais estratégias com suas performances
- **Ativos Promotores**: Adicione ativos com performance positiva
- **Ativos Detratores**: Adicione ativos com performance negativa

### 4. Geração e Exportação
- **Gerar Email**: Clique para gerar o HTML do email
- **Preview**: Visualize o resultado na seção de preview
- **Copiar HTML**: Copie o código HTML para a área de transferência
- **Download**: Baixe o arquivo HTML para uso posterior

## 🔧 Estrutura Técnica

### Serviço Principal: `OutlookCompatibleEmailService`

```typescript
export class OutlookCompatibleEmailService {
  generateOutlookCompatibleEmail(config: EmailConfiguration): string
  validatePortfolioData(portfolio: PortfolioData): boolean
  convertImageToBase64(file: File): Promise<string>
  generateEmailSubject(dataRef: Date): string
}
```

### Interfaces TypeScript

```typescript
interface EmailConfiguration {
  clientName: string;
  dataRef: Date;
  portfolios: PortfolioData[];
  logoBase64?: string;
  customFooter?: string;
}

interface PortfolioData {
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
```

## 📧 Compatibilidade de Email

### ✅ Clientes Suportados
- **Microsoft Outlook** (2007, 2010, 2013, 2016, 2019, 365)
- **Outlook.com** (web)
- **Gmail** (web e app)
- **Apple Mail** (macOS e iOS)
- **Yahoo Mail**
- **Thunderbird**
- **Android Email**

### 🎨 Técnicas de Compatibilidade Implementadas

#### 1. Estilos Inline
```html
<td style="background-color: #0D2035; color: #ffffff; padding: 12px;">
```

#### 2. Conditional Comments para Outlook
```html
<!--[if mso]>
<style type="text/css">
  body, table, td { font-family: Arial, sans-serif !important; }
</style>
<![endif]-->
```

#### 3. Estrutura de Tabelas
```html
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td>Conteúdo aqui</td>
  </tr>
</table>
```

#### 4. Botões VML para Outlook
```html
<!--[if mso]>
<v:roundrect href="link" style="height:44px;width:280px;" fillcolor="#0D2035">
  <center>Texto do Botão</center>
</v:roundrect>
<![endif]-->
```

## 🎯 Melhores Práticas Implementadas

### 1. **Reset CSS Específico para Email**
- Margin e padding zerados
- Box-sizing border-box
- Font-family consistente

### 2. **Estrutura Responsiva**
- Media queries para dispositivos móveis
- Larguras flexíveis
- Fontes escaláveis

### 3. **Imagens Otimizadas**
- Conversão automática para Base64
- Alt text para acessibilidade
- Dimensões fixas para estabilidade

### 4. **Cores e Contrastes**
- Paleta de cores consistente
- Alto contraste para legibilidade
- Cores seguras para email

## 🔍 Debugging e Testes

### Testando Compatibilidade
1. **Teste no Outlook Desktop**: Envie o email para uma conta Outlook
2. **Teste no Gmail**: Verifique renderização no Gmail web
3. **Teste em Dispositivos Móveis**: Confirme responsividade
4. **Validação HTML**: Use validadores específicos para email

### Ferramentas Recomendadas
- **Litmus**: Teste em múltiplos clientes
- **Email on Acid**: Validação de compatibilidade
- **PutsMail**: Teste gratuito de emails
- **Mail Tester**: Verificação de spam score

## 🚨 Problemas Comuns e Soluções

### Outlook não exibe cores de fundo
**Solução**: Use tabelas aninhadas com estilos inline
```html
<table><tr><td style="background-color: #color;">Conteúdo</td></tr></table>
```

### Gmail remove estilos CSS
**Solução**: Todos os estilos foram convertidos para inline

### Imagens quebradas
**Solução**: Logos convertidas para Base64 embutido

### Botões não funcionam no Outlook
**Solução**: Implementação VML com fallback HTML

## 📝 Personalização

### Modificando Cores
Edite as constantes no serviço:
```typescript
private readonly corPrimaria = '#0D2035';
private readonly corSuccesso = '#28a745';
private readonly corPerigo = '#dc3545';
```

### Adicionando Novos Campos
1. Atualize a interface `PortfolioData`
2. Modifique o método `gerarSecaoPortfolio`
3. Adicione campos no componente

### Customizando Layout
Edite os métodos privados no `OutlookCompatibleEmailService`:
- `gerarCabecalho()`
- `gerarRodape()`
- `gerarTabelaPerformance()`

## 🤝 Contribuindo

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 📞 Suporte

Para dúvidas ou problemas:
- Abra uma issue no GitHub
- Entre em contato com a equipe de desenvolvimento

---

**Desenvolvido com ❤️ pela equipe MMZR Family Office** 