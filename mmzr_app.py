import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import json
import webbrowser

# Importar o gerador de e-mail
from mmzr_email_generator import MMZREmailGenerator, process_and_generate_report

class MMZRApp:
    """Aplicativo para geração de relatórios da MMZR Family Office"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("MMZR Family Office - Gerador de Relatórios")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # Estilo
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TButton", background="#061844", foreground="#ffffff", font=("Arial", 10))
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 14, "bold"), background="#f0f0f0")
        self.style.configure("Subheader.TLabel", font=("Arial", 12, "bold"), background="#f0f0f0")
        
        # Variáveis
        self.excel_path = tk.StringVar()
        self.client_name = tk.StringVar()
        self.client_email = tk.StringVar()
        
        self.portfolios = []
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, padding=20, style="TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Interface
        self.create_header()
        self.create_file_selector()
        self.create_client_info()
        self.create_portfolios_section()
        self.create_action_buttons()
        
        # Carregar configurações salvas
        self.load_config()
    
    def create_header(self):
        """Cria o cabeçalho do aplicativo"""
        header_frame = ttk.Frame(self.main_frame, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="Gerador de Relatórios de Performance", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(header_frame, text="Configure os parâmetros e gere relatórios mensais para seus clientes").pack(anchor=tk.W)
    
    def create_file_selector(self):
        """Cria o seletor de arquivo Excel"""
        file_frame = ttk.Frame(self.main_frame, style="TFrame")
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(file_frame, text="Arquivo Excel com dados:", style="TLabel").pack(anchor=tk.W)
        
        file_select_frame = ttk.Frame(file_frame, style="TFrame")
        file_select_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.excel_entry = ttk.Entry(file_select_frame, textvariable=self.excel_path, width=70)
        self.excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(file_select_frame, text="Procurar", command=self.browse_excel).pack(side=tk.RIGHT, padx=(10, 0))
    
    def create_client_info(self):
        """Cria os campos para informações do cliente"""
        client_frame = ttk.Frame(self.main_frame, style="TFrame")
        client_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(client_frame, text="Informações do Cliente", style="Subheader.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        info_frame = ttk.Frame(client_frame, style="TFrame")
        info_frame.pack(fill=tk.X)
        
        # Grid para os campos
        info_frame.columnconfigure(1, weight=1)
        
        # Nome do Cliente
        ttk.Label(info_frame, text="Nome:", style="TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        ttk.Entry(info_frame, textvariable=self.client_name, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, pady=5)
        
        # Email do Cliente
        ttk.Label(info_frame, text="Email:", style="TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        ttk.Entry(info_frame, textvariable=self.client_email, width=50).grid(row=1, column=1, sticky=tk.W+tk.E, pady=5)
    
    def create_portfolios_section(self):
        """Cria a seção para gerenciar carteiras"""
        portfolios_frame = ttk.Frame(self.main_frame, style="TFrame")
        portfolios_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        ttk.Label(portfolios_frame, text="Carteiras", style="Subheader.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # Frame para lista de carteiras
        self.portfolios_list_frame = ttk.Frame(portfolios_frame, style="TFrame")
        self.portfolios_list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Botão para adicionar carteira
        ttk.Button(portfolios_frame, text="Adicionar Carteira", command=self.add_portfolio).pack(anchor=tk.W, pady=(10, 0))
    
    def create_action_buttons(self):
        """Cria os botões de ação no rodapé"""
        buttons_frame = ttk.Frame(self.main_frame, style="TFrame")
        buttons_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(buttons_frame, text="Salvar Configuração", command=self.save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Gerar Relatório", command=self.generate_report).pack(side=tk.RIGHT)
    
    def browse_excel(self):
        """Abre diálogo para selecionar arquivo Excel"""
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
    
    def add_portfolio(self):
        """Adiciona uma nova carteira à lista"""
        portfolio = {
            'name': f"Carteira {len(self.portfolios) + 1}",
            'type': "Diversificada",
            'sheet_name': "",
            'benchmark_name': "IPCA+5%"
        }
        self.portfolios.append(portfolio)
        self.refresh_portfolios_list()
    
    def remove_portfolio(self, index):
        """Remove uma carteira da lista"""
        if 0 <= index < len(self.portfolios):
            del self.portfolios[index]
            self.refresh_portfolios_list()
    
    def refresh_portfolios_list(self):
        """Atualiza a interface com a lista de carteiras"""
        # Limpar o frame
        for widget in self.portfolios_list_frame.winfo_children():
            widget.destroy()
        
        # Adicionar cada carteira
        for i, portfolio in enumerate(self.portfolios):
            portfolio_frame = ttk.Frame(self.portfolios_list_frame, style="TFrame")
            portfolio_frame.pack(fill=tk.X, pady=(0, 10))
            
            # Nome da carteira
            name_frame = ttk.Frame(portfolio_frame, style="TFrame")
            name_frame.pack(fill=tk.X, pady=(0, 5))
            
            ttk.Label(name_frame, text=f"Carteira {i+1}: Nome", style="TLabel").pack(side=tk.LEFT)
            name_entry = ttk.Entry(name_frame, width=30)
            name_entry.insert(0, portfolio['name'])
            name_entry.pack(side=tk.LEFT, padx=(10, 0))
            
            # Botão remover
            ttk.Button(name_frame, text="Remover", command=lambda idx=i: self.remove_portfolio(idx)).pack(side=tk.RIGHT)
            
            # Detalhes da carteira
            details_frame = ttk.Frame(portfolio_frame, style="TFrame")
            details_frame.pack(fill=tk.X)
            
            # Grid para os campos
            details_frame.columnconfigure(1, weight=1)
            details_frame.columnconfigure(3, weight=1)
            
            # Tipo
            ttk.Label(details_frame, text="Tipo:", style="TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 5), pady=5)
            type_entry = ttk.Entry(details_frame, width=20)
            type_entry.insert(0, portfolio['type'])
            type_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 10), pady=5)
            
            # Aba da planilha
            ttk.Label(details_frame, text="Aba Excel:", style="TLabel").grid(row=0, column=2, sticky=tk.W, padx=(0, 5), pady=5)
            sheet_entry = ttk.Entry(details_frame, width=20)
            sheet_entry.insert(0, portfolio['sheet_name'])
            sheet_entry.grid(row=0, column=3, sticky=tk.W, pady=5)
            
            # Benchmark
            ttk.Label(details_frame, text="Benchmark:", style="TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=5)
            benchmark_entry = ttk.Entry(details_frame, width=20)
            benchmark_entry.insert(0, portfolio['benchmark_name'])
            benchmark_entry.grid(row=1, column=1, sticky=tk.W, padx=(0, 10), pady=5)
            
            # Atualizar valores quando os campos forem editados
            def update_portfolio(event=None):
                self.portfolios[i]['name'] = name_entry.get()
                self.portfolios[i]['type'] = type_entry.get()
                self.portfolios[i]['sheet_name'] = sheet_entry.get()
                self.portfolios[i]['benchmark_name'] = benchmark_entry.get()
            
            name_entry.bind("<FocusOut>", update_portfolio)
            type_entry.bind("<FocusOut>", update_portfolio)
            sheet_entry.bind("<FocusOut>", update_portfolio)
            benchmark_entry.bind("<FocusOut>", update_portfolio)
    
    def save_config(self):
        """Salva a configuração atual em um arquivo JSON"""
        config = {
            'excel_path': self.excel_path.get(),
            'client_name': self.client_name.get(),
            'client_email': self.client_email.get(),
            'portfolios': self.portfolios
        }
        
        try:
            with open('mmzr_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Configuração Salva", "As configurações foram salvas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configuração: {str(e)}")
    
    def load_config(self):
        """Carrega a configuração salva de um arquivo JSON"""
        try:
            if os.path.exists('mmzr_config.json'):
                with open('mmzr_config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                self.excel_path.set(config.get('excel_path', ''))
                self.client_name.set(config.get('client_name', ''))
                self.client_email.set(config.get('client_email', ''))
                
                self.portfolios = config.get('portfolios', [])
                if not self.portfolios:
                    # Adicionar carteira padrão se não houver nenhuma
                    self.portfolios = [{
                        'name': 'Carteira Moderada',
                        'type': 'Renda Variável + Renda Fixa',
                        'sheet_name': 'Base Consolidada',
                        'benchmark_name': 'IPCA+5%'
                    }]
                
                self.refresh_portfolios_list()
        except Exception as e:
            messagebox.showwarning("Aviso", f"Erro ao carregar configuração: {str(e)}")
            # Adicionar carteira padrão
            self.portfolios = [{
                'name': 'Carteira Moderada',
                'type': 'Renda Variável + Renda Fixa',
                'sheet_name': 'Base Consolidada',
                'benchmark_name': 'IPCA+5%'
            }]
            self.refresh_portfolios_list()
    
    def generate_report(self):
        """Gera o relatório com base nas configurações atuais"""
        # Verificar se os campos obrigatórios estão preenchidos
        if not self.excel_path.get():
            messagebox.showerror("Erro", "Selecione o arquivo Excel com os dados.")
            return
        
        if not self.client_name.get():
            messagebox.showerror("Erro", "Informe o nome do cliente.")
            return
        
        if not self.portfolios:
            messagebox.showerror("Erro", "Adicione pelo menos uma carteira.")
            return
        
        # Verificar se o arquivo existe
        if not os.path.exists(self.excel_path.get()):
            messagebox.showerror("Erro", f"O arquivo {self.excel_path.get()} não foi encontrado.")
            return
        
        # Configurar o cliente
        client_config = {
            'name': self.client_name.get(),
            'email': self.client_email.get(),
            'portfolios': self.portfolios
        }
        
        try:
            # Processar e gerar o relatório
            result = process_and_generate_report(self.excel_path.get(), client_config)
            
            if result:
                response = messagebox.askyesno(
                    "Relatório Gerado",
                    f"O relatório foi gerado com sucesso!\nDeseja abrir o arquivo agora?",
                    icon=messagebox.INFO
                )
                
                if response:
                    # Abrir o arquivo no navegador padrão
                    webbrowser.open('file://' + os.path.abspath(result))
            else:
                messagebox.showerror("Erro", "Falha ao gerar o relatório. Verifique os logs para mais detalhes.")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o relatório: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MMZRApp(root)
    root.mainloop() 