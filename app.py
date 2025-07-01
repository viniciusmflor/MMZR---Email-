"""
MMZR Family Office - Interface Gráfica
Sistema de geração automatizada de relatórios financeiros

Autor: MMZR Family Office
Versão: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import sys
import os
from datetime import datetime
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Importar módulos do sistema MMZR
try:
    from gerador import gerar_relatorio_integrado, listar_clientes_disponiveis
    from diagnostico import verificar_status_sistema, executar_diagnostico_completo
except ImportError as e:
    logger.error(f"Erro ao importar módulos MMZR: {e}")
    sys.exit(1)


class MMZRInterface:
    """Interface gráfica principal do sistema MMZR."""
    
    def __init__(self):
        """Inicializa a interface gráfica."""
        self.root = tk.Tk()
        self.root.title("MMZR Family Office - Gerador de Relatórios v1.0")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)
        
        # Configurar estilo moderno
        self._configurar_estilo()
        
        # Variáveis de controle
        self.clientes_lista = []
        self.cliente_selecionado = tk.StringVar()
        self.enviar_email = tk.BooleanVar()
        self.processando = False
        
        # Criar interface
        self._criar_interface()
        self._centralizar_janela()
        
        # Inicializar sistema
        self._inicializar_sistema()
    
    def _configurar_estilo(self):
        """Configura estilo visual da aplicação."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurar cores corporativas
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#2E5BBA')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground='#666666')
        style.configure('Section.TLabel', font=('Segoe UI', 10, 'bold'))
        style.configure('Action.TButton', font=('Segoe UI', 9))
    
    def _centralizar_janela(self):
        """Centraliza a janela na tela."""
        self.root.update_idletasks()
        largura = self.root.winfo_width()
        altura = self.root.winfo_height()
        pos_x = (self.root.winfo_screenwidth() // 2) - (largura // 2)
        pos_y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
    
    def _criar_interface(self):
        """Cria todos os elementos da interface."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Cabeçalho
        self._criar_cabecalho(main_frame)
        
        # Seção de seleção de cliente
        self._criar_secao_cliente(main_frame)
        
        # Seção de opções
        self._criar_secao_opcoes(main_frame)
        
        # Seção de ações
        self._criar_secao_acoes(main_frame)
        
        # Seção de diagnóstico
        self._criar_secao_diagnostico(main_frame)
        
        # Área de log/status
        self._criar_area_status(main_frame)
        
        # Configurar redimensionamento
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def _criar_cabecalho(self, parent):
        """Cria o cabeçalho da aplicação."""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        
        titulo = ttk.Label(header_frame, text="MMZR Family Office", style='Title.TLabel')
        titulo.grid(row=0, column=0, sticky=tk.W)
        
        subtitulo = ttk.Label(header_frame, text="Sistema de Geração Automatizada de Relatórios", style='Subtitle.TLabel')
        subtitulo.grid(row=1, column=0, sticky=tk.W, pady=(2, 0))
    
    def _criar_secao_cliente(self, parent):
        """Cria seção de seleção de cliente."""
        cliente_frame = ttk.LabelFrame(parent, text="Seleção de Cliente", padding="15")
        cliente_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(cliente_frame, text="Cliente:", style='Section.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=(0, 8)
        )
        
        self.combo_clientes = ttk.Combobox(
            cliente_frame, textvariable=self.cliente_selecionado, 
            width=60, state="readonly", font=('Segoe UI', 9)
        )
        self.combo_clientes.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        cliente_frame.columnconfigure(0, weight=1)
    
    def _criar_secao_opcoes(self, parent):
        """Cria seção de opções."""
        opcoes_frame = ttk.LabelFrame(parent, text="Opções", padding="15")
        opcoes_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.check_email = ttk.Checkbutton(
            opcoes_frame, text="Preparar email no Microsoft Outlook", 
            variable=self.enviar_email
        )
        self.check_email.grid(row=0, column=0, sticky=tk.W)
    
    def _criar_secao_acoes(self, parent):
        """Cria seção de ações principais."""
        acoes_frame = ttk.LabelFrame(parent, text="Ações", padding="15")
        acoes_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        botoes_frame = ttk.Frame(acoes_frame)
        botoes_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.btn_gerar = ttk.Button(
            botoes_frame, text="Gerar Relatório", 
            command=self.gerar_relatorio, style='Action.TButton', width=20
        )
        self.btn_gerar.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_atualizar = ttk.Button(
            botoes_frame, text="Atualizar Lista", 
            command=self.carregar_clientes, style='Action.TButton', width=20
        )
        self.btn_atualizar.grid(row=0, column=1)
        
        acoes_frame.columnconfigure(0, weight=1)
    
    def _criar_secao_diagnostico(self, parent):
        """Cria seção de diagnóstico."""
        diag_frame = ttk.LabelFrame(parent, text="Diagnóstico", padding="15")
        diag_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        botoes_diag_frame = ttk.Frame(diag_frame)
        botoes_diag_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.btn_status = ttk.Button(
            botoes_diag_frame, text="Verificar Status", 
            command=self.verificar_status, style='Action.TButton', width=20
        )
        self.btn_status.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_diagnostico = ttk.Button(
            botoes_diag_frame, text="Diagnóstico Completo", 
            command=self.executar_diagnostico, style='Action.TButton', width=20
        )
        self.btn_diagnostico.grid(row=0, column=1)
        
        diag_frame.columnconfigure(0, weight=1)
    
    def _criar_area_status(self, parent):
        """Cria área de status e logs."""
        status_frame = ttk.LabelFrame(parent, text="Status do Sistema", padding="15")
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        
        self.text_status = scrolledtext.ScrolledText(
            status_frame, height=12, width=80, font=('Consolas', 9),
            wrap=tk.WORD, state=tk.NORMAL
        )
        self.text_status.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
    
    def _inicializar_sistema(self):
        """Inicializa o sistema e carrega dados iniciais."""
        self.adicionar_status("Sistema MMZR iniciado com sucesso")
        self.adicionar_status("Carregando lista de clientes...")
        self.carregar_clientes()
    
    def adicionar_status(self, mensagem):
        """Adiciona mensagem à área de status."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        linha = f"[{timestamp}] {mensagem}\n"
        
        self.text_status.insert(tk.END, linha)
        self.text_status.see(tk.END)
        self.root.update_idletasks()
        
        # Log também para arquivo
        logger.info(mensagem)
    
    def _desabilitar_botoes(self):
        """Desabilita todos os botões durante processamento."""
        self.btn_gerar.configure(state="disabled")
        self.btn_atualizar.configure(state="disabled")
        self.btn_status.configure(state="disabled")
        self.btn_diagnostico.configure(state="disabled")
        self.combo_clientes.configure(state="disabled")
    
    def _habilitar_botoes(self):
        """Habilita todos os botões após processamento."""
        self.btn_gerar.configure(state="normal")
        self.btn_atualizar.configure(state="normal")
        self.btn_status.configure(state="normal")
        self.btn_diagnostico.configure(state="normal")
        self.combo_clientes.configure(state="readonly")
    
    def carregar_clientes(self):
        """Carrega lista de clientes em thread separada."""
        if self.processando:
            return
        
        def carregar():
            self.processando = True
            self._desabilitar_botoes()
            self.adicionar_status("Iniciando carregamento de clientes...")
            
            try:
                # Capturar saída do sistema
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    clientes = listar_clientes_disponiveis()
                
                # Processar saída capturada
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip() and not line.startswith('=') and '|' not in line:
                        self.adicionar_status(f"Processando: {line}")
                
                self.clientes_lista = clientes
                self.combo_clientes['values'] = clientes
                
                if clientes:
                    self.adicionar_status(f"Carregamento concluído: {len(clientes)} cliente(s) disponível(is)")
                    self.combo_clientes.current(0)
                else:
                    self.adicionar_status("AVISO: Nenhum cliente encontrado nas planilhas")
                    
            except Exception as e:
                error_msg = f"Erro ao carregar clientes: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._habilitar_botoes()
        
        threading.Thread(target=carregar, daemon=True).start()
    
    def gerar_relatorio(self):
        """Gera relatório para o cliente selecionado."""
        if self.processando:
            return
        
        cliente = self.cliente_selecionado.get().strip()
        if not cliente:
            messagebox.showwarning("Seleção Necessária", "Por favor, selecione um cliente da lista")
            return
        
        def gerar():
            self.processando = True
            self._desabilitar_botoes()
            
            try:
                self.adicionar_status(f"Iniciando geração de relatório para: {cliente}")
                if self.enviar_email.get():
                    self.adicionar_status("Modo: Geração com preparação de email")
                else:
                    self.adicionar_status("Modo: Geração apenas de arquivo HTML")
                
                # Capturar saída do processo
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    gerar_relatorio_integrado(
                        nome_ou_email_cliente=cliente,
                        enviar_email=self.enviar_email.get()
                    )
                
                # Processar saída capturada
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip():
                        self.adicionar_status(line)
                
                self.adicionar_status("Processo de geração concluído com sucesso")
                messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso para {cliente}!")
                
            except Exception as e:
                error_msg = f"Erro na geração: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{error_msg}")
            
            finally:
                self.processando = False
                self._habilitar_botoes()
        
        threading.Thread(target=gerar, daemon=True).start()
    
    def verificar_status(self):
        """Verifica status do sistema."""
        if self.processando:
            return
        
        def verificar():
            self.processando = True
            self._desabilitar_botoes()
            
            try:
                self.adicionar_status("Executando verificação de status do sistema...")
                
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    resultado = verificar_status_sistema()
                
                # Processar saída capturada
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip():
                        self.adicionar_status(line)
                
                if resultado:
                    self.adicionar_status("Verificação concluída: Sistema funcionando corretamente")
                    messagebox.showinfo("Status", "Sistema funcionando corretamente!")
                else:
                    self.adicionar_status("Verificação concluída: Problemas detectados")
                    messagebox.showwarning("Status", "Sistema com problemas. Verifique o log de status.")
                
            except Exception as e:
                error_msg = f"Erro na verificação: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._habilitar_botoes()
        
        threading.Thread(target=verificar, daemon=True).start()
    
    def executar_diagnostico(self):
        """Executa diagnóstico completo."""
        if self.processando:
            return
        
        def diagnosticar():
            self.processando = True
            self._desabilitar_botoes()
            
            try:
                self.adicionar_status("Iniciando diagnóstico completo do sistema...")
                
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    resultado = executar_diagnostico_completo()
                
                # Processar saída capturada
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip():
                        self.adicionar_status(line)
                
                if resultado:
                    self.adicionar_status("Diagnóstico concluído: Sistema plenamente funcional")
                    messagebox.showinfo("Diagnóstico", "Diagnóstico concluído com sucesso!")
                else:
                    self.adicionar_status("Diagnóstico concluído: Problemas encontrados")
                    messagebox.showwarning("Diagnóstico", "Problemas encontrados. Verifique o log de status.")
                
            except Exception as e:
                error_msg = f"Erro no diagnóstico: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._habilitar_botoes()
        
        threading.Thread(target=diagnosticar, daemon=True).start()
    
    def executar(self):
        """Inicia a aplicação."""
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
            self.root.mainloop()
        except KeyboardInterrupt:
            self._on_closing()
    
    def _on_closing(self):
        """Gerencia fechamento da aplicação."""
        if self.processando:
            if messagebox.askokcancel("Fechar", "Existe um processo em andamento. Deseja realmente fechar?"):
                self.root.quit()
        else:
            self.root.quit()


def main():
    """Função principal da aplicação."""
    try:
        # Verificar se estamos no diretório correto
        if not os.path.exists("documentos"):
            messagebox.showerror(
                "Diretório Incorreto", 
                "Por favor, execute este programa na pasta raiz do projeto MMZR\n"
                "O diretório deve conter a pasta 'documentos'"
            )
            return False
        
        # Inicializar e executar aplicação
        logger.info("Iniciando aplicação MMZR v1.0")
        app = MMZRInterface()
        app.executar()
        logger.info("Aplicação encerrada")
        return True
        
    except Exception as e:
        error_msg = f"Erro fatal ao iniciar aplicação: {str(e)}"
        logger.error(error_msg)
        messagebox.showerror("Erro Fatal", error_msg)
        return False


if __name__ == "__main__":
    main() 