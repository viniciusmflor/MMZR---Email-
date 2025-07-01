"""
MMZR Family Office - Interface Grafica Principal
Sistema de geracao automatizada de relatorios financeiros

Versao: 1.0.0
Plataforma: Windows (Microsoft Outlook)
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

# Importar modulos do sistema MMZR
try:
    from gerador import gerar_relatorio_integrado, listar_clientes_disponiveis
    from diagnostico import verificar_status_sistema, executar_diagnostico_completo
    from mmzr_email_sender import verificar_outlook_disponivel
except ImportError as e:
    logger.error(f"Erro ao importar modulos MMZR: {e}")
    sys.exit(1)


class MMZRInterface:
    """Interface grafica principal do sistema MMZR."""
    
    def __init__(self):
        """Inicializa a interface grafica."""
        self.root = tk.Tk()
        self.root.title("MMZR Family Office - Gerador de Relatorios v1.0")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        self.root.minsize(700, 600)
        
        # Configurar estilo
        self._configurar_estilo()
        
        # Variaveis de controle
        self.clientes_lista = []
        self.enviar_email = tk.BooleanVar()
        self.outlook_disponivel = tk.BooleanVar()
        self.processando = False
        
        # Criar interface
        self._criar_interface()
        self._centralizar_janela()
        
        # Inicializar sistema
        self._inicializar_sistema()
    
    def _configurar_estilo(self):
        """Configura estilo visual da aplicacao."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurar cores corporativas
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#2E5BBA')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground='#666666')
        style.configure('Section.TLabel', font=('Segoe UI', 10, 'bold'))
        style.configure('Action.TButton', font=('Segoe UI', 9))
        style.configure('Success.TButton', font=('Segoe UI', 9, 'bold'))
    
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
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self._criar_cabecalho(main_frame)
        self._criar_secao_clientes(main_frame)
        self._criar_secao_opcoes(main_frame)
        self._criar_secao_acoes(main_frame)
        self._criar_secao_diagnostico(main_frame)
        self._criar_area_status(main_frame)
        
        # Configurar redimensionamento
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def _criar_cabecalho(self, parent):
        """Cria o cabecalho da aplicacao."""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        
        titulo = ttk.Label(header_frame, text="MMZR Family Office", style='Title.TLabel')
        titulo.grid(row=0, column=0, sticky=tk.W)
        
        subtitulo = ttk.Label(header_frame, text="Sistema de Geracao Automatizada de Relatorios", style='Subtitle.TLabel')
        subtitulo.grid(row=1, column=0, sticky=tk.W, pady=(2, 0))
    
    def _criar_secao_clientes(self, parent):
        """Cria secao de selecao de clientes."""
        clientes_frame = ttk.LabelFrame(parent, text="Selecao de Clientes", padding="15")
        clientes_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # Controles de selecao
        controles_frame = ttk.Frame(clientes_frame)
        controles_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(controles_frame, text="Clientes Disponiveis:", style='Section.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        # Botoes de controle
        botoes_frame = ttk.Frame(controles_frame)
        botoes_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.btn_selecionar_todos = ttk.Button(
            botoes_frame, text="Selecionar Todos", 
            command=self.selecionar_todos_clientes, style='Action.TButton', width=15
        )
        self.btn_selecionar_todos.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_limpar_selecao = ttk.Button(
            botoes_frame, text="Limpar Selecao", 
            command=self.limpar_selecao_clientes, style='Action.TButton', width=15
        )
        self.btn_limpar_selecao.grid(row=0, column=1)
        
        # Lista de clientes
        lista_frame = ttk.Frame(clientes_frame)
        lista_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.listbox_clientes = tk.Listbox(
            lista_frame, 
            selectmode=tk.EXTENDED,
            height=8,
            font=('Segoe UI', 9)
        )
        
        scrollbar_clientes = ttk.Scrollbar(lista_frame, orient=tk.VERTICAL, command=self.listbox_clientes.yview)
        self.listbox_clientes.configure(yscrollcommand=scrollbar_clientes.set)
        
        self.listbox_clientes.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_clientes.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Status da selecao
        self.label_selecao = ttk.Label(clientes_frame, text="Nenhum cliente selecionado", foreground="#666666")
        self.label_selecao.grid(row=3, column=0, sticky=tk.W)
        
        self.listbox_clientes.bind('<<ListboxSelect>>', self._atualizar_contador_selecao)
        
        # Configurar redimensionamento
        clientes_frame.columnconfigure(0, weight=1)
        clientes_frame.rowconfigure(2, weight=1)
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
    
    def _criar_secao_opcoes(self, parent):
        """Cria secao de opcoes."""
        opcoes_frame = ttk.LabelFrame(parent, text="Opcoes de Geracao", padding="15")
        opcoes_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.check_email = ttk.Checkbutton(
            opcoes_frame, text="Preparar emails no Microsoft Outlook", 
            variable=self.enviar_email
        )
        self.check_email.grid(row=0, column=0, sticky=tk.W)
        
        self.label_outlook_status = ttk.Label(
            opcoes_frame, text="Verificando Outlook...", foreground="#666666"
        )
        self.label_outlook_status.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
    
    def _criar_secao_acoes(self, parent):
        """Cria secao de acoes principais."""
        acoes_frame = ttk.LabelFrame(parent, text="Acoes", padding="15")
        acoes_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        botoes_frame = ttk.Frame(acoes_frame)
        botoes_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.btn_gerar = ttk.Button(
            botoes_frame, text="Gerar Relatorios", 
            command=self.gerar_relatorios, style='Success.TButton', width=20
        )
        self.btn_gerar.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_atualizar = ttk.Button(
            botoes_frame, text="Atualizar Lista", 
            command=self.carregar_clientes, style='Action.TButton', width=20
        )
        self.btn_atualizar.grid(row=0, column=1)
        
        acoes_frame.columnconfigure(0, weight=1)
    
    def _criar_secao_diagnostico(self, parent):
        """Cria secao de diagnostico."""
        diag_frame = ttk.LabelFrame(parent, text="Diagnostico do Sistema", padding="15")
        diag_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        botoes_diag_frame = ttk.Frame(diag_frame)
        botoes_diag_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.btn_status = ttk.Button(
            botoes_diag_frame, text="Verificar Status", 
            command=self.verificar_status, style='Action.TButton', width=20
        )
        self.btn_status.grid(row=0, column=0, padx=(0, 10))
        
        self.btn_diagnostico = ttk.Button(
            botoes_diag_frame, text="Diagnostico Completo", 
            command=self.executar_diagnostico, style='Action.TButton', width=20
        )
        self.btn_diagnostico.grid(row=0, column=1)
        
        diag_frame.columnconfigure(0, weight=1)
    
    def _criar_area_status(self, parent):
        """Cria area de status e logs."""
        status_frame = ttk.LabelFrame(parent, text="Log do Sistema", padding="15")
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        
        self.text_status = scrolledtext.ScrolledText(
            status_frame, height=12, width=80, font=('Consolas', 9),
            wrap=tk.WORD, state=tk.NORMAL
        )
        self.text_status.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
    
    def _inicializar_sistema(self):
        """Inicializa o sistema e carrega dados."""
        self.adicionar_status("Inicializando sistema MMZR...")
        threading.Thread(target=self._verificar_outlook, daemon=True).start()
        self.carregar_clientes()
    
    def _verificar_outlook(self):
        """Verifica se o Outlook esta disponivel."""
        try:
            outlook_ok = verificar_outlook_disponivel()
            self.outlook_disponivel.set(outlook_ok)
            
            if outlook_ok:
                self.label_outlook_status.config(text="Microsoft Outlook disponivel", foreground="green")
                self.adicionar_status("Microsoft Outlook detectado e funcional")
            else:
                self.label_outlook_status.config(text="Microsoft Outlook nao disponivel", foreground="red")
                self.adicionar_status("AVISO: Microsoft Outlook nao foi detectado")
                
        except Exception as e:
            self.label_outlook_status.config(text="Erro na verificacao do Outlook", foreground="red")
            self.adicionar_status(f"Erro ao verificar Outlook: {str(e)}")
    
    def adicionar_status(self, mensagem):
        """Adiciona mensagem a area de status."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        mensagem_completa = f"[{timestamp}] {mensagem}\n"
        
        self.text_status.insert(tk.END, mensagem_completa)
        self.text_status.see(tk.END)
        self.root.update_idletasks()
    
    def _controlar_botoes(self, estado):
        """Controla estado dos botoes durante processamento."""
        widgets = [
            self.btn_gerar, self.btn_atualizar, self.btn_status, 
            self.btn_diagnostico, self.btn_selecionar_todos, self.btn_limpar_selecao
        ]
        for widget in widgets:
            widget.config(state=estado)
    
    def selecionar_todos_clientes(self):
        """Seleciona todos os clientes da lista."""
        self.listbox_clientes.select_set(0, tk.END)
        self._atualizar_contador_selecao()
    
    def limpar_selecao_clientes(self):
        """Limpa a selecao de clientes."""
        self.listbox_clientes.selection_clear(0, tk.END)
        self._atualizar_contador_selecao()
    
    def _atualizar_contador_selecao(self, event=None):
        """Atualiza o contador de clientes selecionados."""
        qtd = len(self.listbox_clientes.curselection())
        
        if qtd == 0:
            texto = "Nenhum cliente selecionado"
        elif qtd == 1:
            texto = "1 cliente selecionado"
        else:
            texto = f"{qtd} clientes selecionados"
        
        self.label_selecao.config(text=texto)
    
    def carregar_clientes(self):
        """Carrega lista de clientes disponiveis."""
        if self.processando:
            return
        
        def carregar():
            self.processando = True
            self._controlar_botoes('disabled')
            
            try:
                self.adicionar_status("Carregando lista de clientes...")
                self.listbox_clientes.delete(0, tk.END)
                
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    clientes = listar_clientes_disponiveis()
                
                self.clientes_lista = clientes
                
                for cliente in clientes:
                    self.listbox_clientes.insert(tk.END, cliente)
                
                if clientes:
                    self.adicionar_status(f"Carregamento concluido: {len(clientes)} cliente(s) disponivel(is)")
                    self.listbox_clientes.select_set(0)
                    self._atualizar_contador_selecao()
                else:
                    self.adicionar_status("AVISO: Nenhum cliente encontrado nas planilhas")
                    
            except Exception as e:
                error_msg = f"Erro ao carregar clientes: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._controlar_botoes('normal')
        
        threading.Thread(target=carregar, daemon=True).start()
    
    def gerar_relatorios(self):
        """Gera relatorios para os clientes selecionados."""
        if self.processando:
            return
        
        indices_selecionados = self.listbox_clientes.curselection()
        if not indices_selecionados:
            messagebox.showwarning("Selecao Necessaria", "Por favor, selecione pelo menos um cliente da lista")
            return
        
        clientes_selecionados = [self.clientes_lista[i] for i in indices_selecionados]
        qtd_clientes = len(clientes_selecionados)
        
        # Confirmar acao para multiplos clientes
        if qtd_clientes > 1:
            resposta = messagebox.askyesno(
                "Confirmar Geracao", 
                f"Gerar relatorios para {qtd_clientes} clientes?\n\n"
                f"Clientes selecionados:\n" + "\n".join(f"• {cliente}" for cliente in clientes_selecionados[:5]) +
                (f"\n... e mais {qtd_clientes - 5}" if qtd_clientes > 5 else "")
            )
            if not resposta:
                return
        
        def gerar():
            self.processando = True
            self._controlar_botoes('disabled')
            
            try:
                self.adicionar_status(f"Iniciando geracao de relatorios para {qtd_clientes} cliente(s)")
                
                if self.enviar_email.get():
                    self.adicionar_status("Modo: Geracao com preparacao de emails no Microsoft Outlook")
                else:
                    self.adicionar_status("Modo: Geracao apenas de arquivos HTML")
                
                sucessos = 0
                falhas = 0
                
                for i, cliente in enumerate(clientes_selecionados, 1):
                    self.adicionar_status(f"Processando cliente {i}/{qtd_clientes}: {cliente}")
                    
                    try:
                        import io
                        import contextlib
                        
                        output_buffer = io.StringIO()
                        with contextlib.redirect_stdout(output_buffer):
                            sucesso = gerar_relatorio_integrado(
                                nome_ou_email_cliente=cliente,
                                enviar_email=self.enviar_email.get() and self.outlook_disponivel.get()
                            )
                        
                        output_lines = output_buffer.getvalue().strip().split('\n')
                        for line in output_lines:
                            if line.strip():
                                self.adicionar_status(f"  {line}")
                        
                        if sucesso:
                            sucessos += 1
                            self.adicionar_status(f"Relatorio gerado com sucesso para {cliente}")
                        else:
                            falhas += 1
                            self.adicionar_status(f"Falha na geracao para {cliente}")
                            
                    except Exception as e:
                        falhas += 1
                        self.adicionar_status(f"Erro ao processar {cliente}: {str(e)}")
                
                # Exibir resumo
                self.adicionar_status(f"\nResumo da geracao:")
                self.adicionar_status(f"  Sucessos: {sucessos}")
                self.adicionar_status(f"  Falhas: {falhas}")
                self.adicionar_status(f"  Total: {qtd_clientes}")
                
                if sucessos > 0:
                    messagebox.showinfo(
                        "Processo Concluido", 
                        f"Geracao concluida!\n\n"
                        f"Sucessos: {sucessos}\n"
                        f"Falhas: {falhas}\n"
                        f"Total: {qtd_clientes}"
                    )
                else:
                    messagebox.showerror(
                        "Processo com Problemas",
                        "Nenhum relatorio foi gerado com sucesso.\n\n"
                        "Verifique o log do sistema para mais detalhes."
                    )
                
            except Exception as e:
                error_msg = f"Erro no processo de geracao: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._controlar_botoes('normal')
        
        threading.Thread(target=gerar, daemon=True).start()
    
    def verificar_status(self):
        """Verifica status do sistema."""
        if self.processando:
            return
        
        def verificar():
            self.processando = True
            self._controlar_botoes('disabled')
            
            try:
                self.adicionar_status("Executando verificacao de status do sistema...")
                
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    resultado = verificar_status_sistema()
                
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip():
                        self.adicionar_status(line)
                
                if resultado:
                    self.adicionar_status("Verificacao concluida: Sistema funcionando corretamente")
                    messagebox.showinfo("Status", "Sistema funcionando corretamente!")
                else:
                    self.adicionar_status("Verificacao concluida: Problemas detectados")
                    messagebox.showwarning("Status", "Sistema com problemas. Verifique o log de status.")
                
            except Exception as e:
                error_msg = f"Erro na verificacao: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._controlar_botoes('normal')
        
        threading.Thread(target=verificar, daemon=True).start()
    
    def executar_diagnostico(self):
        """Executa diagnostico completo."""
        if self.processando:
            return
        
        def diagnosticar():
            self.processando = True
            self._controlar_botoes('disabled')
            
            try:
                self.adicionar_status("Iniciando diagnostico completo do sistema...")
                
                import io
                import contextlib
                
                output_buffer = io.StringIO()
                with contextlib.redirect_stdout(output_buffer):
                    resultado = executar_diagnostico_completo()
                
                output_lines = output_buffer.getvalue().strip().split('\n')
                for line in output_lines:
                    if line.strip():
                        self.adicionar_status(line)
                
                if resultado:
                    self.adicionar_status("Diagnostico concluido: Sistema plenamente funcional")
                    messagebox.showinfo("Diagnostico", "Diagnostico concluido com sucesso!")
                else:
                    self.adicionar_status("Diagnostico concluido: Problemas encontrados")
                    messagebox.showwarning("Diagnostico", "Problemas encontrados. Verifique o log de status.")
                
            except Exception as e:
                error_msg = f"Erro no diagnostico: {str(e)}"
                self.adicionar_status(error_msg)
                messagebox.showerror("Erro", error_msg)
            
            finally:
                self.processando = False
                self._controlar_botoes('normal')
        
        threading.Thread(target=diagnosticar, daemon=True).start()
    
    def executar(self):
        """Inicia a aplicacao."""
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
            self.root.mainloop()
        except KeyboardInterrupt:
            self._on_closing()
    
    def _on_closing(self):
        """Gerencia fechamento da aplicacao."""
        if self.processando:
            if messagebox.askokcancel("Fechar", "Existe um processo em andamento. Deseja realmente fechar?"):
                self.root.quit()
        else:
            self.root.quit()


def main():
    """Funcao principal da aplicacao."""
    try:
        # Verificar se estamos no diretorio correto
        if not os.path.exists("documentos"):
            messagebox.showerror(
                "Diretorio Incorreto", 
                "Por favor, execute este programa na pasta raiz do projeto MMZR\n"
                "O diretorio deve conter a pasta 'documentos'"
            )
            return False
        
        # Inicializar e executar aplicacao
        logger.info("Iniciando aplicacao MMZR v1.0")
        app = MMZRInterface()
        app.executar()
        logger.info("Aplicacao encerrada")
        return True
        
    except Exception as e:
        error_msg = f"Erro fatal ao iniciar aplicacao: {str(e)}"
        logger.error(error_msg)
        messagebox.showerror("Erro Fatal", error_msg)
        return False


if __name__ == "__main__":
    main() 