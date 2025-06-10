import os
import platform

class MMZRCompatibilidade:
    """Classe para gerenciar compatibilidade entre sistemas operacionais"""
    
    @staticmethod
    def get_os_info():
        """Retorna informações básicas do sistema operacional"""
        return {
            'sistema': platform.system(),
            'versao': platform.version()
        }
    
    @staticmethod
    def get_planilhas_path():
        """Retorna os caminhos das planilhas"""
        base_path = os.path.join("documentos", "dados", "Planilha Inteli.xlsm")
        rentabilidade_path = os.path.join("documentos", "dados", "Planilha Inteli - dados de rentabilidade.xlsx")
        
        return base_path, rentabilidade_path
    
    @staticmethod
    def enviar_email(destinatario, assunto, caminho_html):
        """Envia email ou simula o envio dependendo do sistema"""
        os_info = MMZRCompatibilidade.get_os_info()
        
        if os_info['sistema'] == 'Windows':
            try:
                import win32com.client as win32
                
                # Abrir o Outlook
                outlook = win32.Dispatch('outlook.application')
                
                # Criar um novo email
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                
                # Configurar o email
                mail.To = destinatario
                mail.Subject = assunto
                
                # Ler o conteúdo HTML
                with open(caminho_html, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                mail.HTMLBody = html_content
                
                # Enviar o email
                mail.Send()
                
                return True
                
            except Exception as e:
                print(f"Erro ao enviar email: {e}")
                return False
        else:
            # Para sistemas não-Windows, simular o envio
            print(f"[SIMULAÇÃO] Email enviado para: {destinatario}")
            print(f"[SIMULAÇÃO] Assunto: {assunto}")
            print(f"[SIMULAÇÃO] HTML: {caminho_html}")
            return True
    
    @staticmethod
    def testar_compatibilidade():
        """Testa a compatibilidade do sistema"""
        os_info = MMZRCompatibilidade.get_os_info()
        base_path, rentabilidade_path = MMZRCompatibilidade.get_planilhas_path()
        
        print("=== TESTE DE COMPATIBILIDADE ===")
        print(f"Sistema operacional: {os_info['sistema']}")
        
        # Verificar se as planilhas existem
        paths_ok = os.path.exists(base_path) and os.path.exists(rentabilidade_path)
        
        # Verificar capacidade de email
        if os_info['sistema'] == 'Windows':
            try:
                import win32com.client
                email_ok = True
                print("✓ Sistema de email do Outlook disponível")
            except ImportError:
                email_ok = False
                print("⚠ win32com não disponível, emails serão simulados")
        else:
            email_ok = True
            print("✓ Sistema de email simulado disponível para desenvolvimento")
        
        print("=== TESTE CONCLUÍDO ===")
        
        return {
            'os_compatible': True,
            'paths_ok': paths_ok,
            'email_ok': email_ok
        }


# Executar o teste quando o arquivo for executado diretamente
if __name__ == "__main__":
    MMZRCompatibilidade.testar_compatibilidade() 