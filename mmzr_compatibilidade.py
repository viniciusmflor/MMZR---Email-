import os
import platform
import pandas as pd
from datetime import datetime

class MMZRCompatibilidade:
    """Classe para garantir a compatibilidade entre macOS e Windows"""
    
    @staticmethod
    def get_os_info():
        """Retorna informações sobre o sistema operacional"""
        return {
            "sistema": platform.system(),
            "versao": platform.version(),
            "arquitetura": platform.architecture(),
            "python": platform.python_version(),
        }
    
    @staticmethod
    def get_path(*args):
        """Retorna um caminho compatível com o sistema operacional atual"""
        return os.path.join(*args)
    
    @staticmethod
    def get_abs_path(*args):
        """Retorna um caminho absoluto compatível com o sistema operacional atual"""
        return os.path.abspath(os.path.join(*args))
    
    @staticmethod
    def get_planilhas_path():
        """Retorna os caminhos das planilhas base e de rentabilidade"""
        base_dir = "documentos"
        dados_dir = "dados"
        
        # Verificar se o diretório documentos/dados existe
        if not os.path.exists(os.path.join(base_dir, dados_dir)):
            # Tentar encontrar o diretório correto baseado no diretório atual
            cwd = os.getcwd()
            if os.path.basename(cwd) == "MMZR - Email":
                base_dir = os.path.join(cwd, "documentos")
                dados_dir = "dados"
            
        planilha_base = os.path.join(base_dir, dados_dir, "Planilha Inteli.xlsm")
        planilha_rentabilidade = os.path.join(base_dir, dados_dir, "Planilha Inteli - dados de rentabilidade.xlsx")
        
        return planilha_base, planilha_rentabilidade
    
    @staticmethod
    def enviar_email(destinatario, assunto, caminho_html, anexos=None):
        """
        Envia um email usando o Outlook (Windows) ou exibe uma mensagem (macOS)
        
        Args:
            destinatario: Email do destinatário
            assunto: Assunto do email
            caminho_html: Caminho para o arquivo HTML do relatório
            anexos: Lista de caminhos para arquivos a serem anexados (opcional)
        
        Returns:
            bool: True se o email foi enviado, False caso contrário
        """
        try:
            # Ler o conteúdo HTML
            with open(caminho_html, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # Verificar o sistema operacional
            if platform.system() == "Windows":
                try:
                    import win32com.client
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    mail = outlook.CreateItem(0)  # 0 = olMailItem
                    mail.To = destinatario
                    mail.Subject = assunto
                    mail.HTMLBody = html_content
                    
                    # Adicionar anexos, se houver
                    if anexos:
                        for anexo in anexos:
                            if os.path.exists(anexo):
                                mail.Attachments.Add(anexo)
                    
                    mail.Save()
                    print(f"Email salvo como rascunho para {destinatario}")
                    
                    #mail.Send()
                    #print(f"Email enviado para {destinatario} via Outlook")
                    #return True
                except win32com.client.pywintypes.com_error as e:
                    print(f"ERRO COM: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    return False
                except ImportError:
                    print("ERRO: win32com não está instalado. A integração com Outlook não funcionará")
                    print(f"Email seria enviado para {destinatario}")
                    return False
                except Exception as e:
                    print(f"ERRO GERAL: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    return False
            else:
                # No macOS, apenas exibir uma mensagem
                print(f"[SIMULAÇÃO] Email enviado para {destinatario}")
                print(f"  Assunto: {assunto}")
                print(f"  Arquivo HTML: {caminho_html}")
                return True
                
        except Exception as e:
            print(f"ERRO ao abrir arquivo: {str(e)}")
            return False
    
    @staticmethod
    def testar_compatibilidade():
        """Testa a compatibilidade entre Mac e Windows"""
        info = MMZRCompatibilidade.get_os_info()
        print("\n=== TESTE DE COMPATIBILIDADE ===")
        print(f"Sistema operacional: {info['sistema']}")
        print(f"Versão: {info['versao']}")
        print(f"Python: {info['python']}")
        
        # Testar caminhos
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        print("\nVerificando caminhos de arquivos:")
        print(f"1. Planilha base: {planilha_base}")
        print(f"   Existe: {os.path.exists(planilha_base)}")
        
        print(f"2. Planilha rentabilidade: {planilha_rentabilidade}")
        print(f"   Existe: {os.path.exists(planilha_rentabilidade)}")
        
        # Testar disponibilidade do win32com
        if info['sistema'] == "Windows":
            print("\nTestando integração com Outlook:")
            try:
                import win32com.client
                print("✓ win32com está disponível para integração com Outlook")
            except ImportError:
                print("✗ ERRO: win32com não está instalado. A integração com Outlook não funcionará")
        else:
            print("\nSistema não é Windows, win32com não será utilizado")
            print("✓ Sistema de email simulado disponível para desenvolvimento")
        
        print("\n=== TESTE CONCLUÍDO ===")
        
        return {
            "sistema": info['sistema'],
            "paths_ok": os.path.exists(planilha_base) and os.path.exists(planilha_rentabilidade),
            "outlook_ok": info['sistema'] != "Windows" or MMZRCompatibilidade._check_win32com()
        }
    
    @staticmethod
    def _check_win32com():
        """Verifica se win32com está disponível (apenas para Windows)"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False


# Executar o teste quando o arquivo for executado diretamente
if __name__ == "__main__":
    MMZRCompatibilidade.testar_compatibilidade() 