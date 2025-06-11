"""
MMZR Family Office - Módulo de Compatibilidade

Este módulo garante a compatibilidade entre diferentes sistemas operacionais
(macOS e Windows) para o sistema de geração de relatórios MMZR.

Autor: MMZR Family Office
Versão: 2.0.0
Data: 2025-01-11
"""

import os
import platform
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Union, Any
import logging
import json

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class MMZRCompatibilidade:
    """
    Classe para garantir a compatibilidade entre macOS e Windows.
    
    Esta classe fornece métodos estáticos para operações que dependem do
    sistema operacional, como caminhos de arquivos e envio de emails.
    """
    
    @staticmethod
    def get_os_info() -> Dict[str, str]:
        """
        Retorna informações sobre o sistema operacional.
        
        Returns:
            Dict[str, str]: Dicionário com informações do sistema operacional
        """
        try:
            return {
                "sistema": platform.system(),
                "versao": platform.version(),
                "arquitetura": str(platform.architecture()),
                "python": platform.python_version(),
            }
        except Exception as e:
            logger.error(f"Erro ao obter informações do sistema: {e}")
            return {"sistema": "Desconhecido", "versao": "", "arquitetura": "", "python": ""}
    
    @staticmethod
    def get_path(*args: str) -> str:
        """
        Retorna um caminho compatível com o sistema operacional atual.
        
        Args:
            *args (str): Componentes do caminho
            
        Returns:
            str: Caminho formatado para o sistema atual
        """
        return os.path.join(*args)
    
    @staticmethod
    def get_abs_path(*args: str) -> str:
        """
        Retorna um caminho absoluto compatível com o sistema operacional atual.
        
        Args:
            *args (str): Componentes do caminho
            
        Returns:
            str: Caminho absoluto formatado para o sistema atual
        """
        return os.path.abspath(os.path.join(*args))
    
    @staticmethod
    def get_planilhas_path() -> Tuple[str, str]:
        """
        Obtém os caminhos das planilhas Excel necessárias.
        Pode usar configuração personalizada ou detecção automática.
        
        Returns:
            Tuple[str, str]: Tupla com (caminho_planilha_base, caminho_planilha_rentabilidade)
        """
        try:
            base_dir = "documentos"
            dados_dir = "dados"
            
            # Verificar se o diretório documentos/dados existe
            dados_path = os.path.join(base_dir, dados_dir)
            if not os.path.exists(dados_path):
                # Tentar encontrar o diretório correto baseado no diretório atual
                cwd = os.getcwd()
                if os.path.basename(cwd) == "MMZR - Email":
                    dados_path = os.path.join(cwd, base_dir, dados_dir)
            
            # Tentar carregar configuração personalizada
            config_path = "config_planilhas.json"
            config = MMZRCompatibilidade._load_config(config_path)
            
            if config and not config.get("auto_detectar", True):
                # Usar nomes especificados na configuração
                planilha_base = config["planilhas"]["planilha_base"]
                planilha_rentabilidade = config["planilhas"]["planilha_rentabilidade"]
                
                if planilha_base and planilha_rentabilidade:
                    planilha_base = os.path.join(dados_path, planilha_base)
                    planilha_rentabilidade = os.path.join(dados_path, planilha_rentabilidade)
                    
                    logger.info("Usando configuração personalizada de planilhas")
                    logger.info(f"Base: {planilha_base}")
                    logger.info(f"Rentabilidade: {planilha_rentabilidade}")
                    
                    return planilha_base, planilha_rentabilidade
            
            # Detecção automática
            logger.info("Detectando planilhas automaticamente...")
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade._detectar_planilhas(dados_path)
            
            if planilha_base and planilha_rentabilidade:
                logger.info(f"Planilhas detectadas: Base={os.path.basename(planilha_base)}, Rentabilidade={os.path.basename(planilha_rentabilidade)}")
                return planilha_base, planilha_rentabilidade
            
            # Fallback para nomes padrão (compatibilidade com versão anterior)
            logger.warning("Usando nomes de planilhas padrão como fallback")
            planilha_base = os.path.join(dados_path, "Planilha Inteli.xlsm")
            planilha_rentabilidade = os.path.join(dados_path, "Planilha Inteli - dados de rentabilidade.xlsx")
            
            return planilha_base, planilha_rentabilidade
            
        except Exception as e:
            logger.error(f"Erro ao configurar caminhos das planilhas: {e}")
            return "", ""
    
    @staticmethod
    def _load_config(config_path: str) -> Optional[Dict[str, Any]]:
        """
        Carrega configuração de planilhas de um arquivo JSON.
        
        Args:
            config_path (str): Caminho para o arquivo de configuração
            
        Returns:
            Optional[Dict[str, Any]]: Configuração carregada ou None se erro
        """
        try:
            if not os.path.exists(config_path):
                return None
                
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
                
        except Exception as e:
            logger.warning(f"Erro ao carregar configuração: {e}")
            return None
    
    @staticmethod
    def _detectar_planilhas(dados_path: str) -> Tuple[str, str]:
        """
        Detecta automaticamente planilhas Excel na pasta de dados.
        
        Args:
            dados_path (str): Caminho para a pasta de dados
            
        Returns:
            Tuple[str, str]: Tupla com caminhos das planilhas detectadas
        """
        try:
            if not os.path.exists(dados_path):
                logger.error(f"Pasta de dados não encontrada: {dados_path}")
                return "", ""
            
            # Listar arquivos Excel na pasta
            excel_files = []
            for file in os.listdir(dados_path):
                if file.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                    excel_files.append(os.path.join(dados_path, file))
            
            if len(excel_files) == 0:
                logger.error("Nenhum arquivo Excel encontrado na pasta de dados")
                return "", ""
            
            if len(excel_files) == 1:
                logger.warning("Apenas um arquivo Excel encontrado. Usando o mesmo para ambas as funções.")
                return excel_files[0], excel_files[0]
            
            # Identificar planilhas por estratégia melhorada
            planilha_base = ""
            planilha_rentabilidade = ""
            
            # Estratégia 1: Identificar por tipo de arquivo e palavras-chave
            xlsm_files = []
            xlsx_files = []
            
            for file_path in excel_files:
                filename = os.path.basename(file_path).lower()
                
                if filename.endswith('.xlsm'):
                    xlsm_files.append(file_path)
                else:
                    xlsx_files.append(file_path)
            
            # Estratégia 2: Priorizar .xlsm para planilha base (geralmente tem macros)
            for file_path in xlsm_files:
                filename = os.path.basename(file_path).lower()
                # Verificar se NÃO é planilha de rentabilidade
                if not any(keyword in filename for keyword in ['rentabilidade', 'dados de rentabilidade']):
                    planilha_base = file_path
                    break
            
            # Estratégia 3: Identificar planilha de rentabilidade por palavras-chave específicas
            for file_path in excel_files:
                filename = os.path.basename(file_path).lower()
                if any(keyword in filename for keyword in ['rentabilidade', 'dados de rentabilidade', 'performance']):
                    planilha_rentabilidade = file_path
                    break
            
            # Estratégia 4: Se ainda não identificou planilha base, usar por exclusão
            if not planilha_base:
                for file_path in excel_files:
                    if file_path != planilha_rentabilidade:
                        filename = os.path.basename(file_path).lower()
                        # Verificar se tem características de planilha base
                        if any(keyword in filename for keyword in ['base', 'cliente', 'inteli']) or filename.endswith('.xlsm'):
                            planilha_base = file_path
                            break
            
            # Estratégia 5: Fallback - usar os primeiros arquivos encontrados
            if not planilha_base and len(excel_files) >= 1:
                # Excluir a planilha de rentabilidade já identificada
                for file_path in excel_files:
                    if file_path != planilha_rentabilidade:
                        planilha_base = file_path
                        break
                
                # Se ainda não tem base, usar o primeiro arquivo
                if not planilha_base:
                    planilha_base = excel_files[0]
            
            if not planilha_rentabilidade and len(excel_files) >= 2:
                # Usar arquivo diferente da planilha base
                for file_path in excel_files:
                    if file_path != planilha_base:
                        planilha_rentabilidade = file_path
                        break
            
            # Validar se as planilhas têm as abas necessárias
            if planilha_base:
                if MMZRCompatibilidade._validar_abas(planilha_base, ["Base Clientes"]):
                    logger.info(f"✓ Planilha base validada: {os.path.basename(planilha_base)}")
                else:
                    logger.warning(f"⚠ Planilha base pode não ter a aba 'Base Clientes': {os.path.basename(planilha_base)}")
                    # Se a planilha base não tem a aba correta, tentar trocar
                    if planilha_rentabilidade and MMZRCompatibilidade._validar_abas(planilha_rentabilidade, ["Base Clientes"]):
                        logger.info("🔄 Trocando planilhas: rentabilidade tinha a aba 'Base Clientes'")
                        planilha_base, planilha_rentabilidade = planilha_rentabilidade, planilha_base
            
            if planilha_rentabilidade:
                logger.info(f"✓ Planilha rentabilidade: {os.path.basename(planilha_rentabilidade)}")
            
            return planilha_base, planilha_rentabilidade
            
        except Exception as e:
            logger.error(f"Erro na detecção automática de planilhas: {e}")
            return "", ""
    
    @staticmethod
    def _validar_abas(file_path: str, abas_necessarias: List[str]) -> bool:
        """
        Valida se uma planilha Excel tem as abas necessárias.
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            abas_necessarias (List[str]): Lista de nomes de abas necessárias
            
        Returns:
            bool: True se todas as abas necessárias existem
        """
        try:
            import pandas as pd
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            for aba in abas_necessarias:
                if aba not in sheet_names:
                    return False
            return True
            
        except Exception:
            return False
    
    @staticmethod
    def enviar_email(destinatario: str, assunto: str, caminho_html: str, anexos: Optional[List[str]] = None) -> bool:
        """
        Envia um email usando o Outlook (Windows) ou exibe uma mensagem (macOS).
        
        Args:
            destinatario (str): Email do destinatário
            assunto (str): Assunto do email
            caminho_html (str): Caminho para o arquivo HTML do relatório
            anexos (Optional[List[str]]): Lista de caminhos para arquivos a serem anexados
        
        Returns:
            bool: True se o email foi enviado, False caso contrário
        """
        try:
            # Validar parâmetros de entrada
            if not destinatario or not assunto or not caminho_html:
                logger.error("Parâmetros obrigatórios não fornecidos para envio de email")
                return False
                
            if not os.path.exists(caminho_html):
                logger.error(f"Arquivo HTML não encontrado: {caminho_html}")
                return False
            
            # Ler o conteúdo HTML
            with open(caminho_html, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # Verificar o sistema operacional
            sistema = platform.system()
            logger.info(f"Enviando email no sistema: {sistema}")
            
            if sistema == "Windows":
                return MMZRCompatibilidade._enviar_email_windows(
                    destinatario, assunto, html_content, anexos
                )
            else:
                return MMZRCompatibilidade._simular_envio_email(
                    destinatario, assunto, caminho_html
                )
                
        except Exception as e:
            logger.error(f"Erro ao enviar email: {e}")
            return False
    
    @staticmethod
    def _enviar_email_windows(destinatario: str, assunto: str, html_content: str, anexos: Optional[List[str]]) -> bool:
        """
        Envia email usando Outlook no Windows.
        
        Args:
            destinatario (str): Email do destinatário
            assunto (str): Assunto do email
            html_content (str): Conteúdo HTML do email
            anexos (Optional[List[str]]): Lista de anexos
            
        Returns:
            bool: True se enviado com sucesso
        """
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
                        logger.info(f"Anexo adicionado: {anexo}")
            
            mail.Send()
            logger.info(f"Email enviado para {destinatario} via Outlook")
            return True
            
        except ImportError:
            logger.error("win32com não está instalado. A integração com Outlook não funcionará")
            logger.info(f"Email seria enviado para {destinatario}")
            return False
        except Exception as e:
            logger.error(f"Erro ao enviar email via Outlook: {e}")
            return False
    
    @staticmethod
    def _simular_envio_email(destinatario: str, assunto: str, caminho_html: str) -> bool:
        """
        Simula o envio de email em sistemas não-Windows.
        
        Args:
            destinatario (str): Email do destinatário
            assunto (str): Assunto do email
            caminho_html (str): Caminho do arquivo HTML
            
        Returns:
            bool: Sempre True (simulação)
        """
        logger.info(f"[SIMULAÇÃO] Email enviado para {destinatario}")
        logger.info(f"  Assunto: {assunto}")
        logger.info(f"  Arquivo HTML: {caminho_html}")
        return True
    
    @staticmethod
    def testar_compatibilidade() -> Dict[str, Union[str, bool]]:
        """
        Testa a compatibilidade entre Mac e Windows.
        
        Returns:
            Dict[str, Union[str, bool]]: Resultado dos testes de compatibilidade
        """
        try:
            info = MMZRCompatibilidade.get_os_info()
            logger.info("\n=== TESTE DE COMPATIBILIDADE ===")
            logger.info(f"Sistema operacional: {info['sistema']}")
            logger.info(f"Versão: {info['versao']}")
            logger.info(f"Python: {info['python']}")
            
            # Testar caminhos
            planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
            
            logger.info("\nVerificando caminhos de arquivos:")
            base_exists = os.path.exists(planilha_base)
            rent_exists = os.path.exists(planilha_rentabilidade)
            
            logger.info(f"1. Planilha base: {planilha_base}")
            logger.info(f"   Existe: {base_exists}")
            
            logger.info(f"2. Planilha rentabilidade: {planilha_rentabilidade}")
            logger.info(f"   Existe: {rent_exists}")
            
            # Testar disponibilidade do win32com
            outlook_ok = True
            if info['sistema'] == "Windows":
                logger.info("\nTestando integração com Outlook:")
                outlook_ok = MMZRCompatibilidade._check_win32com()
                if outlook_ok:
                    logger.info("✓ win32com está disponível para integração com Outlook")
                else:
                    logger.error("✗ win32com não está instalado. A integração com Outlook não funcionará")
            else:
                logger.info("\nSistema não é Windows, win32com não será utilizado")
                logger.info("✓ Sistema de email simulado disponível para desenvolvimento")
            
            logger.info("\n=== TESTE CONCLUÍDO ===")
            
            return {
                "sistema": info['sistema'],
                "paths_ok": base_exists and rent_exists,
                "outlook_ok": outlook_ok
            }
            
        except Exception as e:
            logger.error(f"Erro durante teste de compatibilidade: {e}")
            return {"sistema": "Erro", "paths_ok": False, "outlook_ok": False}
    
    @staticmethod
    def _check_win32com() -> bool:
        """
        Verifica se win32com está disponível (apenas para Windows).
        
        Returns:
            bool: True se win32com estiver disponível
        """
        try:
            import win32com.client
            return True
        except ImportError:
            return False


# Executar o teste quando o arquivo for executado diretamente
if __name__ == "__main__":
    MMZRCompatibilidade.testar_compatibilidade() 