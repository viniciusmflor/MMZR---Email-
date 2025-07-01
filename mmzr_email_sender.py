"""
MMZR Family Office - Modulo de Integracao com Microsoft Outlook
Sistema de envio automatizado de emails via Outlook

Versao: 1.0.0
Plataforma: Windows (Microsoft Outlook)
"""

import os
import logging
from typing import Optional, List

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def enviar_email_outlook(
    para: str, 
    assunto: str, 
    corpo_html: str, 
    anexos: Optional[List[str]] = None,
    copia: Optional[str] = None,
    copia_oculta: Optional[str] = None
) -> bool:
    """
    Prepara um email no Microsoft Outlook (Windows).
    
    Args:
        para: Email do destinatario
        assunto: Assunto do email
        corpo_html: Conteudo HTML do email
        anexos: Lista de caminhos para arquivos de anexo (opcional)
        copia: Email para copia (opcional)
        copia_oculta: Email para copia oculta (opcional)
        
    Returns:
        bool: True se email foi preparado com sucesso, False caso contrario
    """
    try:
        logger.info(f"Preparando email para: {para}")
        return _enviar_email_outlook_windows(para, assunto, corpo_html, anexos, copia, copia_oculta)
        
    except Exception as e:
        error_msg = f"Erro ao preparar email para {para}: {str(e)}"
        logger.error(error_msg)
        print(f"ERRO: {error_msg}")
        return False


def _enviar_email_outlook_windows(para, assunto, corpo_html, anexos, copia, copia_oculta):
    """Envia email usando Microsoft Outlook no Windows."""
    try:
        import win32com.client
        
        logger.info("Conectando ao Microsoft Outlook...")
        
        # Conectar ao Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Criar novo email
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Configurar destinatarios
        mail.To = para
        if copia:
            mail.CC = copia
        if copia_oculta:
            mail.BCC = copia_oculta
            
        # Configurar conteudo
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        
        # Adicionar anexos se fornecidos
        if anexos:
            for anexo in anexos:
                if os.path.exists(anexo):
                    mail.Attachments.Add(anexo)
                    logger.info(f"Anexo adicionado: {anexo}")
                else:
                    logger.warning(f"Anexo nao encontrado: {anexo}")
        
        # Exibir email (nao enviar automaticamente)
        mail.Display(False)
        
        logger.info(f"Email preparado com sucesso no Outlook para {para}")
        print(f"Email preparado no Microsoft Outlook para: {para}")
        print(f"  - Assunto: {assunto}")
        print(f"  - Destinatario: {para}")
        if copia:
            print(f"  - Copia: {copia}")
        if anexos:
            print(f"  - Anexos: {len(anexos)} arquivo(s)")
        
        return True
        
    except ImportError:
        error_msg = "Modulo pywin32 nao instalado. Execute: pip install pywin32"
        logger.error(error_msg)
        print(f"ERRO: {error_msg}")
        print("   Para instalar: pip install pywin32")
        return False
        
    except Exception as e:
        error_msg = f"Erro no Microsoft Outlook: {str(e)}"
        logger.error(error_msg)
        print(f"ERRO: {error_msg}")
        print("\nPossiveis solucoes:")
        print("1. Verificar se o Microsoft Outlook esta instalado")
        print("2. Verificar se o Outlook esta configurado com uma conta")
        print("3. Executar como administrador se necessario")
        print("4. Fechar e reabrir o Outlook")
        return False





def enviar_emails_multiplos(emails_data: List[dict]) -> dict:
    """
    Prepara multiplos emails no Microsoft Outlook.
    
    Args:
        emails_data: Lista de dicionarios com dados dos emails
                    Cada dicionario deve conter: 'para', 'assunto', 'corpo_html'
                    
    Returns:
        dict: Estatisticas do envio (sucessos, falhas, total)
    """
    stats = {'sucessos': 0, 'falhas': 0, 'total': len(emails_data)}
    
    logger.info(f"Iniciando preparacao de {stats['total']} emails")
    print(f"\nPreparando {stats['total']} emails no Outlook...")
    print("=" * 60)
    
    for i, email_data in enumerate(emails_data, 1):
        print(f"\nProcessando email {i}/{stats['total']}")
        
        sucesso = enviar_email_outlook(
            para=email_data['para'],
            assunto=email_data['assunto'],
            corpo_html=email_data['corpo_html'],
            anexos=email_data.get('anexos'),
            copia=email_data.get('copia'),
            copia_oculta=email_data.get('copia_oculta')
        )
        
        if sucesso:
            stats['sucessos'] += 1
        else:
            stats['falhas'] += 1
    
    print("\n" + "=" * 60)
    print(f"Preparacao concluida:")
    print(f"  Sucessos: {stats['sucessos']}")
    print(f"  Falhas: {stats['falhas']}")
    print(f"  Total: {stats['total']}")
    
    logger.info(f"Preparacao concluida - Sucessos: {stats['sucessos']}, Falhas: {stats['falhas']}")
    
    return stats


def verificar_outlook_disponivel() -> bool:
    """
    Verifica se o Microsoft Outlook esta disponivel.
    
    Returns:
        bool: True se Outlook esta disponivel, False caso contrario
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        logger.info("Microsoft Outlook disponivel e funcionando")
        return True
    except ImportError:
        logger.warning("pywin32 nao instalado - instale com: pip install pywin32")
        return False
    except Exception as e:
        logger.warning(f"Microsoft Outlook nao esta funcionando: {str(e)}")
        return False


if __name__ == "__main__":
    # Teste do modulo
    print("MMZR Email Sender - Teste de Funcionalidade")
    print("=" * 50)
    print("Plataforma: Windows com Microsoft Outlook")
    
    if verificar_outlook_disponivel():
        print("Microsoft Outlook esta disponivel")
        
        # Teste com email ficticio
        teste_sucesso = enviar_email_outlook(
            para="teste@exemplo.com",
            assunto="Teste MMZR - Email Sender",
            corpo_html="<h1>Teste</h1><p>Este e um teste do sistema de email MMZR.</p>"
        )
        
        if teste_sucesso:
            print("Teste realizado com sucesso")
        else:
            print("Falha no teste")
    else:
        print("Microsoft Outlook nao esta disponivel")
        print("  Execute: pip install pywin32") 