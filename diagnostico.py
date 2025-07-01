"""
MMZR Family Office - Sistema de Compatibilidade e Diagnóstico
Detecção automática de planilhas e verificação do sistema

Autor: MMZR Family Office
Versão: 1.0.0
"""

import os
import pandas as pd
import logging
from typing import Tuple, Dict, Any, Optional

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class MMZRCompatibilidade:
    """Gerencia detecção automática de planilhas e compatibilidade do sistema."""
    
    @classmethod
    def get_planilhas_path(cls) -> Tuple[str, str]:
        """
        Detecta automaticamente as planilhas base e rentabilidade.
        
        Returns:
            Tuple[str, str]: Tupla com caminhos (planilha_base, planilha_rentabilidade)
            
        Raises:
            FileNotFoundError: Se pasta ou arquivos não forem encontrados
        """
        dados_dir = "documentos/dados"
        
        if not os.path.exists(dados_dir):
            raise FileNotFoundError(f"Pasta não encontrada: {dados_dir}")
        
        # Listar arquivos Excel
        excel_files = []
        for arquivo in os.listdir(dados_dir):
            if arquivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                caminho_completo = os.path.join(dados_dir, arquivo)
                excel_files.append(caminho_completo)
        
        if len(excel_files) == 0:
            raise FileNotFoundError("Nenhum arquivo Excel encontrado em documentos/dados/")
        
        if len(excel_files) == 1:
            # Se só há um arquivo, usar para ambos
            logger.info(f"Apenas um arquivo Excel encontrado, usando para ambas as funções: {excel_files[0]}")
            return excel_files[0], excel_files[0]
        
        # Detectar qual é qual baseado no conteúdo
        planilha_base = None
        planilha_rentabilidade = None
        
        for arquivo in excel_files:
            try:
                excel_file = pd.ExcelFile(arquivo)
                
                # Se tem aba "Base Clientes", é a planilha base
                if "Base Clientes" in excel_file.sheet_names:
                    planilha_base = arquivo
                    logger.info(f"Planilha base identificada: {os.path.basename(arquivo)}")
                else:
                    # Senão, provavelmente é a de rentabilidade
                    planilha_rentabilidade = arquivo
                    logger.info(f"Planilha de rentabilidade identificada: {os.path.basename(arquivo)}")
                    
            except Exception as e:
                logger.warning(f"Erro ao verificar {arquivo}: {e}")
                continue
        
        # Se não identificou a base, usar o primeiro arquivo
        if not planilha_base:
            planilha_base = excel_files[0]
            logger.warning("Planilha base não identificada automaticamente, usando primeiro arquivo")
        
        # Se não identificou a rentabilidade, usar um arquivo diferente da base
        if not planilha_rentabilidade:
            for arquivo in excel_files:
                if arquivo != planilha_base:
                    planilha_rentabilidade = arquivo
                    break
            
            # Se ainda não tem, usar o mesmo arquivo
            if not planilha_rentabilidade:
                planilha_rentabilidade = planilha_base
                logger.warning("Usando mesmo arquivo para base e rentabilidade")
        
        logger.info(f"Planilhas detectadas - Base: {os.path.basename(planilha_base)}, "
                   f"Rentabilidade: {os.path.basename(planilha_rentabilidade)}")
        
        return planilha_base, planilha_rentabilidade
    
    @classmethod
    def testar_compatibilidade(cls) -> bool:
        """
        Testa compatibilidade do sistema.
        
        Returns:
            bool: True se sistema está funcionando corretamente
        """
        try:
            logger.info("Iniciando teste de compatibilidade")
            
            # Testar acesso às planilhas
            planilha_base, planilha_rentabilidade = cls.get_planilhas_path()
            
            # Testar leitura das planilhas
            excel_base = pd.ExcelFile(planilha_base)
            excel_rent = pd.ExcelFile(planilha_rentabilidade)
            
            # Verificar abas essenciais
            if "Base Clientes" not in excel_base.sheet_names:
                logger.error("Aba 'Base Clientes' não encontrada na planilha base")
                print("ERRO: Aba 'Base Clientes' não encontrada na planilha base")
                return False
            
            # Testar leitura básica
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes", nrows=5)
            df_rent = pd.read_excel(excel_rent, sheet_name=excel_rent.sheet_names[0], nrows=5)
            
            # Verificar colunas essenciais
            colunas_necessarias_clientes = ['Nome cliente', 'Código carteira smart']
            for coluna in colunas_necessarias_clientes:
                if coluna not in df_clientes.columns:
                    logger.error(f"Coluna '{coluna}' não encontrada na aba 'Base Clientes'")
                    print(f"ERRO: Coluna '{coluna}' não encontrada na aba 'Base Clientes'")
                    return False
            
            logger.info("Teste de compatibilidade concluído com sucesso")
            print("SUCESSO: Sistema compatível e funcional")
            return True
            
        except Exception as e:
            error_msg = f"Erro de compatibilidade: {e}"
            logger.error(error_msg)
            print(f"ERRO: {error_msg}")
            return False
    
    @classmethod
    def verificar_estrutura_dados(cls) -> Dict[str, Any]:
        """
        Verifica a estrutura dos dados nas planilhas.
        
        Returns:
            Dict[str, Any]: Dicionário com informações da estrutura ou erro
        """
        try:
            logger.info("Verificando estrutura dos dados")
            planilha_base, planilha_rentabilidade = cls.get_planilhas_path()
            
            # Verificar planilha base
            excel_base = pd.ExcelFile(planilha_base)
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
            
            # Verificar planilha de rentabilidade
            excel_rent = pd.ExcelFile(planilha_rentabilidade)
            df_rent = pd.read_excel(excel_rent, sheet_name=excel_rent.sheet_names[0])
            
            # Verificar integridade dos dados
            clientes_validos = df_clientes[df_clientes['Nome cliente'].notna() & 
                                         (df_clientes['Nome cliente'] != '') &
                                         (df_clientes['Nome cliente'] != 'Nome Cliente')]
            
            # Contar clientes únicos (sem duplicatas por múltiplas carteiras)
            clientes_unicos = clientes_validos['Nome cliente'].unique()
            
            # Contar registros válidos de rentabilidade (sem valores nulos)
            rent_validos = df_rent[df_rent['Nome cliente'].notna() & 
                                  (df_rent['Nome cliente'] != '') &
                                  (df_rent['Nome cliente'].astype(str) != 'nan')]
            
            resultado = {
                "status": "sucesso",
                "planilha_base": {
                    "arquivo": os.path.basename(planilha_base),
                    "abas": excel_base.sheet_names,
                    "clientes_total": len(df_clientes),
                    "clientes_validos": len(clientes_unicos),  # Agora conta clientes únicos
                    "carteiras_total": len(clientes_validos),  # Total de carteiras
                    "colunas": list(df_clientes.columns)
                },
                "planilha_rentabilidade": {
                    "arquivo": os.path.basename(planilha_rentabilidade),
                    "abas": excel_rent.sheet_names,
                    "registros_total": len(rent_validos),  # Agora conta apenas válidos
                    "registros_brutos": len(df_rent),  # Total incluindo vazios
                    "colunas": list(df_rent.columns)
                }
            }
            
            logger.info("Verificação de estrutura concluída com sucesso")
            return resultado
            
        except Exception as e:
            error_msg = f"Erro na verificação de estrutura: {str(e)}"
            logger.error(error_msg)
            return {
                "status": "erro",
                "mensagem": error_msg
            }
    
    @classmethod
    def executar_verificacao_completa(cls) -> bool:
        """
        Executa verificação completa do sistema.
        
        Returns:
            bool: True se tudo está funcionando corretamente
        """
        try:
            logger.info("Iniciando verificação completa do sistema")
            
            print("VERIFICAÇÃO COMPLETA DO SISTEMA MMZR")
            print("=" * 50)
            
            # 1. Verificar estrutura de arquivos
            print("\n1. Verificando estrutura de arquivos...")
            if not os.path.exists("documentos"):
                print("ERRO: Pasta 'documentos' não encontrada")
                return False
            
            if not os.path.exists("documentos/dados"):
                print("ERRO: Pasta 'documentos/dados' não encontrada")
                return False
            
            print("   Estrutura de pastas: OK")
            
            # 2. Detectar planilhas
            print("\n2. Detectando planilhas...")
            try:
                planilha_base, planilha_rentabilidade = cls.get_planilhas_path()
                print(f"   Base: {os.path.basename(planilha_base)}")
                print(f"   Rentabilidade: {os.path.basename(planilha_rentabilidade)}")
            except Exception as e:
                print(f"ERRO na detecção: {e}")
                return False
            
            # 3. Testar compatibilidade
            print("\n3. Testando compatibilidade...")
            if not cls.testar_compatibilidade():
                return False
            
            # 4. Verificar estrutura de dados
            print("\n4. Verificando estrutura de dados...")
            estrutura = cls.verificar_estrutura_dados()
            
            if estrutura["status"] == "erro":
                print(f"ERRO na estrutura: {estrutura['mensagem']}")
                return False
            
            print(f"   Clientes únicos: {estrutura['planilha_base']['clientes_validos']}")
            print(f"   Total de carteiras: {estrutura['planilha_base']['carteiras_total']}")
            print(f"   Registros de rentabilidade válidos: {estrutura['planilha_rentabilidade']['registros_total']}")
            print(f"   Registros brutos (incluindo vazios): {estrutura['planilha_rentabilidade']['registros_brutos']}")
            
            print("\n" + "=" * 50)
            print("VERIFICAÇÃO CONCLUÍDA: Sistema plenamente funcional")
            print("=" * 50)
            
            logger.info("Verificação completa concluída com sucesso")
            return True
            
        except Exception as e:
            error_msg = f"Erro na verificação completa: {e}"
            logger.error(error_msg)
            print(f"ERRO: {error_msg}")
            return False


def verificar_status_sistema() -> bool:
    """
    Verifica status geral do sistema de forma simplificada.
    
    Returns:
        bool: True se sistema está funcional
    """
    try:
        logger.info("Verificando status do sistema")
        
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        print("STATUS DO SISTEMA MMZR")
        print("=" * 30)
        print(f"Planilha base: {os.path.basename(planilha_base)}")
        print(f"Planilha rentabilidade: {os.path.basename(planilha_rentabilidade)}")
        
        # Verificar dados básicos
        estrutura = MMZRCompatibilidade.verificar_estrutura_dados()
        if estrutura["status"] == "sucesso":
            print(f"Clientes únicos: {estrutura['planilha_base']['clientes_validos']}")
            print(f"Total de carteiras: {estrutura['planilha_base']['carteiras_total']}")
            print(f"Registros de rentabilidade válidos: {estrutura['planilha_rentabilidade']['registros_total']}")
            print("\nSTATUS: Sistema funcionando corretamente")
            return True
        else:
            print(f"ERRO: {estrutura['mensagem']}")
            print("\nSTATUS: Sistema com problemas")
            return False
            
    except Exception as e:
        error_msg = f"Erro no sistema: {e}"
        logger.error(error_msg)
        print(f"ERRO: {error_msg}")
        print("\nDica: Verifique se há arquivos Excel em documentos/dados/")
        return False


def executar_diagnostico_completo() -> bool:
    """
    Executa diagnóstico completo do sistema.
    
    Returns:
        bool: True se sistema está plenamente funcional
    """
    return MMZRCompatibilidade.executar_verificacao_completa()


# Interface de linha de comando
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--status":
            sucesso = verificar_status_sistema()
            sys.exit(0 if sucesso else 1)
        elif sys.argv[1] == "--diagnostico":
            sucesso = executar_diagnostico_completo()
            sys.exit(0 if sucesso else 1)
        else:
            print("Uso: python diagnostico.py [--status | --diagnostico]")
            sys.exit(1)
    else:
        # Executar verificação de status por padrão
        sucesso = verificar_status_sistema()
        sys.exit(0 if sucesso else 1) 