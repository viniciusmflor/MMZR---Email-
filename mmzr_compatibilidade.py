"""
MMZR Family Office - Sistema de Compatibilidade
Detecção automática de planilhas

Autor: MMZR Family Office
"""

import os
import pandas as pd
from typing import Tuple, Dict, Any


class MMZRCompatibilidade:
    """Gerencia detecção automática de planilhas e compatibilidade do sistema."""
    
    @classmethod
    def get_planilhas_path(cls) -> Tuple[str, str]:
        """Detecta automaticamente as planilhas base e rentabilidade."""
        dados_dir = "documentos/dados"
        
        if not os.path.exists(dados_dir):
            raise FileNotFoundError(f"Pasta não encontrada: {dados_dir}")
        
        # Listar arquivos Excel
        excel_files = []
        for arquivo in os.listdir(dados_dir):
            if arquivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                excel_files.append(os.path.join(dados_dir, arquivo))
        
        if len(excel_files) == 0:
            raise FileNotFoundError("Nenhum arquivo Excel encontrado em documentos/dados/")
        
        if len(excel_files) == 1:
            # Se só há um arquivo, usar para ambos
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
                else:
                    # Senão, provavelmente é a de rentabilidade
                    planilha_rentabilidade = arquivo
                    
            except Exception as e:
                print(f"Erro ao verificar {arquivo}: {e}")
                continue
        
        # Se não identificou a base, usar o primeiro arquivo
        if not planilha_base:
            planilha_base = excel_files[0]
        
        # Se não identificou a rentabilidade, usar um arquivo diferente da base
        if not planilha_rentabilidade:
            for arquivo in excel_files:
                if arquivo != planilha_base:
                    planilha_rentabilidade = arquivo
                    break
            
            # Se ainda não tem, usar o mesmo arquivo
            if not planilha_rentabilidade:
                planilha_rentabilidade = planilha_base
        
        return planilha_base, planilha_rentabilidade
    
    @classmethod
    def testar_compatibilidade(cls) -> bool:
        """Testa compatibilidade do sistema."""
        try:
            # Testar acesso às planilhas
            planilha_base, planilha_rentabilidade = cls.get_planilhas_path()
            
            # Testar leitura das planilhas
            excel_base = pd.ExcelFile(planilha_base)
            excel_rent = pd.ExcelFile(planilha_rentabilidade)
            
            # Verificar abas essenciais
            if "Base Clientes" not in excel_base.sheet_names:
                print("❌ Aba 'Base Clientes' não encontrada na planilha base")
                return False
            
            # Testar leitura básica
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes", nrows=5)
            df_rent = pd.read_excel(excel_rent, sheet_name=excel_rent.sheet_names[0], nrows=5)
            
            print("✅ Sistema compatível e funcional")
            return True
            
        except Exception as e:
            print(f"❌ Erro de compatibilidade: {e}")
            return False
    
    @classmethod
    def verificar_estrutura_dados(cls) -> Dict[str, Any]:
        """Verifica a estrutura dos dados nas planilhas."""
        try:
            planilha_base, planilha_rentabilidade = cls.get_planilhas_path()
            
            # Verificar planilha base
            excel_base = pd.ExcelFile(planilha_base)
            df_clientes = pd.read_excel(excel_base, sheet_name="Base Clientes")
            
            # Verificar planilha de rentabilidade
            excel_rent = pd.ExcelFile(planilha_rentabilidade)
            df_rent = pd.read_excel(excel_rent, sheet_name=excel_rent.sheet_names[0])
            
            resultado = {
                "status": "sucesso",
                "planilha_base": {
                    "arquivo": os.path.basename(planilha_base),
                    "abas": excel_base.sheet_names,
                    "clientes_total": len(df_clientes)
                },
                "planilha_rentabilidade": {
                    "arquivo": os.path.basename(planilha_rentabilidade),
                    "abas": excel_rent.sheet_names,
                    "registros_total": len(df_rent)
                }
            }
            
            return resultado
            
        except Exception as e:
            return {
                "status": "erro",
                "mensagem": str(e)
            }


# Funcionalidades de teste e diagnóstico
def verificar_status_sistema():
    """Verifica status geral do sistema."""
    try:
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        
        print("=== STATUS DO SISTEMA MMZR ===\n")
        print(f"Planilha base detectada: {os.path.basename(planilha_base)}")
        print(f"Planilha rentabilidade detectada: {os.path.basename(planilha_rentabilidade)}")
        
        # Verificar dados
        estrutura = MMZRCompatibilidade.verificar_estrutura_dados()
        if estrutura["status"] == "sucesso":
            print(f"\nClientes disponíveis: {estrutura['planilha_base']['clientes_total']}")
            print(f"Registros de rentabilidade: {estrutura['planilha_rentabilidade']['registros_total']}")
            print("\n✅ Sistema funcionando corretamente")
            return True
        else:
            print(f"\n❌ Erro: {estrutura['mensagem']}")
            return False
            
    except Exception as e:
        print(f"❌ Erro no sistema: {e}")
        print("\nDica: Verifique se há arquivos Excel em documentos/dados/")
        return False


def executar_diagnostico_completo():
    """Executa diagnóstico completo do sistema."""
    print("=== DIAGNÓSTICO COMPLETO DO SISTEMA MMZR ===\n")
    
    try:
        # Verificar arquivos
        print("1. Detectando planilhas...")
        planilha_base, planilha_rentabilidade = MMZRCompatibilidade.get_planilhas_path()
        print(f"   Base: {os.path.basename(planilha_base)}")
        print(f"   Rentabilidade: {os.path.basename(planilha_rentabilidade)}")
        
        # Teste de compatibilidade
        print("\n2. Testando compatibilidade...")
        compativel = MMZRCompatibilidade.testar_compatibilidade()
        
        if not compativel:
            print("❌ Sistema não compatível. Verifique os arquivos de planilha.")
            return False
        
        # Verificar estrutura de dados
        print("\n3. Verificando estrutura de dados...")
        estrutura = MMZRCompatibilidade.verificar_estrutura_dados()
        
        if estrutura["status"] == "erro":
            print(f"❌ Erro na estrutura: {estrutura['mensagem']}")
            return False
        
        print(f"✅ Base: {estrutura['planilha_base']['clientes_total']} clientes")
        print(f"✅ Rentabilidade: {estrutura['planilha_rentabilidade']['registros_total']} registros")
        
        print("\n✅ DIAGNÓSTICO CONCLUÍDO - Sistema funcional!")
        return True
        
    except Exception as e:
        print(f"❌ Erro no diagnóstico: {e}")
        return False


# Interface de linha de comando
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        comando = sys.argv[1]
        
        if comando == "--diagnostico":
            executar_diagnostico_completo()
        elif comando == "--status":
            verificar_status_sistema()
        elif comando == "--testar":
            compativel = MMZRCompatibilidade.testar_compatibilidade()
            sys.exit(0 if compativel else 1)
        else:
            print("Comandos disponíveis:")
            print("  --diagnostico  : Executa diagnóstico completo")
            print("  --status       : Mostra status do sistema")
            print("  --testar       : Testa compatibilidade básica")
    else:
        # Por padrão, mostrar status
        verificar_status_sistema() 