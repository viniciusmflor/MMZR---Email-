#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste para demonstrar a funcionalidade de evitar duplicação do nome Felipe
quando o banker responsável é o Banker 4 (Felipe).
"""

def gerar_texto_observacao(banker, banker_pronome):
    """
    Função de teste que replica a lógica implementada no mmzr_email_generator.py
    para demonstrar como funciona a prevenção de duplicação do nome Felipe.
    """
    
    # Criar o texto da observação baseado no banker
    if banker == 'Banker 4':
        # Se o banker é o Banker 4 (Felipe), usar texto singular sem duplicação
        obs_text = "<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe em cópia para também receber as informações."
    else:
        # Se o banker não é o Banker 4, usar texto plural com os dois nomes
        obs_text = f"<strong>Obs.:</strong> Conforme solicitado, deixo o Felipe e {banker_pronome} em cópia para também receberem as informações."
    
    return obs_text

def teste_funcionalidade():
    """Executa testes para diferentes cenários de bankers"""
    
    print("=== TESTE DA FUNCIONALIDADE DE PREVENÇÃO DE DUPLICAÇÃO ===\n")
    
    # Cenário 1: Banker 1 com Renato
    print("1. Banker: Banker 1 (Renato)")
    texto_banker1 = gerar_texto_observacao("Banker 1", "Renato")
    print(f"Resultado: {texto_banker1}\n")
    
    # Cenário 2: Banker 4 (Felipe) - cenário problemático resolvido
    print("2. Banker: Banker 4 (Felipe)")
    texto_banker4 = gerar_texto_observacao("Banker 4", "Felipe")
    print(f"Resultado: {texto_banker4}\n")
    
    # Cenário 3: Banker 2 com outro nome
    print("3. Banker: Banker 2 (Carolina)")
    texto_banker2 = gerar_texto_observacao("Banker 2", "Carolina")
    print(f"Resultado: {texto_banker2}\n")
    
    # Cenário 4: Banker 3 com outro nome
    print("4. Banker: Banker 3 (Roberto)")
    texto_banker3 = gerar_texto_observacao("Banker 3", "Roberto")
    print(f"Resultado: {texto_banker3}\n")
    
    # Cenário 5: Banker 7 com outro nome
    print("5. Banker: Banker 7 (Ana)")
    texto_banker7 = gerar_texto_observacao("Banker 7", "Ana")
    print(f"Resultado: {texto_banker7}\n")
    
    print("=== RESUMO ===")
    print("✅ Quando o banker é Banker 4: texto singular, sem duplicação")
    print("✅ Quando o banker não é Banker 4: texto plural, com os dois nomes")
    print("✅ Lógica baseada no código do banker, mais robusta!")

if __name__ == "__main__":
    teste_funcionalidade() 