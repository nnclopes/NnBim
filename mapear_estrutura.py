# -*- coding: utf-8 -*-
import os

def diagnostico_full_nn(diretorio_raiz):
    log = []
    log.append("DIAGNÓSTICO DE RAIZ NnBim - v4.0")
    log.append("="*60)
    log.append("RAIZ: " + diretorio_raiz)
    log.append("="*60 + "\n")

    # 1. VERIFICAÇÃO DO EXTENSION.JSON (O MAIS IMPORTANTE)
    json_path = os.path.join(diretorio_raiz, "extension.json")
    if os.path.exists(json_path):
        log.append("[VITAL] extension.json: ENCONTRADO NA RAIZ (CORRETO)\n")
    else:
        log.append("[VITAL] extension.json: NÃO ENCONTRADO NA RAIZ! <--- [ERRO CRÍTICO]\n")

    def listar_arvore(caminho, nivel=0):
        try:
            itens = sorted(os.listdir(caminho))
        except:
            return

        for item in itens:
            if item in ['.git', '__pycache__', 'bin', 'obj']: continue
            
            caminho_full = os.path.join(caminho, item)
            indent = "    " * nivel
            
            if os.path.isdir(caminho_full):
                # Validar sufixos de pastas
                erro_pasta = ""
                if nivel == 0 and not item.endswith('.tab') and item not in ['lib', '_TEMPLATES']:
                    erro_pasta = " <--- [AVISO: Pasta na raiz sem sufixo .tab ou .panel]"
                
                log.append("{}|--- {}/{}".format(indent, item, erro_pasta))
                listar_arvore(caminho_full, nivel + 1)
            else:
                # Validar arquivos
                if item.lower() == 'script.py':
                    log.append("{}|    [OK] script.py".format(indent))
                elif item.lower() == 'extension.json' and nivel > 0:
                    log.append("{}|    [!] extension.json (REDUNDANTE/FORA DO LUGAR)")
                elif item.endswith(('.png', '.xaml')):
                    log.append("{}|    [F] {}".format(indent, item))

    listar_arvore(diretorio_raiz)
    return "\n".join(log)

if __name__ == "__main__":
    # CAMINHO RAIZ DA EXTENSÃO
    caminho_da_nn = r'C:\Users\Nn_1tb\OneDrive\Documentos\GitHub\NnBim\Nn_v1_3.extension'
    
    if os.path.exists(caminho_da_nn):
        resultado = diagnostico_full_nn(caminho_da_nn)
        with open("RELATORIO_FINAL_NN.txt", "w", encoding="utf-8") as f:
            f.write(resultado)
        print("Relatório gerado! Verifique 'RELATORIO_FINAL_NN.txt' e cole aqui.")
    else:
        print("Erro: O caminho especificado não foi encontrado.")