# -*- coding: utf-8 -*-
import os

def limpeza_profunda_unicode():
    # Detecta onde o script está rodando
    raiz = os.path.dirname(os.path.abspath(__file__))
    print("Iniciando Limpeza Profunda em: {}\n".format(raiz))
    
    contador = 0
    
    for root, dirs, files in os.walk(raiz):
        for file in files:
            if file.endswith(('.py', '.yaml', '.json', '.xaml')):
                caminho = os.path.join(root, file)
                
                # Pula este próprio script
                if "FIX_DNA_V3" in file: continue

                try:
                    # Tenta ler o arquivo de várias formas
                    content = None
                    for encoding in ['utf-8-sig', 'utf-16', 'cp1252', 'latin-1']:
                        try:
                            with open(caminho, 'rb') as f:
                                raw = f.read()
                            # Tenta decodificar o binário
                            content = raw.decode(encoding)
                            break
                        except:
                            continue
                    
                    if content:
                        # O SEGREDO: Salvar em binário forçando UTF-8 puro
                        # Remove caracteres nulos e garante o cabeçalho
                        content = content.replace('\x00', '') 
                        if file.endswith('.py') and "# -*- coding: utf-8 -*-" not in content:
                            content = "# -*- coding: utf-8 -*-\n" + content
                            
                        with open(caminho, 'wb') as f:
                            f.write(content.encode('utf-8'))
                        
                        print("[OK] Corrigido: {}".format(file))
                        contador += 1
                except Exception as e:
                    print("[ERRO] Falha em {}: {}".format(file, e))

    print("\nTOTAL DE ARQUIVOS NORMALIZADOS: {}".format(contador))
    print("DICA: De um RELOAD no pyRevit agora.")

if __name__ == "__main__":
    limpeza_profunda_unicode()