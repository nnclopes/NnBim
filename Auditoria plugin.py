# -*- coding: utf-8 -*-
"""
NnBim: AUDITORIA COM LOG ESTRUTURADO
DESCRICAO:
Realiza a varredura de integridade do plugin e gera um arquivo .log detalhado
em C:\Users\Nn_1tb\OneDrive\Documentos\GitHub\NnBim\Nn_v1_3.extension\lib\Logs.

COMO USAR:
1. Execute o script.
2. Verifique a tabela de resumo no Revit.
3. Consulte o arquivo 'audit_log.txt' para detalhes técnicos de erros.
"""

__title__ = 'Auditar\n(com Log)'
__author__ = 'Nn_Dev'

import os
import datetime
from pyrevit import forms, script

# --- CONFIGURAÇÕES DE CAMINHO ---
BASE_PATH = r'C:\Users\Nn_1tb\OneDrive\Documentos\GitHub\NnBim\Nn_v1_3.extension'
LOG_DIR = os.path.join(BASE_PATH, 'lib', 'Logs')

if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, 'audit_report_{}.log'.format(datetime.date.today()))

class NnLogger:
    def __init__(self, filepath):
        self.filepath = filepath
        
    def log(self, level, message):
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        line = "[{}][{}] {}\n".format(timestamp, level, message)
        with open(self.filepath, 'a') as f:
            f.write(line)

def realizar_auditoria():
    logger = NnLogger(LOG_FILE)
    logger.log("INFO", "Iniciando auditoria completa da extensão NnBim.")
    
    dados_tabela = []
    
    for root, dirs, files in os.walk(BASE_PATH):
        for directory in dirs:
            if directory.endswith('.pushbutton'):
                btn_full_path = os.path.join(root, directory)
                btn_name = directory.replace('.pushbutton', '')
                
                # Verificações
                check_script = os.path.exists(os.path.join(btn_full_path, 'script.py'))
                check_icon = os.path.exists(os.path.join(btn_full_path, 'icon.png'))
                
                status_final = "OK"
                
                if not check_script:
                    logger.log("ERROR", "Botão '{}' está sem script.py!".format(btn_name))
                    status_final = "ERRO"
                if not check_icon:
                    logger.log("WARNING", "Botão '{}' está sem ícone personalizado.".format(btn_name))
                    if status_final != "ERRO": status_final = "AVISO"
                
                dados_tabela.append([btn_name, status_final, "Sim" if check_icon else "Não"])

    logger.log("INFO", "Auditoria finalizada. Total de botões verificados: {}".format(len(dados_tabela)))
    return dados_tabela

# --- INTERFACE DE SAÍDA ---
output = script.get_output()
output.print_md("# 📋 Relatório de Auditoria e Logs")
print("Arquivo de Log gerado em: {}".format(LOG_FILE))

resultados = realizar_auditoria()

output.print_table(
    table_data=resultados,
    columns=['NOME DA FERRAMENTA', 'STATUS SISTEMA', 'POSSUI ÍCONE?']
)

# Alerta caso haja erros críticos
erros = [r for r in resultados if r[1] == "ERRO"]
if erros:
    forms.alert("Foram encontrados {} erros críticos. Verifique o Log!".format(len(erros)), icon="warning")