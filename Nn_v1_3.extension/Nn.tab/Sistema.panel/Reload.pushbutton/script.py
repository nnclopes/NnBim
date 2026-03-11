# -*- coding: utf-8 -*-
__title__ = 'Recarregar\npyRevit'
__author__ = 'NnBim Dev'
__doc__ = 'Recarrega a sessão do pyRevit aguardando 2 segundos.'

import time
from pyrevit.loader import sessionmgr

def recarregar_sessao():
    # Pausa a execução por exatos 2 segundos
    time.sleep(2)
    
    # O comando mágico que recarrega tudo
    sessionmgr.reload_pyrevit()

# Execução
if __name__ == '__main__':
    recarregar_sessao()