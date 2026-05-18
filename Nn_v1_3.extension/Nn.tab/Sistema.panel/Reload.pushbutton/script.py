# -*- coding: utf-8 -*-
__title__ = 'Recarregar\npyRevit'
__author__ = 'NnBim Dev'

import time
from pyrevit.loader import sessionmgr

def recarregar_sessao():
    # Pausa para garantir que o Windows liberou os arquivos
    time.sleep(2)
    # Comando oficial de recarregamento
    sessionmgr.reload_pyrevit()

if __name__ == '__main__':
    recarregar_sessao()

