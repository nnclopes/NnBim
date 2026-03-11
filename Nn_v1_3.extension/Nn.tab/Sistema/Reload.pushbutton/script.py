"""Recarrega a sessão do pyRevit para aplicar alterações nos scripts."""
# -*- coding=utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except

# Metadados do Botão
__title__ = 'Recarregar\npyRevit' # Quebra de linha no nome
__author__ = 'NnBim Dev'
__doc__ = 'Recarrega a sessão do pyRevit.\nUse isso após editar scripts para ver as mudanças sem fechar o Revit.'

from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

# Logger para debugar se necessário
logger = script.get_logger()

def recarregar_sessao():
    # Verifica se o clique veio da interface (botão)
    res = True
    if EXEC_PARAMS.executed_from_ui:
        # Pergunta de segurança (Padrão NnBim: Segurança em primeiro lugar)
        res = forms.alert(
            'Recarregar a sessão consome memória e limpa variáveis temporárias.\n\n'
            'Você deve usar isso quando:\n'
            ' - Adicionou/Removeu botões.\n'
            ' - Alterou ícones.\n'
            ' - Mudou códigos core (C# ou configurações).\n\n'
            'Tem certeza que deseja recarregar agora?',
            title='Sistema NnBim',
            ok=False, 
            yes=True, 
            no=True
        )

    if res:
        results = script.get_results()

        # Feedback visual no output window
        logger.info('Iniciando recarregamento do pyRevit...')
        
        # O comando mágico que recarrega tudo
        sessionmgr.reload_pyrevit()

        # Feedback final (geralmente a janela fecha sozinha ao recarregar, mas registramos o log)
        results.newsession = sessioninfo.get_session_uuid()

# Execução
if __name__ == '__main__':
    recarregar_sessao()