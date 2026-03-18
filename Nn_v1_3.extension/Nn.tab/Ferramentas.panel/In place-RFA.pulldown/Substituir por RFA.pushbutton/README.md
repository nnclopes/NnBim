# Ferramenta: Transformar\nem RFA

## DescriÃ§Ã£o
NnBim Ferramentas: Transformar em RFA
DESCRIÇÃO:
Converte elementos In-Loco ou DirectShapes soltos em Famílias Carregáveis (.rfa) externas.
O script extrai a geometria, abre um template de família (.rft), insere o sólido na origem 0,0,0
e carrega a nova família de volta para o projeto na posição exata do original.

COMO USAR:
1. Selecione os elementos soltos (DirectShapes) que deseja converter.
2. Selecione o arquivo de template (.rft) adequado para a categoria.
3. O script criará arquivos temporários e substituirá os originais por instâncias de família (.rfa).
