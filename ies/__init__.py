"""Pipeline para consolidar listas de IES do INEP e do e-MEC.

Etapa 1 (``ies.censo``): combina os arquivos anuais do Censo da Educação
Superior em ``outputs/lista_ies_consolidada.xlsx``.

Etapa 2 (``ies.final``): mescla a base do e-MEC (``Microdados/ies_siglas.csv``,
gerada por ``enrich_ies_sigla.py``) com o resultado da Etapa 1 e produz
``outputs/list_ies_final.xlsx``.

Executar o pipeline completo: ``python -m ies``.
"""
