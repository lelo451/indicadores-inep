"""Pipeline para consolidar listas de IES do INEP e do e-MEC.

Etapa 1 (``ies.censo``): combina os arquivos anuais do Censo da Educação
Superior em ``outputs/lista_ies_consolidada.xlsx``.

Etapa 2 (``ies.final.run``): mescla a base do e-MEC (``Microdados/ies_siglas.csv``,
gerada por ``enrich_ies_sigla.py``) com o resultado da Etapa 1 e produz
``outputs/list_ies_final.xlsx``.

Etapa 3 (``ies.final.apply_to_indicadores``): reaplica
``nome/sigla/organizacao/categoria`` de ``list_ies_final.xlsx`` em
``outputs/indicadores_consolidados.xlsx``, pulando IES com ``complemento='i'``
(fonte circular). Também aplica ``ORG_ACAD_LEGACY_CODES`` para sanar códigos
numéricos antigos remanescentes.

Executar o pipeline completo: ``python -m ies``.
"""
