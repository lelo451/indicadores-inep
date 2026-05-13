# Indicadores de Qualidade da Educação Superior

Scripts para consolidar os indicadores de qualidade da educação superior
publicados pelo INEP (ENADE, CPC, IDD, IGC) em uma única planilha pronta para
análise, e para consolidar os dicionários de variáveis dos microdados ENADE e
IDD.

## Estrutura

```
.
├── consolidar_indicadores.py        # Consolida todos os arquivos de data/
├── consolidar_dicionario_enade.py   # Consolida dicionários ENADE
├── consolidar_dicionario_idd.py     # Consolida dicionários IDD
├── dicionario_comum.py              # Módulo compartilhado entre os dois acima
│
├── data/                            # Planilhas de indicadores (renomeadas)
│   ├── IGC_2007.xlsx … IGC_2023.xlsx
│   ├── cpc_2007.xls  … cpc_2023.xlsx
│   ├── enade_2004.xls … enade_2023.xlsx
│   ├── enade_cpc_2008.xls … enade_cpc_2011.xls
│   └── idd_2016.xlsx … idd_2023.xlsx
│
├── Microdados/                      # Microdados e tabelas auxiliares
│   ├── enade/dicionarios enade YYYY.xlsx
│   ├── idd/dicionarios idd YYYY.xls(x)
│   ├── municipios.csv               # Catálogo IBGE (códigos + nomes)
│   └── municipios_enade.csv
│
├── dicionario_enade_consolidado.xlsx     # Saída do consolidador ENADE
├── dicionario_idd_consolidado.xlsx       # Saída do consolidador IDD
└── indicadores_consolidados.xlsx         # Saída do consolidador principal
```

## Como rodar

Requer Python 3.10+ com `pandas`, `openpyxl` e `xlrd` (para arquivos `.xls`):

```bash
pip install pandas openpyxl xlrd
```

Os três scripts são independentes. Rode na ordem:

```bash
python3 consolidar_dicionario_enade.py   # gera dicionario_enade_consolidado.xlsx
python3 consolidar_dicionario_idd.py     # gera dicionario_idd_consolidado.xlsx
python3 consolidar_indicadores.py        # gera indicadores_consolidados.xlsx
```

O consolidador principal usa os dois dicionários para decodificar variáveis
codificadas (`co_grupo`, `co_categad`, etc.) presentes em arquivos antigos.

Todos os scripts resolvem caminhos relativos a partir do diretório onde estão
gravados — podem ser executados de qualquer `cwd`.

## Indicadores consolidados

O arquivo `indicadores_consolidados.xlsx` traz 22 colunas canônicas por linha,
mesclando os arquivos por (Ano, Código do Curso) no nível de curso e fazendo
LEFT JOIN com o IGC no nível de IES.

| Coluna | Origem |
| --- | --- |
| Ano | filename ou coluna do arquivo |
| Código da Área, Área de Avaliação | ENADE, CPC, IDD |
| Grau Acadêmico | arquivos 2018+ |
| Código da IES, Nome da IES, Sigla da IES | todos |
| Organização Acadêmica, Categoria Administrativa | todos (em IGC pré-2017 a Organização é inferida da aba) |
| Código do Curso, Modalidade de Ensino | ENADE/CPC/IDD pós-2014 |
| Código do Município, Município do Curso, Sigla da UF | todos (município preenchido por catálogo IBGE quando faltava) |
| Conceito Enade (Contínuo/Faixa) | arquivos ENADE e enade_cpc |
| CPC (Contínuo/Faixa) | arquivos CPC e enade_cpc |
| IDD (Contínuo/Faixa) | arquivos IDD; o Contínuo também vem de `Nota IDD` em CPCs antigos |
| IGC (Contínuo/Faixa) | arquivos IGC (junção por Código da IES) |

## Decisões de projeto

- **Aliases por ano.** INEP renomeia colunas a cada ciclo (`Cód.IES` ↔
  `Código da IES` ↔ `co_ies`). O dicionário `COLUMN_ALIASES` em
  `consolidar_indicadores.py` reúne todas as grafias observadas.
- **Códigos decodificados.** Arquivos pré-2010 trazem apenas códigos
  (`co_grupo=5` em vez de "Medicina Veterinária"). O consolidador usa as
  planilhas de dicionário do INEP (consolidadas previamente) para preencher os
  rótulos.
- **IES normalizada para o ano mais recente.** Cada `Código da IES` é
  reescrito em todos os anos com o Nome/Sigla/Categoria/Organização mais
  recentes. Assim, "CEFET/PR" (2004) e "UTFPR" (2010+) — ambos IES código 588 —
  aparecem como UTFPR/Universidade/Pública Federal em todas as linhas.
- **IGC multi-aba.** Arquivos IGC pré-2017 dividem os dados por organização
  acadêmica em abas separadas (Universidades, Centros Universitários,
  Faculdades). A consolidação lê todas as abas relevantes e injeta a
  Organização Acadêmica pelo nome da aba.
- **Município por nome+UF.** Linhas que trazem o nome do município sem o
  código IBGE são resolvidas com `Microdados/municipios.csv`, mais um pequeno
  conjunto de aliases para grafias que o INEP escreveu fora do padrão IBGE
  (ex.: `Campos do Goytacazes` → `Campos dos Goytacazes`).

## Dicionários consolidados

Cada um destes arquivos traz uma aba `TODAS` mais uma aba por variável:

- `dicionario_enade_consolidado.xlsx` — ENADE 2004–2023
- `dicionario_idd_consolidado.xlsx` — IDD 2014–2023

Variáveis cobertas em ambos: `CO_CATEGAD`, `CO_ORGACAD`, `CO_GRUPO`,
`CO_MODALIDADE`, `CO_UF_CURSO`, `CO_MUNIC_CURSO`. Para cada variável só são
mantidos os nomes distintos (case- e accent-insensitive) e, quando o mesmo
nome aparece em vários anos, prevalece o ano mais recente.

`CO_MUNIC_CURSO` é populado a partir do catálogo IBGE em
`Microdados/municipios.csv` (5.569 municípios), já que o dicionário do INEP só
remete à planilha de municípios.
