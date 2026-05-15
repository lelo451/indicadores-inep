# Indicadores de Qualidade da Educação Superior

Scripts para consolidar os indicadores de qualidade da educação superior
publicados pelo INEP (ENADE, CPC, IDD, IGC) em uma única planilha pronta para
análise, e para consolidar os dicionários de variáveis dos microdados ENADE e
IDD.

## Estrutura

```
.
├── pipeline.py                      # Entrada recomendada: roda consolidar → ies
├── normalizacao.py                  # Tabelas canônicas + helpers de texto (shared)
├── consolidar_indicadores.py        # Consolida todos os arquivos de data/
├── enrich_ies_sigla.py              # Raspa sigla/org./categoria do e-MEC
│
├── dicionario/                      # Pacote: consolidação dos dicionários INEP
│   ├── __init__.py                  #   engine de parsing
│   ├── enade.py                     #   entry point ENADE
│   ├── idd.py                       #   entry point IDD
│   └── __main__.py                  #   roda os dois (python -m dicionario)
│
├── ies/                             # Pacote: lista mestra de IES
│   ├── _common.py                   #   caminhos e helpers
│   ├── censo.py                     #   etapa 1: consolida Censo Superior
│   ├── final.py                     #   etapa 2: mescla e-MEC + Censo;
│   │                                #   etapa 3: reaplica metadados em
│   │                                #            indicadores_consolidados
│   └── __main__.py                  #   roda as três etapas (python -m ies)
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
│   ├── municipios_enade.csv
│   └── ies_siglas.csv               # Cache gerado por enrich_ies_sigla.py
│
├── Lista INEP Censo Superior/       # Listas anuais do Censo (input do ies.censo)
│   └── Lista_de_IES_YYYY.xlsx
│
├── list_ies.csv                     # Whitelist de IES (input)
│
└── outputs/                         # Todos os artefatos gerados
    ├── dicionario_enade_consolidado.xlsx
    ├── dicionario_idd_consolidado.xlsx
    ├── indicadores_consolidados.xlsx
    ├── lista_ies_consolidada.xlsx
    └── list_ies_final.xlsx
```

## Como rodar

Requer Python 3.10+. Instale as dependências com:

```bash
pip install -r requirements.txt
```

`undetected-chromedriver`, `selenium` e `beautifulsoup4` só são necessários se
você for rodar o `enrich_ies_sigla.py` (que abre um Chrome real para passar do
Cloudflare do e-MEC).

### Pipeline (ordem de dependência)

```
dicionario.enade  ┐
dicionario.idd    ┴─→ consolidar_indicadores  ┬─→ enrich_ies_sigla ┐
                                              │                     │
                                              └──────────────────┐  ├─→ ies.final ──→ reaplica em indicadores_consolidados
                                                                 │  │
                                                  ies.censo  ────┴──┘
```

`ies.final` combina três fontes em ordem de prioridade:

1. **e-MEC** (`Microdados/ies_siglas.csv`) — base oficial, marca `complemento='n'`.
2. **Censo da Educação Superior** (`outputs/lista_ies_consolidada.xlsx`) —
   para IES com `match_type=not_found` no e-MEC. Marca `complemento='y'`.
3. **Indicadores INEP** (`outputs/indicadores_consolidados.xlsx`) — recupera
   metadados de IES descredenciadas/fundidas que nem e-MEC nem Censo cobrem.
   Marca `complemento='i'`.

Após gerar `list_ies_final.xlsx`, o pacote `ies` reaplica `nome/sigla/
organizacao/categoria` em `indicadores_consolidados.xlsx` (pulando IES com
`complemento='i'`, que vieram dos próprios indicadores). Por isso, **sempre
rode `python -m ies` depois de `consolidar_indicadores.py`** — ou use
`pipeline.py`, que encadeia os dois.

Comandos:

```bash
# Pré-requisitos (rodam raramente):
python -m dicionario                     # → dicionários do INEP (ENADE + IDD)
python enrich_ies_sigla.py               # → Microdados/ies_siglas.csv (Chrome)

# Pipeline principal (consolidar + ies):
python pipeline.py                       # → todos os outputs/ regenerados

# Ou, separadamente:
python consolidar_indicadores.py         # → outputs/indicadores_consolidados.xlsx
python -m ies                            # → list_ies_final.xlsx + reaplica em indicadores
```

Todos os scripts resolvem caminhos relativos a partir do diretório onde estão
gravados — podem ser executados de qualquer `cwd`.

## Indicadores consolidados

O arquivo `indicadores_consolidados.xlsx` tem três abas:

**Aba `Dados`** — granularidade de curso (Ano + Código do Curso, com
LEFT JOIN de IGC por Código da IES). Atributos da IES (Sigla/Organização/
Categoria) e do município (Sigla da UF) ficam apenas nas respectivas
dimensões. Nome da IES e Município do Curso permanecem inline mas são
refrescados das abas dimensionais para consistência referencial.

| Coluna | Origem |
| --- | --- |
| Ano | filename ou coluna do arquivo |
| Código da Área, Área de Avaliação | ENADE, CPC, IDD |
| Grau Acadêmico | arquivos 2018+ |
| Código da IES, Nome da IES | todos (Nome refrescado da aba `IES`) |
| Código do Curso, Modalidade de Ensino | ENADE/CPC/IDD pós-2014 |
| Código do Município, Município do Curso | todos (Nome refrescado da aba `Municípios`) |
| Conceito Enade (Contínuo/Faixa) | arquivos ENADE e enade_cpc |
| CPC (Contínuo/Faixa) | arquivos CPC e enade_cpc |
| IDD (Contínuo/Faixa) | arquivos IDD; o Contínuo também vem de `Nota IDD` em CPCs antigos |
| IGC (Contínuo/Faixa) | arquivos IGC (junção por Código da IES) |

**Aba `IES`** — uma linha por `Código da IES`, com Sigla/Nome/Organização/
Categoria já normalizadas. Vem de `list_ies_final.xlsx` (fontes priorizadas:
e-MEC > Censo > Indicadores), sem a coluna diagnóstica `complemento`.

**Aba `Municípios`** — uma linha por `Código do Município` presente em
`Dados`, com Município do Curso (nome canônico IBGE) e Sigla da UF. Vem
do catálogo `Microdados/municipios.csv`.

Para juntar dimensões com `Dados` no Excel:
`=VLOOKUP([@[Código da IES]]; IES!A:E; n; 0)` e
`=VLOOKUP([@[Código do Município]]; Municípios!A:C; n; 0)`.

## Decisões de projeto

- **Aliases por ano.** INEP renomeia colunas a cada ciclo (`Cód.IES` ↔
  `Código da IES` ↔ `co_ies`). O dicionário `COLUMN_ALIASES` em
  `consolidar_indicadores.py` reúne todas as grafias observadas.
- **Códigos decodificados.** Arquivos pré-2010 trazem apenas códigos
  (`co_grupo=5` em vez de "Medicina Veterinária"). O consolidador usa as
  planilhas de dicionário do INEP (consolidadas previamente) para preencher os
  rótulos.
- **IES normalizada em duas camadas.** Nome/Sigla/Categoria/Organização de
  cada `Código da IES` são unificados em todas as linhas. Primeiro
  `normalize_ies_metadata` propaga o valor do ano mais recente dentro do
  próprio indicadores (faz "CEFET/PR" 2004 virar UTFPR como em 2010+).
  Depois a etapa 3 de `ies.final` sobrescreve com `list_ies_final.xlsx`
  (priorizando e-MEC > Censo). IES marcadas com `complemento='i'` em
  `list_ies_final` (descredenciadas/fundidas) são puladas nessa segunda
  camada para evitar circularidade — para essas, prevalece o valor do
  ano-mais-recente da primeira camada.
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
