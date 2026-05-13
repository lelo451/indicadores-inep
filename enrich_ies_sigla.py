# -*- coding: utf-8 -*-
"""Busca a Sigla, Organização Acadêmica e Categoria Administrativa das IES
no e-MEC para preencher as colunas faltantes em indicadores_consolidados.xlsx.

Para cada IES com 'Sigla da IES' vazia/0, faz uma consulta no formulário
'Consulta Avançada' (https://emec.mec.gov.br/emec/nova#avancada) procurando
pelo Nome da IES. A tabela de resultados tem colunas bem definidas
(Instituição, Sigla, Município/UF, Org. Acadêmica, Cat. Administrativa),
então extraímos cada campo do td correspondente — bem mais robusto do que
parsear a célula 'Nome da IES - Sigla' da página de detalhes.

Cada linha de resultado tem o código no texto da td[0] ('(N) NOME') e
também no onclick da lupa (td[8]: 'abrirPopUpDetalhesIes( N )'). Quando o
código da nossa base aparece entre os candidatos, usamos esse registro
(match_type='code_match'); caso contrário, escolhemos o melhor nome
(match_type='name_match') ou marcamos 'not_found'.

Resultados ficam em Microdados/ies_siglas.csv e o consolidador os usa
em fill_sigla_from_cache / fill_org_categ_from_cache.

Requisitos:
- undetected-chromedriver (pip install undetected-chromedriver)
- Chrome real em /usr/bin/google-chrome
- DISPLAY disponível (Cloudflare bloqueia modo headless)
"""

from __future__ import annotations

import base64
import csv
import os
import re
import time
import unicodedata
from pathlib import Path

import pandas as pd

DATA_ROOT = Path(__file__).parent.resolve()
INDICADORES = DATA_ROOT / "indicadores_consolidados.xlsx"
CACHE = DATA_ROOT / "Microdados" / "ies_siglas.csv"

EMEC_FORM = "https://emec.mec.gov.br/emec/nova#avancada"
EMEC_DETAIL = (
    "https://emec.mec.gov.br/emec/consulta-cadastro/detalhes-ies/"
    "d96957f455f6405d14c6542552b0f6eb/{b64}"
)

# Notas administrativas que o e-MEC anexa depois da sigla na célula de nome
ADMIN_NOTE_RE = re.compile(
    r"\s+(?:Unifica[çc][ãa]o|Ades[ãa]o|Extinta|Processo n|Transformada|Migra[çc][ãa]o|"
    r"Recredenciament|Credenciament|Situa[çc][ãa]o)\b.*$",
    re.IGNORECASE,
)


def _encode_code(code) -> str:
    return base64.b64encode(str(code).encode("utf-8")).decode("utf-8")

CACHE_FIELDS = [
    "codigo_ies", "sigla", "nome_emec",
    "organizacao_emec", "categoria_emec",
    "match_type", "fetched_at",
]

# Quão similar dois nomes precisam ser para um 'name_match' ser aceito.
NAME_MATCH_THRESHOLD = 0.5

CLOUDFLARE_TITLES = ("just a moment", "checking your browser", "verifying")


# ---------------------------------------------------------------------------
# Comparação de nomes (tolera typos e reordenamentos)
# ---------------------------------------------------------------------------

def _norm_name(value) -> str:
    s = unicodedata.normalize("NFKD", str(value))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def _name_similarity(a: str, b: str) -> float:
    """Maior entre Jaccard (tokens) e SequenceMatcher (char-level)."""
    from difflib import SequenceMatcher

    norm_a = _norm_name(a)
    norm_b = _norm_name(b)
    if not norm_a or not norm_b:
        return 0.0
    tokens_a = set(norm_a.split())
    tokens_b = set(norm_b.split())
    jaccard = len(tokens_a & tokens_b) / len(tokens_a | tokens_b) if tokens_a and tokens_b else 0.0
    char_ratio = SequenceMatcher(None, norm_a, norm_b).ratio()
    return max(jaccard, char_ratio)


# ---------------------------------------------------------------------------
# Selenium helpers (Chrome + Cloudflare)
# ---------------------------------------------------------------------------

def _start_browser():
    import undetected_chromedriver as uc

    opts = uc.ChromeOptions()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1280,800")
    return uc.Chrome(
        options=opts,
        browser_executable_path="/usr/bin/google-chrome",
        version_main=None,
    )


def _wait_for_real_page(driver, target_locator, timeout: float = 120) -> bool:
    """Espera o Cloudflare resolver e o elemento alvo existir."""
    by, value = target_locator
    deadline = time.time() + timeout
    last_report = 0.0
    while time.time() < deadline:
        try:
            if driver.find_element(by, value) is not None:
                return True
        except Exception:
            pass
        if time.time() - last_report >= 10:
            try:
                title = driver.title
            except Exception:
                title = "<unavailable>"
            print(f"      ...waiting (title={title!r})", flush=True)
            last_report = time.time()
        time.sleep(1)
    return False


# ---------------------------------------------------------------------------
# Consulta Avançada
# ---------------------------------------------------------------------------

def _parse_result_row(tr) -> dict | None:
    """Extrai {codigo, nome, sigla, organizacao, categoria, municipio_uf}
    de uma <tr class='linha_tr_body_nova_grid'>. Colunas:
      td[0]: '(codigo) NOME'  td[1]: sigla  td[2]: município/UF
      td[3]: organização       td[4]: categoria
      td[8]: <img onclick='...abrirPopUpDetalhesIes( N )'>
    """
    tds = tr.find_all("td")
    if len(tds) < 5:
        return None
    nome_td_text = tds[0].get_text(" ", strip=True)
    m = re.match(r"\s*\((\d+)\)\s*(.+?)\s*$", nome_td_text)
    if m:
        codigo = m.group(1)
        nome = m.group(2).strip()
    else:
        if len(tds) >= 9:
            img = tds[8].find("img", attrs={"onclick": True})
            if img:
                m2 = re.search(r"abrirPopUpDetalhesIes\(\s*(\d+)", img.get("onclick", ""))
                if m2:
                    codigo = m2.group(1)
                    nome = nome_td_text
                else:
                    return None
            else:
                return None
        else:
            return None
    # Strip the admin notes that eMEC concatenates onto the name cell
    nome = ADMIN_NOTE_RE.sub("", nome).strip()
    nome = re.sub(r"\s*-\s*$", "", nome).strip()
    sigla = tds[1].get_text(" ", strip=True)
    municipio_uf = tds[2].get_text(" ", strip=True)
    organizacao = tds[3].get_text(" ", strip=True)
    categoria = tds[4].get_text(" ", strip=True)
    return {
        "codigo": codigo,
        "nome": nome,
        "sigla": sigla or None,
        "municipio_uf": municipio_uf or None,
        "organizacao": organizacao or None,
        "categoria": categoria or None,
    }


def search_consulta_avancada(driver, name: str, timeout: float = 120) -> list[dict]:
    """Submete o nome no formulário Consulta Avançada e devolve
    [{codigo, nome, sigla, organizacao, categoria, municipio_uf}, ...].

    Limpa o filtro de situação para incluir IES extintas/unificadas. Sai
    cedo se aparecer 'Nenhum registro encontrado!'."""
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By

    # Pequena pausa antes de cada busca para não bater no e-MEC em rajada.
    time.sleep(0.5)

    driver.get(EMEC_FORM)
    if not _wait_for_real_page(driver, (By.ID, "txt_no_ies"), timeout=timeout):
        try:
            link = driver.find_element(By.PARTIAL_LINK_TEXT, "Consulta Avançada")
            driver.execute_script("arguments[0].click();", link)
        except Exception:
            return []
        if not _wait_for_real_page(driver, (By.ID, "txt_no_ies"), timeout=30):
            return []

    # Seleciona radio "IES"
    for radio in driver.find_elements(By.NAME, "data[CONSULTA_AVANCADA][rad_buscar_por]"):
        if (radio.get_attribute("value") or "").upper() == "IES":
            try:
                driver.execute_script("arguments[0].click();", radio)
            except Exception:
                pass
            break

    # Inclui IES inativas no resultado
    try:
        driver.execute_script(
            "var s = document.getElementById('sel_co_situacao_funcionamento_ies');"
            "if (s) { s.value = ''; }"
        )
    except Exception:
        pass

    name_input = driver.find_element(By.ID, "txt_no_ies")
    name_input.clear()
    name_input.send_keys(name)

    driver.execute_script(
        "arguments[0].click();", driver.find_element(By.ID, "btnPesqAvancada")
    )

    # Aguarda resultado ou mensagem 'Nenhum registro encontrado!'
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            src = driver.page_source
        except Exception:
            src = ""
        if "Nenhum registro encontrado" in src:
            return []
        if driver.find_elements(By.CSS_SELECTOR, "tr.linha_tr_body_nova_grid"):
            break
        time.sleep(0.5)
    else:
        return []

    soup = BeautifulSoup(driver.page_source, "html.parser")
    container = soup.find(id="div_listar_consulta_avancada")
    if not container:
        return []
    results = []
    for tr in container.select("tr.linha_tr_body_nova_grid"):
        parsed = _parse_result_row(tr)
        if parsed:
            results.append(parsed)
    return results


def pick_match(candidates: list[dict], target_code: str, target_name: str) -> tuple[dict | None, str]:
    """Escolhe o melhor candidato. Prioriza coincidência de código; depois
    similaridade de nome acima do threshold. Devolve (candidato, match_type)."""
    if not candidates:
        return None, "not_found"
    target_code = str(target_code).strip()
    for c in candidates:
        if c["codigo"] == target_code:
            return c, "code_match"
    scored = sorted(
        ((c, _name_similarity(target_name, c["nome"])) for c in candidates),
        key=lambda x: x[1],
        reverse=True,
    )
    best, score = scored[0]
    if score >= NAME_MATCH_THRESHOLD:
        return best, "name_match"
    return None, "code_mismatch"


# ---------------------------------------------------------------------------
# Cache (Microdados/ies_siglas.csv)
# ---------------------------------------------------------------------------

def load_cache() -> dict[str, dict]:
    if not CACHE.exists():
        return {}
    with CACHE.open(encoding="utf-8") as f:
        return {row["codigo_ies"]: row for row in csv.DictReader(f)}


def save_cache(cache: dict[str, dict]) -> None:
    CACHE.parent.mkdir(parents=True, exist_ok=True)

    def sort_key(r):
        try:
            return (0, int(r["codigo_ies"]))
        except ValueError:
            return (1, r["codigo_ies"])

    with CACHE.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CACHE_FIELDS)
        writer.writeheader()
        for row in sorted(cache.values(), key=sort_key):
            writer.writerow({k: row.get(k, "") for k in CACHE_FIELDS})


def needs_lookup(df: pd.DataFrame) -> pd.DataFrame:
    s = df["Sigla da IES"].astype("string").str.strip()
    mask = s.isna() | s.isin(["", "0", "0.0"])
    pairs = (
        df.loc[mask, ["Código da IES", "Nome da IES"]]
        .dropna(subset=["Código da IES"])
        .drop_duplicates()
        .copy()
    )
    pairs["Código da IES"] = pairs["Código da IES"].astype(str).str.strip()
    pairs = pairs[pairs["Código da IES"].str.match(r"^\d+$", na=False)]
    return pairs.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def main() -> None:
    from selenium.common.exceptions import InvalidSessionIdException, WebDriverException

    df = pd.read_excel(INDICADORES, dtype={"Código da IES": str})
    pairs = needs_lookup(df)
    cache = load_cache()
    todo = [
        (row["Código da IES"], row["Nome da IES"])
        for _, row in pairs.iterrows()
        if row["Código da IES"] not in cache
    ]
    print(f"{len(pairs)} IES sem sigla; {len(todo)} ainda não estão no cache.", flush=True)
    if not todo:
        return

    driver = _start_browser()
    today = time.strftime("%Y-%m-%d")

    def safe_call(fn, *args, **kwargs):
        nonlocal driver
        try:
            return fn(driver, *args, **kwargs)
        except (InvalidSessionIdException, WebDriverException) as exc:
            print(f"      browser died ({type(exc).__name__}); restarting...", flush=True)
            try:
                driver.quit()
            except Exception:
                pass
            driver = _start_browser()
            try:
                return fn(driver, *args, **kwargs)
            except Exception as exc2:
                print(f"      retry failed: {exc2}", flush=True)
                return None

    try:
        for i, (code, name) in enumerate(todo, 1):
            print(f"[{i}/{len(todo)}] code={code} name={name!r}", flush=True)
            # 1) Busca por código primeiro (o campo aceita um número e devolve
            # exatamente uma linha, caso o código exista).
            candidates = safe_call(search_consulta_avancada, code) or []
            chosen, match_type = pick_match(candidates, code, name)
            # 2) Se o código não trouxe nada utilizável, tenta por nome
            if not chosen:
                candidates_name = safe_call(search_consulta_avancada, name) or []
                chosen2, match_type2 = pick_match(candidates_name, code, name)
                if chosen2:
                    chosen, match_type = chosen2, match_type2
            if chosen:
                print(
                    f"      {match_type}: code={chosen['codigo']} sigla={chosen['sigla']!r} "
                    f"org={chosen['organizacao']!r} cat={chosen['categoria']!r}",
                    flush=True,
                )
            else:
                print(f"      {match_type}", flush=True)

            picked = chosen or {}
            cache[code] = {
                "codigo_ies": code,
                "sigla": picked.get("sigla") or "",
                "nome_emec": picked.get("nome") or "",
                "organizacao_emec": picked.get("organizacao") or "",
                "categoria_emec": picked.get("categoria") or "",
                "match_type": match_type,
                "fetched_at": today,
            }
            if i % 10 == 0:
                save_cache(cache)
            time.sleep(1)
    finally:
        save_cache(cache)
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
