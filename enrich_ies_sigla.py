# -*- coding: utf-8 -*-
"""Busca a sigla das IES no e-MEC para preencher 'Sigla da IES' onde está
vazia/0 no consolidado.

A primeira tentativa usa o Código da IES. Se a página do e-MEC para aquele
código não carrega ou retorna um nome muito diferente do que temos em
'Nome da IES', registramos o resultado como 'code_mismatch' para inspeção
posterior (busca por nome via formulário fica como TODO — o e-MEC bloqueia
URLs simples e exige interação no formulário).

Resultados são gravados em 'Microdados/ies_siglas.csv'. O consolidador lê
esse cache em 'fill_sigla_from_cache' e usa para preencher a coluna.

Requisitos:
- undetected-chromedriver  (pip install undetected-chromedriver)
- Chrome real instalado em /usr/bin/google-chrome
- DISPLAY disponível (e-MEC é protegido por Cloudflare; modo headless é
  bloqueado, então o navegador abre uma janela)
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

EMEC_DETAIL = (
    "https://emec.mec.gov.br/emec/consulta-cadastro/detalhes-ies/"
    "d96957f455f6405d14c6542552b0f6eb/{b64}"
)

CACHE_FIELDS = [
    "codigo_ies", "sigla", "nome_emec",
    "organizacao_emec", "categoria_emec",
    "match_type", "fetched_at",
]

# Quão similar (Jaccard sobre tokens normalizados) os dois nomes precisam
# ser para considerarmos que o e-MEC retornou a IES certa.
NAME_MATCH_THRESHOLD = 0.5


def _encode_code(code) -> str:
    return base64.b64encode(str(code).encode("utf-8")).decode("utf-8")


def _norm_name(value) -> str:
    s = unicodedata.normalize("NFKD", str(value))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def _name_similarity(a: str, b: str) -> float:
    """Combina similaridade Jaccard (tokens) com SequenceMatcher (char-level).
    Jaccard pega bem reordenamentos; SequenceMatcher tolera typos ('Facudade'
    vs 'Faculdade'). Usamos o maior dos dois para ser generoso com matches."""
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


# Notas administrativas que o e-MEC concatena depois da sigla na mesma célula.
# Cortamos a partir delas para não contaminarem o nome/sigla extraídos.
ADMIN_NOTE_RE = re.compile(
    r"\s+(?:Unifica[çc][ãa]o|Ades[ãa]o|Extinta|Processo n|Transformada|Migra[çc][ãa]o|"
    r"Recredenciament|Credenciament|Situa[çc][ãa]o)\b.*$",
    re.IGNORECASE,
)


def _parse_label_value(value: str) -> dict | None:
    """Parse '(codigo) NOME - SIGLA [notas administrativas]' em {codigo, nome, sigla}.

    O e-MEC costuma anexar texto adicional (ex.: 'Unificação de Mantidas:
    Processo nº...') na mesma célula da sigla. Removemos esse rastro antes de
    separar nome/sigla; depois usamos rpartition no último ' - ' (siglas como
    'UNI-BAN' contêm hífen, então split simples atrapalha)."""
    m = re.match(r"\s*\((\d+)\)\s*(.+?)\s*$", value)
    if not m:
        return None
    codigo = m.group(1)
    rest = ADMIN_NOTE_RE.sub("", m.group(2))
    # Remove um traço residual deixado por 'NOME - <nota removida>'
    rest = re.sub(r"\s*-\s*$", "", rest).strip()
    if " - " in rest:
        nome, _, sigla = rest.rpartition(" - ")
        return {"codigo": codigo, "nome": nome.strip(), "sigla": sigla.strip() or None}
    return {"codigo": codigo, "nome": rest, "sigla": None}


def _extract_value_for_label(palette, label_prefix: str) -> str | None:
    """Encontra a célula cujo texto começa com label_prefix e retorna o texto
    da célula seguinte. Vários tds aninhados podem começar com o mesmo
    prefixo (o pai contém o filho na busca recursiva), então pulamos
    candidatos sem irmão td e continuamos procurando."""
    for td in palette.find_all("td"):
        text = td.get_text(" ", strip=True)
        if not text.startswith(label_prefix):
            continue
        value_td = td.find_next_sibling("td")
        if not value_td:
            continue
        div = value_td.find("div")
        return (div.get_text(" ", strip=True) if div else value_td.get_text(" ", strip=True))
    return None


def _parse_ies_detail(html: str) -> dict | None:
    """Retorna {codigo, nome, sigla, organizacao, categoria} da página de
    detalhes do IES."""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    palette = soup.find(id="paletaCadastro")
    if not palette:
        return None
    nome_sigla = _extract_value_for_label(palette, "Nome da IES")
    if not nome_sigla:
        return None
    parsed = _parse_label_value(nome_sigla)
    if not parsed:
        return None
    parsed["organizacao"] = _extract_value_for_label(palette, "Organização Acadêmica")
    parsed["categoria"] = _extract_value_for_label(palette, "Categoria Administrativa")
    return parsed


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


CLOUDFLARE_TITLES = ("just a moment", "checking your browser", "verifying")


def _wait_for_real_page(driver, target_locator, timeout: float = 120) -> bool:
    """Espera o desafio do Cloudflare resolver e o elemento alvo aparecer.
    Pesquisa o elemento a cada 1s; imprime progresso a cada 10s."""
    by, value = target_locator
    deadline = time.time() + timeout
    last_report = 0.0
    while time.time() < deadline:
        try:
            el = driver.find_element(by, value)
            # Aceita qualquer aparição do elemento (visível ou não); o caller
            # confirma o conteúdo logo em seguida.
            if el is not None:
                return True
        except Exception:
            # NoSuchElement, StaleElement, sessão errada — tudo se traduz em
            # "ainda não pronto"; continua aguardando.
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


# XPath para o subtítulo "IES" dentro de paletaCadastro — só aparece depois
# que o conteúdo é carregado por AJAX dentro do wrapper inicial.
IES_SUBTITLE_XPATH = "/html/body/div[1]/div[2]/div/table/tbody/tr[3]/td/div"


def fetch_by_code(driver, code: str, timeout: float = 120) -> dict | None:
    """Carrega a página do código no e-MEC e devolve o registro analisado.

    A URL /detalhes-ies devolve só um wrapper; o conteúdo real é carregado
    por jQuery .load() dentro de div#div_conteudo, e a página inteira só
    fica utilizável depois que o AJAX termina. Esperamos:
    1) Cloudflare resolver (#paletaCadastro existir)
    2) AJAX popular o conteúdo (o subtítulo 'IES' aparecer)"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    driver.get(EMEC_DETAIL.format(b64=_encode_code(code)))

    if not _wait_for_real_page(driver, (By.ID, "paletaCadastro"), timeout=timeout):
        print(f"      paletaCadastro nunca apareceu (Cloudflare?)", flush=True)
        return None

    try:
        WebDriverWait(driver, 90).until(
            EC.presence_of_element_located((By.XPATH, IES_SUBTITLE_XPATH))
        )
    except TimeoutException:
        print(f"      AJAX não carregou conteúdo dentro do paletaCadastro", flush=True)
        return None
    time.sleep(1)
    return _parse_ies_detail(driver.page_source)


def fetch_by_name(driver, name: str, timeout: float = 120) -> list[dict]:
    """Pesquisa o e-MEC pela aba 'Consulta Avançada' filtrando por Nome da IES
    e retorna [{codigo, nome, sigla}, ...]."""
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    driver.get("https://emec.mec.gov.br/emec/nova#avancada")
    if not _wait_for_real_page(driver, (By.ID, "txt_no_ies"), timeout=timeout):
        # tenta clicar no link da aba avançada (Cloudflare pode ter passado mas a aba não)
        try:
            link = driver.find_element(By.PARTIAL_LINK_TEXT, "Consulta Avançada")
            driver.execute_script("arguments[0].click();", link)
        except Exception:
            return []
        if not _wait_for_real_page(driver, (By.ID, "txt_no_ies"), timeout=30):
            return []

    # garante interatividade (campo deve estar visível e habilitado)
    try:
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "txt_no_ies")))
    except TimeoutException:
        return []

    for radio in driver.find_elements(By.NAME, "data[CONSULTA_AVANCADA][rad_buscar_por]"):
        if (radio.get_attribute("value") or "").upper() == "IES":
            try:
                driver.execute_script("arguments[0].click();", radio)
            except Exception:
                pass
            break

    # Limpa o filtro de situação (default 'Ativa' = 10035) para que IES
    # extintas/unificadas também apareçam — código antigo da nossa base muitas
    # vezes aponta para IES inativa.
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

    btn = driver.find_element(By.ID, "btnPesqAvancada")
    driver.execute_script("arguments[0].click();", btn)

    # O AJAX devolve o resultado dentro de #div_listar_consulta_avancada.
    # Esperamos por: linhas de resultado, OU mensagem 'Nenhum registro
    # encontrado!' — o que vier primeiro.
    deadline = time.time() + 60
    no_results = False
    found_rows = False
    while time.time() < deadline:
        try:
            src = driver.page_source
        except Exception:
            src = ""
        if "Nenhum registro encontrado" in src:
            no_results = True
            break
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.linha_tr_body_nova_grid")
            if rows:
                found_rows = True
                break
        except Exception:
            pass
        time.sleep(0.5)

    if no_results or not found_rows:
        return []

    soup = BeautifulSoup(driver.page_source, "html.parser")
    container = soup.find(id="div_listar_consulta_avancada")
    if not container:
        return []
    results: list[dict] = []
    for tr in container.select("tr.linha_tr_body_nova_grid"):
        text = tr.get_text(" ", strip=True)
        parsed = _parse_label_value(text)
        if parsed:
            results.append(parsed)
    return results


def pick_best_match(candidates: list[dict], our_name: str) -> dict | None:
    """Escolhe o resultado com maior similaridade com nosso Nome da IES,
    desde que passe o threshold mínimo."""
    if not candidates:
        return None
    scored = sorted(
        ((cand, _name_similarity(our_name, cand["nome"])) for cand in candidates),
        key=lambda x: x[1],
        reverse=True,
    )
    best, score = scored[0]
    return best if score >= NAME_MATCH_THRESHOLD else None


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
    """Retorna pares únicos (código, nome) cuja Sigla está vazia/0."""
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


def main() -> None:
    df = pd.read_excel(INDICADORES, dtype={"Código da IES": str})
    pairs = needs_lookup(df)
    cache = load_cache()
    todo = [
        (row["Código da IES"], row["Nome da IES"])
        for _, row in pairs.iterrows()
        if row["Código da IES"] not in cache
    ]
    print(f"{len(pairs)} IES sem sigla; {len(todo)} ainda não estão no cache.")
    if not todo:
        return

    from selenium.common.exceptions import InvalidSessionIdException, WebDriverException

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
            parsed = safe_call(fetch_by_code, code)
            match_type = None
            chosen = None
            if parsed:
                similarity = _name_similarity(name, parsed["nome"])
                if similarity >= NAME_MATCH_THRESHOLD:
                    chosen = parsed
                    match_type = "code_match"
                else:
                    print(f"      code mismatch (sim={similarity:.2f}); trying name search...", flush=True)
            else:
                print(f"      code not found; trying name search...", flush=True)

            if chosen is None:
                candidates = safe_call(fetch_by_name, name) or []
                best = pick_best_match(candidates, name)
                if best:
                    full = safe_call(fetch_by_code, best["codigo"])
                    chosen = full or best
                    match_type = "name_match"
                    print(f"      name match: code={best['codigo']} sigla={chosen.get('sigla')!r}", flush=True)
                else:
                    match_type = "not_found" if not parsed else "code_mismatch"

            picked = chosen or parsed or {}
            cache[code] = {
                "codigo_ies": code,
                "sigla": picked.get("sigla") or "",
                "nome_emec": picked.get("nome") or "",
                "organizacao_emec": picked.get("organizacao") or "",
                "categoria_emec": picked.get("categoria") or "",
                "match_type": match_type or "not_found",
                "fetched_at": today,
            }
            if i % 10 == 0:
                save_cache(cache)
            time.sleep(1)
    finally:
        save_cache(cache)
        driver.quit()


if __name__ == "__main__":
    main()
