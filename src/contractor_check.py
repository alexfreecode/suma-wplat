"""
contractor_check.py — porównanie klientów z wyciągu z bazą kontrahentów
wyeksportowaną z Saldeo Smart (CSV).

Cel: ostrzec użytkownika PRZED importem faktur o sytuacjach prowadzących
do powstania w Saldeo zduplikowanych kart tego samego kontrahenta
(literówka, inna kolejność imienia i nazwiska, inna wielkość liter itp.),
a także poinformować o klientach, których w bazie Saldeo jeszcze nie ma —
dla nich przy imporcie zostanie utworzona nowa karta.

Użycie:
    contractors = load_saldeo_contractors("eksport_kontrahentow.csv")
    results = check_clients(["Jan Kowalski", ...], contractors)
    # results: [{"name": ..., "status": "exact"|"similar"|"new", "matched": ...}, ...]
"""

import csv
import re
import unicodedata
from difflib import SequenceMatcher

# Próg podobieństwa tekstów (0..1) dla oznaczenia „podejrzanie podobnej nazwy”.
# Dobrany empirycznie: wyłapuje literówki i drobne różnice, ale nie myli
# różnych osób o podobnych, lecz różnych nazwiskach.
SIMILARITY_THRESHOLD = 0.84


# ── Normalizacja nazw do porównania ──────────────────────────────────────────

def _normalize(s: str) -> str:
    """Sprowadza nazwę do postaci porównywalnej: bez diakrytyków, wielkości liter i zbędnych spacji."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s.strip().lower())


def _tokens(s: str) -> frozenset:
    """Zbiór słów nazwy — by wyłapać „Kowalski Jan” == „Jan Kowalski”."""
    return frozenset(_normalize(s).split())


def _sorted_norm(s: str) -> str:
    """
    Znormalizowana nazwa ze słowami w kolejności alfabetycznej — potrzebna,
    by porównanie rozmyte „widziało” podobieństwo także wtedy, gdy
    jednocześnie inna jest kolejność słów I występuje drobna literówka
    (np. „Jan Kowalski” vs „Kowalski Joan”).
    Zwykłe porównanie znak po znaku daje wtedy niski współczynnik z powodu
    przestawienia słów, a porównanie zbiorów tokenów wymaga dokładnej
    zgodności słów i nie wyłapuje literówki.
    """
    return " ".join(sorted(_normalize(s).split()))


# ── Wczytywanie bazy kontrahentów Saldeo ─────────────────────────────────────

def load_saldeo_contractors(csv_path: str) -> list[dict]:
    """
    Wczytuje eksport CSV bazy kontrahentów z Saldeo Smart
    (sekcja „Kontrahenci” → „Eksportuj”).

    Porównanie opiera się na kolumnach „Nazwa skrócona:” i „Nazwa pełna” —
    ich indeksy ustalane są na podstawie nagłówka pliku (na wypadek zmiany
    kolejności kolumn w przyszłych wersjach Saldeo).

    Zwraca listę {"short": ..., "full": ..., "nip": ...} z oryginalnymi
    wartościami (potrzebne do pokazania użytkownikowi). „nip” jest pusty,
    gdy kontrahent to osoba prywatna albo plik nie ma tej kolumny.
    """
    contractors: list[dict] = []

    # encoding="utf-8-sig" — eksporty Saldeo zwykle zawierają BOM
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        if not header:
            return contractors

        def _find(col_prefix: str):
            for i, h in enumerate(header):
                if h.strip().lower().startswith(col_prefix.lower()):
                    return i
            return None

        idx_short = _find("Nazwa skrócona")
        idx_full  = _find("Nazwa pełna")
        idx_nip   = _find("NIP")
        if idx_short is None and idx_full is None:
            return contractors

        for row in reader:
            if not row or all(not cell.strip() for cell in row):
                continue
            short = row[idx_short].strip() if idx_short is not None and idx_short < len(row) else ""
            full  = row[idx_full].strip()  if idx_full  is not None and idx_full  < len(row) else ""
            nip   = row[idx_nip].strip()   if idx_nip   is not None and idx_nip   < len(row) else ""
            if short or full:
                contractors.append({"short": short, "full": full, "nip": nip})

    return contractors


def names_with_nip(contractors: list[dict]) -> set[str]:
    """Znormalizowane nazwy kontrahentów, którzy mają w Saldeo NIP.

    Bank nie podaje NIP-u nadawcy, a jednoosobowa firma przychodzi na
    wyciągu tak samo jak osoba prywatna: imię, nazwisko, adres. NIP w bazie
    Saldeo jest więc jedyną wskazówką, że to firma. Służy tylko do
    podpowiedzi w raporcie, decyzję podejmuje użytkownik.
    """
    out: set[str] = set()
    for c in contractors:
        if not c.get("nip"):
            continue
        for name in (c.get("short", ""), c.get("full", "")):
            if name:
                out.add(_normalize(name))
    return out


def has_nip(client_name: str, nip_names: set[str]) -> bool:
    """Czy klient z wyciągu ma w bazie Saldeo kartę z NIP-em (dokładna nazwa)."""
    return _normalize(client_name) in nip_names


# ── Porównanie klientów z bazą ───────────────────────────────────────────────

def check_clients(client_names: list[str], contractors: list[dict]) -> list[dict]:
    """
    Porównuje nazwy klientów z wyciągu z bazą kontrahentów Saldeo.

    Zwraca listę {"name", "status", "matched"} dla każdej nazwy:
      "exact"   — dokładna zgodność (bez uwzględniania wielkości liter/spacji/
                  diakrytyków) — faktura podepnie się pod istniejącą kartę
                  kontrahenta;
      "similar" — znaleziono podejrzanie podobną, lecz nieidentyczną nazwę —
                  duże prawdopodobieństwo, że to ta sama osoba, a import
                  utworzy zduplikowaną kartę;
      "new"     — nie znaleziono żadnej zgodności — Saldeo utworzy nową kartę.

    Logika porównania używa „Nazwy pełnej” jako głównego celu wyszukiwania —
    jest ona zawsze bliższa rzeczywistemu imieniu i nazwisku z wyciągu.
    „Nazwa skrócona” to dowolny pseudonim o ograniczonej długości, nadawany
    ręcznie — może być czymkolwiek (skrótem, NIP-em itp.);
    używana jest wyłącznie jako wariant zapasowy, gdy brak nazwy pełnej.

    Pole "matched" ZAWSZE zawiera „Nazwę skróconą” — to klucz, po którym
    Saldeo szuka istniejącej karty przy imporcie faktur.
    """
    # ref: (canonical_short, full_form, short_form_fallback)
    #   canonical_short    — Nazwa skrócona; zwracana w matched (klucz importu)
    #   full_form          — formy Nazwy pełnej; główny cel porównania
    #   short_form_fallback — formy Nazwy skróconej; używana TYLKO gdy
    #                         nazwa pełna nie istnieje lub równa się skróconej
    _Form = tuple[str, frozenset, str] | None
    ref: list[tuple[str, _Form, _Form]] = []
    seen: set[str] = set()
    for c in contractors:
        canonical = (c["short"] or c["full"]).strip()
        if not canonical or canonical in seen:
            continue
        seen.add(canonical)
        short = c["short"].strip()
        full  = c["full"].strip()
        # Główna forma porównania — nazwa pełna; gdy jej brak, bierzemy skróconą
        cmp_full  = full if full else short
        # Forma zapasowa — nazwa skrócona, tylko jeśli różni się od pełnej
        cmp_short = short if short and short != cmp_full else ""
        full_form  = (_normalize(cmp_full),  _tokens(cmp_full),  _sorted_norm(cmp_full))  if cmp_full  else None
        short_form = (_normalize(cmp_short), _tokens(cmp_short), _sorted_norm(cmp_short)) if cmp_short else None
        ref.append((canonical, full_form, short_form))

    results: list[dict] = []
    for name in client_names:
        norm_name   = _normalize(name)
        name_tokens = _tokens(name)
        status, matched = "new", None

        # 1) dokładna zgodność z nazwą PEŁNĄ — najpewniejszy sygnał:
        #    „Nazwa pełna” jest zawsze najbliższa zapisowi z wyciągu
        for canonical, full_form, _ in ref:
            if full_form and norm_name == full_form[0]:
                status, matched = "exact", canonical
                break

        # 2) dokładna zgodność z nazwą SKRÓCONĄ (wariant zapasowy):
        #    działa, gdy nazwa pełna jest pusta albo skrócony pseudonim
        #    przypadkiem pokrywa się z nazwą z wyciągu
        if status == "new":
            for canonical, _, short_form in ref:
                if short_form and norm_name == short_form[0]:
                    status, matched = "exact", canonical
                    break

        # 3) inna kolejność słów — „Kowalski Jan” / „Jan Kowalski”
        #    sprawdzamy najpierw nazwę pełną, potem skróconą
        if status == "new":
            for canonical, full_form, short_form in ref:
                for form in (full_form, short_form):
                    if form and name_tokens and name_tokens == form[1]:
                        status, matched = "similar", canonical
                        break
                if status == "similar":
                    break

        # 4) nazwa pełna Saldeo jest podzbiorem tokenów z wyciągu —
        #    w Saldeo nazwa dwuczłonowa, w wyciągu trójczłonowa (z drugim
        #    imieniem). Np.: {wieczorek, sabina} ⊆ {wieczorek, sabina, dorota}.
        #    Sprawdzamy tylko nazwę pełną (≥ 2 tokeny).
        if status == "new":
            for canonical, full_form, _ in ref:
                if full_form and len(full_form[1]) >= 2 and full_form[1].issubset(name_tokens):
                    status, matched = "similar", canonical
                    break

        # 5) rozmyte podobieństwo tekstów — literówki, drobne różnice zapisu.
        #    Porównujemy z nazwą pełną; gdy jej brak — ze skróconą.
        if status == "new":
            best_ratio, best_match = 0.0, None
            for canonical, full_form, short_form in ref:
                form = full_form or short_form
                if not form:
                    continue
                ratio = SequenceMatcher(None, norm_name, form[0]).ratio()
                if ratio > best_ratio:
                    best_ratio, best_match = ratio, canonical
            if best_ratio >= SIMILARITY_THRESHOLD:
                status, matched = "similar", best_match

        # 6) inna kolejność słów + literówka jednocześnie.
        #    Sortujemy tokeny alfabetycznie — przestawienie przestaje
        #    przeszkadzać, a literówka staje się jedyną różnicą.
        #    Porównujemy z nazwą pełną; gdy jej brak — ze skróconą.
        if status == "new":
            sorted_name = _sorted_norm(name)
            best_ratio, best_match = 0.0, None
            for canonical, full_form, short_form in ref:
                form = full_form or short_form
                if not form:
                    continue
                ratio = SequenceMatcher(None, sorted_name, form[2]).ratio()
                if ratio > best_ratio:
                    best_ratio, best_match = ratio, canonical
            if best_ratio >= SIMILARITY_THRESHOLD:
                status, matched = "similar", best_match

        results.append({"name": name, "status": status, "matched": matched})

    return results
