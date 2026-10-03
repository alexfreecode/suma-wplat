"""
parser.py — Suma Wpłat: analiza wyciągu bankowego mBank
Grupuje wpłaty przychodzące od osób fizycznych według klienta.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

try:
    from tabulate import tabulate
    HAS_TABULATE = True
except ImportError:
    HAS_TABULATE = False


# ─── Stałe ────────────────────────────────────────────────────────────────────

HEADER_MARKER = "#Data operacji"

# Rdzeń do rozpoznawania przelewów przychodzących.
# Wszystkie polskie typy operacji przychodzących zawierają „przychodzący”:
#   PRZELEW ZEWNĘTRZNY/WEWNĘTRZNY PRZYCHODZĄCY
#   BLIK P2P-PRZYCHODZĄCY
#   PRZELEW EXPRESS ELIXIR PRZYCH.
#   ... i wszelkie przyszłe warianty
# Rdzeń PRZYCH obejmuje wszystkie warianty bez wymieniania konkretnych fraz.
INCOMING_ROOT = "PRZYCH"

# Słowa kluczowe firm i operacji bankowych — do wykluczenia
COMPANY_KEYWORDS = [
    "SPÓŁKA", "S.A.", "SP. Z O.O.", "SP.Z O.O.",
    "ZAKŁAD", "URZĄD", "FUNDUSZ", "TOWARZYSTWO",
    "UZNANIE NATYCH",
    # Świadczenia i przelewy urzędowe. Trafiają na konto jak każda inna
    # wpłata, a fakturą na ZUS skończyłby się problem, nie wygoda.
    # „ZUS ” ze spacją, żeby nie łapać nazwisk kończących się na „…ZUS”.
    "ZUS ", "CENTRUM OBSŁUGI", "ŚWIADCZEŃ",
]


# ─── Parsowanie danych ─────────────────────────────────────────────────────────

def find_header_row(file_path: str, encoding: str) -> int:
    """Znajduje numer wiersza z nagłówkiem tabeli (indeksowany od 0)."""
    with open(file_path, encoding=encoding, errors="replace") as f:
        for i, line in enumerate(f):
            if HEADER_MARKER in line:
                return i
    raise ValueError(
        f"Nie znaleziono wiersza nagłówka ze znacznikiem '{HEADER_MARKER}'. "
        "Upewnij się, że plik jest wyciągiem mBank."
    )


def parse_kwota(value: str) -> float:
    """Zamienia '1 300,00 PLN' → 1300.0, '-22,15 PLN' → -22.15"""
    cleaned = (
        str(value)
        .replace(" PLN", "")
        .replace("\xa0", "")   # twarda spacja
        .replace(" ", "")
        .replace(",", ".")
    )
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _clean_address(name_part: str) -> str:
    """Usuwa adres z tekstu „IMIĘ [ADRES]”, pozostawiając samo imię i nazwisko."""
    # Usuwamy sufiks " ." (format BLIK: „TIUPA YELYZAVETA .”)
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()
    # Usuwamy adres z wyraźnym prefiksem ulicy
    name_part = re.split(r"\s+(?:UL\.|AL\.|OS\.|PL\.)", name_part)[0].strip()
    # Usuwamy kod pocztowy i wszystko po nim
    name_part = re.sub(r"\s+\d{2}-\d{3}.*$", "", name_part).strip()
    # Usuwamy „NAZWA_ULICY NUMER_DOMU” bez prefiksu (na końcu tekstu)
    # Przykład: „BOHDAN HUDYMA GRANICZNA 53/74” → „BOHDAN HUDYMA”
    name_part = re.sub(r"\s+\S+\s+\d+[\w/]*\s*$", "", name_part).strip()
    return " ".join(name_part.split()).upper()


def parse_address_text(text: str, contains_name: bool = True) -> tuple[str, str, str]:
    """
    Rozkłada tekst adresu na (ulica, kod_pocztowy, miejscowość).

    `contains_name` mówi, czy w tekście przed adresem stoi jeszcze imię
    i nazwisko:

      • mBank (True) — adres siedzi w jednym worku razem z nazwiskiem
        („JAN KOWALSKI UL. KWIATOWA 5 00-001 WARSZAWA”), więc początku ulicy
        trzeba się domyślać;
      • PKO (False) — adres przychodzi osobną kolumną „Adres nadawcy”,
        więc wszystko przed kodem pocztowym JEST ulicą i nie ma czego zgadywać.
    """
    text = text.strip()

    # ── Szukamy kodu pocztowego XX-XXX ───────────────────────────────────────
    postal_match = re.search(r'\b(\d{2}-\d{3})\b', text)
    if postal_match:
        postal = postal_match.group(1)
        city   = text[postal_match.end():].strip().upper()
        # Część banków dokleja do miejscowości kod kraju („WARSZAWA PL”)
        city   = re.sub(r'\s+PL$', '', city).strip()
        before = text[:postal_match.start()].strip()

        # Ulica z prefiksem UL./AL./OS./PL.
        pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', before, re.I)
        if pref:
            street = "UL. " + pref.group(1).strip().upper()
        elif not contains_name:
            # Nie ma się czego domyślać — całość przed kodem to ulica
            street = before.upper()
        else:
            # Ulica bez prefiksu, wymieszana z nazwiskiem: bierzemy ostatnie
            # „SŁOWO NUMER”. Rozwiązanie z konieczności — przy adresach
            # z numerem mieszkania („... 25 5”) potrafi uciąć za dużo.
            no_pref = re.search(r'(\S+)\s+(\d+[\w/]*)\s*$', before)
            street  = (no_pref.group(1) + " " + no_pref.group(2)).upper() if no_pref else ""

        return (street, postal, city)

    # ── Sam prefiks bez kodu pocztowego ──────────────────────────────────────
    pref = re.search(r'(?:UL\.|AL\.|OS\.|PL\.)\s*(.+)$', text, re.I)
    if pref:
        return ("UL. " + pref.group(1).strip().upper(), "", "")

    # Bez prefiksu i bez kodu pocztowego nie ma pewności, że to w ogóle adres
    return ("", "", "")


def extract_address(desc: str) -> tuple[str, str, str]:
    """
    Wyodrębnia adres z pola „Opis operacji” wyciągu mBank.
    Zwraca (ulica, kod_pocztowy, miejscowość).
    Zwraca ("", "", "") gdy adresu brak (BLIK, Express Elixir itd.).

    Format standardowy: „IMIĘ [UL. ULICA NUMER] [KOD MIASTO], TYTUŁ ...”
    Adres znajduje się w bloku przed pierwszym przecinkiem, po imieniu.
    """
    # Express Elixir (brak przecinka) — adres niewiarygodny, pomijamy
    if "," not in desc:
        return ("", "", "")

    name_part = desc.split(",")[0].strip()
    # Usuwamy sufiks BLIK " ."
    name_part = re.sub(r"\s+\.\s*$", "", name_part).strip()

    return parse_address_text(name_part, contains_name=True)


def extract_name(desc: str) -> str:
    """
    Wyodrębnia nazwę klienta z pola Opis operacji.

    Dwa formaty:

    1. Standardowy (PRZELEW / BLIK P2P):
       „IMIĘ [ADRES], TYTUŁ   TYP_OPERACJI   NUMER_RACHUNKU”
       → imię i nazwisko brane do pierwszego przecinka

    2. Express Elixir (brak przecinka):
       „PRZELEW EXPRESS ELIXIR PRZYCH.  IMIĘ  ADRES  .  ...  PRZYCH.  RACHUNEK”
       → imię i nazwisko brane z drugiego bloku (po podwójnej spacji)
    """
    if "," not in desc:
        # Express Elixir i podobne formaty bez przecinka
        blocks = [
            b.strip()
            for b in re.split(r"\s{2,}", desc.strip())
            if b.strip() and b.strip() != "."
        ]
        # blocks[0] — typ operacji, blocks[1] — imię i nazwisko (ewent. z adresem)
        name_part = blocks[1] if len(blocks) >= 2 else (blocks[0] if blocks else "")
    else:
        # Format standardowy
        name_part = desc.split(",")[0].strip()

    result = _clean_address(name_part)
    return result if result else "NIEZNANY"


def extract_title(desc: str) -> str:
    """
    Wyodrębnia krótki tytuł płatności z pola Opis operacji.

    Format standardowy: bierze tekst po pierwszym przecinku, do podwójnej spacji.
      Przykład: „..., FV 1/03/2026  KOWALSKA...” → „FV 1/03/2026”

    Express Elixir (brak przecinka): zwraca „Express Elixir”.
    """
    if "," not in desc:
        # Express Elixir — w opisie nie ma osobnego tytułu płatności
        return "Express Elixir"

    after_comma = desc.split(",", 1)[1].strip()
    # Usuwamy sufiks „.” (BLIK)
    after_comma = re.sub(r"\s+\.\s*$", "", after_comma)
    # Bierzemy pierwszy człon do 2+ spacji
    parts = re.split(r"\s{2,}", after_comma)
    title = parts[0].strip().rstrip(".").strip()
    return title


def is_individual(desc: str) -> bool:
    """True, gdy płatność pochodzi od osoby fizycznej (nie od firmy).

    Dwa niezależne sygnały — jeśli którykolwiek zadziała, to firma:
    1. Czarna lista słów kluczowych (SPÓŁKA, S.A., URZĄD, ...)
    2. NIP — 10-cyfrowy numer podatkowy polskiej firmy.
       U osób fizycznych NIP nie występuje w opisach bankowych.
       Wzorzec: dokładnie 10 cyfr niebędących częścią dłuższej liczby
       (IBAN z 26 cyfr tu nie pasuje — tam ciąg cyfr jest nieprzerwany).
    """
    desc_upper = desc.upper()
    if any(kw in desc_upper for kw in COMPANY_KEYWORDS):
        return False
    if re.search(r'(?<!\d)\d{10}(?!\d)', desc):
        return False
    return True


def is_incoming(desc: str) -> bool:
    """True, gdy operacja jest płatnością przychodzącą.
    Sprawdza rdzeń PRZYCH, obecny w każdym polskim typie przelewu
    przychodzącego niezależnie od konkretnego sformułowania.
    """
    return INCOMING_ROOT in desc.upper()


# ─── Formatowanie wyniku ───────────────────────────────────────────────────────

def fmt_amount(value: float) -> str:
    """1300.0 → '1 300,00'"""
    formatted = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    # Python formatuje 1300.0 jako „1,300.00” — poprawiamy ręcznie
    # Bardziej niezawodny sposób:
    integer_part = int(abs(value))
    frac_part = round(abs(value) % 1 * 100)
    sign = "-" if value < 0 else ""

    # Podział na grupy po 3 cyfry
    s = str(integer_part)
    groups = []
    while s:
        groups.append(s[-3:])
        s = s[:-3]
    integer_str = " ".join(reversed(groups))  # wąska twarda spacja

    return f"{sign}{integer_str},{frac_part:02d}"


def print_separator(char: str = "─", width: int = 62) -> None:
    print("  " + char * width)


def print_header(char: str = "=", width: int = 64) -> None:
    print(char * width)


# ─── Logika główna ─────────────────────────────────────────────────────────────

BANK_MBANK = "mbank"
BANK_PKO   = "pko"

# Kolumny, po których poznajemy wyciąg PKO w formacie XLS
PKO_KOLUMNY = ("Data operacji", "Typ transakcji", "Kwota", "Nazwa nadawcy")

# Typy operacji PKO, które są wpływami, ale nie płatnościami klientów:
# wymiana walut i zwrot płatności kartą. Bez tego trafiłyby na faktury.
PKO_TYPY_NIE_KLIENT = ("WYMIANA W KANTORZE", "ZWROT W TERMINALU")

# Blokada to kwota zablokowana przez autoryzację kartą, jeszcze nie
# zaksięgowana — nie ma nawet daty operacji. Później wraca na wyciąg jako
# zwykła „Płatność kartą”, więc licząc ją razem z wydatkami policzylibyśmy
# ten sam zakup dwa razy. Pokazujemy ją osobno, zamiast po cichu wyrzucać.
PKO_TYP_BLOKADA = "BLOKADA"
KATEGORIA_BLOKADY = "Blokady kartowe"

KOMUNIKAT_PKO_CSV = (
    "To wygląda na wyciąg PKO w formacie CSV.\n\n"
    "Dla PKO program czyta pliki XLS — tam nazwa i adres nadawcy stoją\n"
    "w osobnych kolumnach, a w CSV są wymieszane w jednym polu opisu.\n\n"
    "Wyeksportuj ten sam okres jeszcze raz i w oknie wyboru formatu\n"
    "zaznacz XLS zamiast CSV."
)

# Kolumny wspólnej tabeli, którą zwraca load_transactions niezależnie od banku.
# Wszystko powyżej tej linii jest bankowe, wszystko poniżej — już nie.
KOLUMNY_WSPOLNE = ["date", "amount", "desc", "name", "title",
                   "addr_street", "addr_postal", "addr_city",
                   "account", "is_client", "is_hold", "is_company"]


# ─── Firmy oznaczone przez użytkownika ─────────────────────────────────────────
#
# Jednoosobowa firma przychodzi na wyciągu tak samo jak osoba prywatna: imię,
# nazwisko, adres. Ani mBank, ani PKO nie podają NIP-u nadawcy, więc żadna
# reguła tego nie rozpozna. Użytkownik oznacza taką firmę raz, program
# zapamiętuje ją w config.json („companies”) i od tej pory odkłada jej wpłaty
# poza raport. Rozpoznajemy po numerze konta nadawcy albo po dokładnej nazwie:
# konto łapie tę samą firmę mimo innego zapisu nazwy, nazwa — tę samą firmę
# płacącą z innego konta.

def norm_name(name: str) -> str:
    """Nazwa do porównania: wielkie litery, pojedyncze spacje."""
    return " ".join(str(name or "").upper().split())


def norm_account(account: str) -> str:
    """Numer konta do porównania: same cyfry."""
    return re.sub(r"\D", "", str(account or ""))


def _mbank_account(desc: str) -> str:
    """Konto nadawcy z opisu mBanku: 26 cyfr na końcu opisu."""
    found = re.findall(r"(?<!\d)\d{26}(?!\d)", desc)
    return found[-1] if found else ""


def mark_companies(df: pd.DataFrame, companies: list | None) -> pd.DataFrame:
    """Przenosi wpłaty firm oznaczonych przez użytkownika z klientów do firm.

    Taka wpłata przestaje być „klientem” (nie trafia do raportu, faktur
    Saldeo ani do kategorii „Klienci” kontroli kompletności), ale zostaje
    w tabeli z flagą is_company, żeby raport mógł ją pokazać osobno.
    """
    df["is_company"] = False
    if not companies or df.empty:
        return df
    names = {norm_name(c.get("name")) for c in companies if c.get("name")}
    accounts = {norm_account(c.get("account")) for c in companies}
    accounts.discard("")
    hit = df["is_client"] & (
        df["name"].apply(norm_name).isin(names)
        | df["account"].apply(norm_account).isin(accounts)
    )
    df.loc[hit, "is_company"] = True
    df.loc[hit, "is_client"] = False
    return df


def detect_bank(file_path: str) -> str:
    """Rozpoznaje bank po zawartości pliku.

    Świadomie nie ma ustawienia „wybierz bank”: ustawienie dałoby się
    przestawić błędnie i wtedy plik nie wczytywałby się bez zrozumiałego
    powodu, a użytkownik z kontami w dwóch bankach musiałby o nim pamiętać
    przy każdym wyciągu.
    """
    suffix = Path(file_path).suffix.lower()

    if suffix in (".xls", ".xlsx"):
        try:
            head = pd.read_excel(file_path, nrows=0)
        except Exception as exc:
            raise ValueError(
                f"Nie udało się otworzyć pliku Excel:\n{exc}"
            ) from exc
        columns = [str(c).strip() for c in head.columns]
        if all(c in columns for c in PKO_KOLUMNY):
            return BANK_PKO
        raise ValueError(
            "Plik Excel nie wygląda na wyciąg PKO — brakuje kolumn "
            f"{list(PKO_KOLUMNY)}.\n\nZnalezione kolumny: {columns}"
        )

    # Pliki tekstowe. Markery są czysto ASCII, więc do samego rozpoznania
    # kodowanie nie ma znaczenia i nie trzeba go zgadywać.
    with open(file_path, "rb") as f:
        start = f.read(4096).decode("latin-1")

    if HEADER_MARKER in start:
        return BANK_MBANK
    if '"Data operacji","Data waluty"' in start:
        raise ValueError(KOMUNIKAT_PKO_CSV)

    raise ValueError(
        "Nie rozpoznano formatu pliku.\n\n"
        "Program czyta:\n"
        "  • wyciąg mBank — plik CSV (Lista operacji),\n"
        "  • wyciąg PKO — plik XLS (Zestawienie operacji).\n\n"
        "Upewnij się, że plik pochodzi prosto z bankowości internetowej "
        "i nie był po drodze przerabiany w Excelu."
    )


# ─── Wczytywanie: mBank ────────────────────────────────────────────────────────

def _load_mbank(file_path: str, encoding: str) -> pd.DataFrame:
    """Wyciąg mBank (CSV) → wspólna tabela operacji."""
    header_row = find_header_row(file_path, encoding)

    raw = pd.read_csv(
        file_path,
        sep=";",
        encoding=encoding,
        skiprows=header_row,
        dtype=str,
        index_col=False,   # dane mają nadmiarową kolumnę — bez auto-indeksu
    )

    # Usuwamy zbędne spacje z nazw kolumn
    raw.columns = [c.strip() for c in raw.columns]

    required = ["#Data operacji", "#Opis operacji", "#Kwota"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            f"W pliku nie znaleziono kolumn: {missing}\n"
            f"Dostępne kolumny: {list(raw.columns)}"
        )

    # Usuwamy puste wiersze
    raw = raw.dropna(subset=["#Data operacji", "#Kwota"])

    desc = raw["#Opis operacji"].astype(str)
    df = pd.DataFrame(index=raw.index)
    df["date"]   = raw["#Data operacji"].str.strip()
    df["amount"] = raw["#Kwota"].apply(parse_kwota)
    df["desc"]   = desc
    df["account"] = desc.apply(_mbank_account)

    # U mBanku wszystko siedzi w jednym polu opisu, więc kierunek operacji
    # rozpoznajemy po typie przelewu zapisanym w tym samym tekście.
    df["is_client"] = (
        (df["amount"] > 0)
        & desc.apply(is_incoming)
        & desc.apply(is_individual)
    )
    df["is_hold"] = False              # mBank nie pokazuje blokad na wyciągu

    # Nazwę, tytuł i adres wyciągamy tylko tam, gdzie mają sens — na
    # pozostałych operacjach (opłaty, karty) dałyby przypadkowe śmieci.
    df["name"]  = ""
    df["title"] = ""
    df["addr_street"] = df["addr_postal"] = df["addr_city"] = ""
    klienci = df["is_client"]
    if klienci.any():
        df.loc[klienci, "name"]  = desc[klienci].apply(extract_name)
        df.loc[klienci, "title"] = desc[klienci].apply(extract_title)
        addr = desc[klienci].apply(extract_address)
        df.loc[klienci, "addr_street"] = addr.apply(lambda t: t[0])
        df.loc[klienci, "addr_postal"] = addr.apply(lambda t: t[1])
        df.loc[klienci, "addr_city"]   = addr.apply(lambda t: t[2])

    return df


# ─── Wczytywanie: PKO ──────────────────────────────────────────────────────────

def _pko_tekst(value) -> str:
    """Wartość komórki jako tekst. Pusta komórka Excela to NaN — a `str(NaN)`
    daje słowo „nan”, które wcześniej lądowało w opisach operacji."""
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).split())


def _pko_etykieta(row, columns, etykieta: str) -> str:
    """Wyszukuje w wierszu wartość opisaną etykietą, np. „Lokalizacja”.

    PKO wpisuje takie pary do kolumn bez nazwy, a ich kolejność zmienia się
    zależnie od rodzaju operacji — dlatego szukamy po etykiecie, nie po
    numerze kolumny.
    """
    for c in columns:
        value = row.get(c)
        if value is None or pd.isna(value):
            continue
        text = str(value).strip()
        if text.lower().startswith(etykieta.lower()):
            reszta = text[len(etykieta):].lstrip()
            if reszta.startswith(":"):
                return reszta[1:].strip()
    return ""


def _pko_tytul(value) -> str:
    """Czyści pole „Opis transakcji”: zdejmuje etykietę „Tytuł :” oraz ogon
    z numerami telefonów, który PKO dokleja przy przelewach na telefon."""
    if value is None or pd.isna(value):
        return ""
    text = " ".join(str(value).split())
    text = re.sub(r'^Tytu[łl]\s*:\s*', '', text, flags=re.I)
    text = re.sub(r'\s*OD:\s*\d+\s*DO:\s*\d+\s*$', '', text, flags=re.I)
    return text.strip()


def _pko_data(value) -> str:
    """Data operacji → tekst „RRRR-MM-DD”. W XLS przychodzi już jako data,
    ale nie polegamy na tym."""
    if value is None or pd.isna(value):
        return ""                      # blokady kartowe nie mają jeszcze daty
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _load_pko(file_path: str) -> pd.DataFrame:
    """Wyciąg PKO (XLS) → wspólna tabela operacji.

    W przeciwieństwie do mBanku nazwa i adres nadawcy przychodzą osobnymi
    kolumnami, więc nie trzeba ich wyłuskiwać z tekstu.
    """
    raw = pd.read_excel(file_path)
    raw.columns = [str(c).strip() for c in raw.columns]

    missing = [c for c in PKO_KOLUMNY if c not in raw.columns]
    if missing:
        raise ValueError(
            f"W pliku nie znaleziono kolumn: {missing}\n"
            f"Dostępne kolumny: {list(raw.columns)}"
        )

    # Filtrujemy tylko po kwocie. Po dacie NIE — bez daty przychodzą blokady
    # kartowe, a te mają trafić do raportu kontrolnego, nie zniknąć.
    raw = raw.dropna(subset=["Kwota"])
    bez_nazwy = [c for c in raw.columns if str(c).startswith("Unnamed")]

    df = pd.DataFrame(index=raw.index)
    df["date"] = raw["Data operacji"].apply(_pko_data)
    # Kwota przychodzi liczbą ze znakiem; parse_kwota tylko na wypadek,
    # gdyby Excel podał ją jako tekst
    df["amount"] = raw["Kwota"].apply(
        lambda v: float(v) if isinstance(v, (int, float)) else parse_kwota(str(v))
    )
    df["name"]  = raw["Nazwa nadawcy"].fillna("").astype(str).str.strip()
    df["account"] = (raw["Rachunek nadawcy"].apply(_pko_tekst).apply(norm_account)
                     if "Rachunek nadawcy" in raw.columns else "")
    df["title"] = raw["Opis transakcji"].apply(_pko_tytul) \
        if "Opis transakcji" in raw.columns else ""

    typ = raw["Typ transakcji"].fillna("").astype(str).str.strip()

    # Kierunek bierzemy ze ZNAKU KWOTY, nie z typu operacji. U PKO ten sam typ
    # („Przelew z rachunku”) bywa i wpływem, i wydatkiem, a rdzeń „PRZYCH”
    # pojawia się tylko w części nazw — na wyciągu testowym złapałby
    # 2 wpływy z 13.
    nie_klient = typ.str.upper().apply(
        lambda t: any(t.startswith(x) for x in PKO_TYPY_NIE_KLIENT)
    )
    df["is_hold"] = typ.str.upper().str.startswith(PKO_TYP_BLOKADA)
    df["is_client"] = (
        (df["amount"] > 0)
        & (df["name"] != "")
        & df["name"].apply(is_individual)
        & ~nie_klient
        & ~df["is_hold"]
    )

    # Adres tylko dla klientów — u pozostałych operacji kolumna dotyczy
    # odbiorcy albo jest pusta
    df["addr_street"] = df["addr_postal"] = df["addr_city"] = ""
    if "Adres nadawcy" in raw.columns:
        klienci = df["is_client"]
        if klienci.any():
            addr = raw.loc[klienci, "Adres nadawcy"].fillna("").astype(str).apply(
                lambda t: parse_address_text(t, contains_name=False)
            )
            df.loc[klienci, "addr_street"] = addr.apply(lambda t: t[0])
            df.loc[klienci, "addr_postal"] = addr.apply(lambda t: t[1])
            df.loc[klienci, "addr_city"]   = addr.apply(lambda t: t[2])

    # Opis do raportu kontrolnego: powtarzalny, żeby jednakowe operacje
    # (np. zakupy w tym samym sklepie) zwijały się do jednego wiersza.
    # Tytuł pomijamy, gdy jest samym numerem referencyjnym — inaczej każdy
    # wiersz byłby inny i grupowanie nic by nie dało.
    def opis(idx) -> str:
        row = raw.loc[idx]
        kontrahent = (_pko_tekst(row.get("Nazwa nadawcy"))
                      or _pko_tekst(row.get("Nazwa odbiorcy"))
                      or _pko_etykieta(row, bez_nazwy, "Lokalizacja"))
        tytul = df.at[idx, "title"]
        if re.fullmatch(r'[\d\s*]*', tytul or ""):
            tytul = ""
        czesci = [typ.at[idx], kontrahent, tytul]
        return " · ".join(c for c in czesci if c)

    df["desc"] = [opis(i) for i in raw.index]

    return df


# ─── Wczytywanie: wspólne wejście ──────────────────────────────────────────────

def load_transactions(file_path: str, encoding: str,
                      companies: list | None = None) -> pd.DataFrame:
    """Wczytuje wyciąg dowolnego obsługiwanego banku i zwraca tabelę
    o stałych kolumnach (KOLUMNY_WSPOLNE). Cała reszta programu — raporty,
    kontrola kompletności, faktury Saldeo — pracuje już tylko na niej
    i o bankach nic nie wie.

    companies — firmy oznaczone przez użytkownika (z config.json); ich
    wpłaty dostają is_company zamiast is_client."""
    bank = detect_bank(file_path)
    if bank == BANK_PKO:
        df = _load_pko(file_path)
    else:
        df = _load_mbank(file_path, encoding)
    df = mark_companies(df, companies)
    return df[KOLUMNY_WSPOLNE]


KOLUMNY_KLIENTA = ["date", "name", "amount", "title",
                   "addr_street", "addr_postal", "addr_city", "account"]


def analyze_split(file_path: str, encoding: str,
                  companies: list | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Jak analyze(), ale zwraca też wpłaty od firm oznaczonych przez
    użytkownika: (klienci, firmy)."""
    df = load_transactions(file_path, encoding, companies)
    clients = df[df["is_client"]][KOLUMNY_KLIENTA].reset_index(drop=True)
    firms = df[df["is_company"]][KOLUMNY_KLIENTA].reset_index(drop=True)
    return clients, firms


def analyze(file_path: str, encoding: str,
            companies: list | None = None) -> pd.DataFrame:
    """
    Wczytuje wyciąg, filtruje płatności od osób fizycznych,
    zwraca DataFrame z kolumnami: date, name, amount, title, adres i konto.
    """
    return analyze_split(file_path, encoding, companies)[0]


def categorize_transaction(amount: float, is_client: bool,
                          is_hold: bool = False) -> str:
    """
    Przypisuje operację do jednej z trzech kategorii:
      "Klienci"          — wpłata od osoby fizycznej (to, co trafia na faktury)
      "Pozostałe wpływy" — każdy inny wpływ (od firm, zwroty, odsetki itp.)
      "Wydatki"          — każdy wydatek

    Rozpoznanie klienta robi już wczytywanie wyciągu — każdy bank po swojemu —
    więc tutaj przychodzi gotową flagą.
    """
    if is_hold:
        return KATEGORIA_BLOKADY
    if amount > 0:
        return "Klienci" if is_client else "Pozostałe wpływy"
    return "Wydatki"


def _collapse_desc(desc: str) -> str:
    """Zwija wszystkie ciągi białych znaków do pojedynczej spacji."""
    return " ".join(str(desc).split())


def _shorten(text: str, max_len: int = 70) -> str:
    if len(text) > max_len:
        return text[: max_len - 1] + "…"
    return text


def print_reconciliation_report(file_path: str, encoding: str, output_lines: list,
                                companies: list | None = None) -> None:
    """
    Generuje osobny raport „Kontrola kompletności wyciągu” — dzieli
    WSZYSTKIE operacje wyciągu na trzy kategorie (Klienci / Pozostałe
    wpływy / Wydatki), aby użytkownik mógł sprawdzić, że nic nie zginęło
    i suma łączna zgadza się z wyciągiem.

    Wewnątrz każdej kategorii operacje grupowane są po identycznym tekście
    „Opis operacji” — powtarzające się opisy (np. prowizje bankowe)
    zwijane są do jednego wiersza z liczbą operacji i sumą.
    """

    def out(line: str = "") -> None:
        try:
            print(line)
        except Exception:
            pass
        output_lines.append(line)

    df = load_transactions(file_path, encoding, companies)
    df["category"] = df.apply(
        lambda r: categorize_transaction(r["amount"], bool(r["is_client"]),
                                         bool(r["is_hold"])),
        axis=1,
    )
    # Blokady trzymamy osobno: nie są jeszcze operacją, więc nie wchodzą
    # do sum kontrolnych, ale muszą być widoczne
    blokady = df[df["is_hold"]]
    df_ops  = df[~df["is_hold"]]

    out("═" * 64)
    out("  KONTROLA KOMPLETNOŚCI WYCIĄGU")
    out("═" * 64)
    out()
    incoming = df_ops[df_ops["amount"] > 0]
    out(f"  Wszystkich operacji w wyciągu: {len(df_ops)}")
    out(f"  Suma wpływów:                  {fmt_amount(incoming['amount'].sum())} PLN"
        f"  ({len(incoming)} operacji)")
    out(f"  Suma wszystkich operacji:      {fmt_amount(df_ops['amount'].sum())} PLN")
    if not blokady.empty:
        out(f"  Blokady kartowe (poza sumą):   {fmt_amount(blokady['amount'].sum())} PLN"
            f"  ({len(blokady)} operacji)")

    kategorie = ["Klienci", "Pozostałe wpływy", "Wydatki"]
    if not blokady.empty:
        kategorie.append(KATEGORIA_BLOKADY)

    for category in kategorie:
        sub = df[df["category"] == category]

        out()
        out("─" * 64)
        out(f"  {category.upper()}  —  {len(sub)} operacji,"
            f"  {fmt_amount(sub['amount'].sum())} PLN")
        out("─" * 64)

        if sub.empty:
            out("  (brak operacji)")
            continue

        sub = sub.copy()
        sub["_desc"] = sub["desc"].apply(_collapse_desc)

        grouped = (
            sub.groupby("_desc", sort=False)
            .agg(n=("amount", "count"),
                 s=("amount", "sum"),
                 d=("date", "first"))
            .reset_index()
        )
        grouped = grouped.reindex(
            grouped["s"].abs().sort_values(ascending=False).index
        )

        for _, row in grouped.iterrows():
            label = _shorten(row["_desc"])
            amount_str = fmt_amount(row["s"]) + " PLN"
            if row["n"] == 1:
                out(f"  {str(row['d']).strip():<12}  {amount_str:>16}   {label}")
            else:
                out(f"  {'':<12}  {amount_str:>16}   {label}  (× {int(row['n'])})")

    out()
    out("═" * 64)
    out(f"  SUMA KONTROLNA (wszystkie operacje):  "
        f"{fmt_amount(df_ops['amount'].sum())} PLN")
    if not blokady.empty:
        out("  (bez blokad kartowych — te nie są jeszcze zaksięgowane")
        out("   i wrócą na wyciąg jako zwykłe płatności kartą)")
    out("═" * 64)


def print_report(df: pd.DataFrame, output_lines: list,
                 firms: pd.DataFrame | None = None) -> None:
    """Generuje szczegółowy raport per klient + tabelę zbiorczą.

    firms — wpłaty od firm oznaczonych przez użytkownika; pokazujemy je na
    końcu osobnym blokiem, żeby nie zniknęły z oczu."""

    def out(line: str = "") -> None:
        # print() pisze do sys.stdout — potrzebne tylko w trybie CLI.
        # W trybie GUI (a tym bardziej w zbudowanym „--windowed” .exe, gdzie
        # sys.stdout jest wręcz równy None) ten wydruk nie jest potrzebny —
        # raport i tak zbierany jest w output_lines i trafia do pola
        # tekstowego/pliku. Przy tym wydruk może SIĘ WYSYPAĆ: jeśli kodowanie
        # konsoli nie „rozumie” polskich znaków (np. cp1252/„charmap” nie
        # zawiera Ą Ć Ę Ł Ń Ś Ź Ż), print() rzuca UnicodeEncodeError i cały
        # raport urywa się błędem — dlatego wyciszamy wszelkie błędy druku.
        try:
            print(line)
        except Exception:
            pass
        output_lines.append(line)

    if df.empty:
        out("Nie znaleziono płatności od osób fizycznych.")
        _print_firms(firms, out)
        return

    # Ustalamy datę pierwszej wpłaty każdego klienta
    first_payment = df.groupby("name")["date"].min().rename("first_date")
    df_sorted = df.join(first_payment, on="name")

    # Sortujemy: klienci — wg daty pierwszej wpłaty, w ramach klienta — wg daty
    df_sorted = df_sorted.sort_values(["first_date", "name", "date"])

    # ── Bloki per klient (sort=False — kolejność już ustalona wyżej) ──
    for name, group in df_sorted.groupby("name", sort=False):
        group = group.sort_values("date")
        total = group["amount"].sum()
        count = len(group)

        out()
        out("=" * 64)
        out(f"  KLIENT: {name}")
        out("-" * 64)

        # Tabela transakcji
        rows = []
        for _, row in group.iterrows():
            rows.append([row["date"], fmt_amount(row["amount"]) + " PLN", row["title"]])

        if HAS_TABULATE:
            table = tabulate(
                rows,
                headers=["Data", "Kwota", "Opis"],
                tablefmt="plain",
            )
            for line in table.splitlines():
                out("  " + line)
        else:
            out(f"  {'Data':<14}  {'Kwota':<16}  Opis")
            out("  " + "─" * 60)
            for r in rows:
                out(f"  {r[0]:<14}  {r[1]:<16}  {r[2]}")

        out("  " + "─" * 44)
        out(f"  RAZEM:  {fmt_amount(total) + ' PLN':>16}  ({count} transakcji)")

    # ── Tabela zbiorcza ──
    out()
    out()
    out("═" * 64)
    out("  TABELA ZBIORCZA")
    out("═" * 64)

    summary = (
        df.groupby("name")
        .agg(transakcji=("amount", "count"), razem=("amount", "sum"))
        .sort_values("razem", ascending=False)
        .reset_index()
    )
    summary.rename(columns={"name": "Klient"}, inplace=True)

    summary_rows = [
        [row["Klient"], int(row["transakcji"]), fmt_amount(row["razem"]) + " PLN"]
        for _, row in summary.iterrows()
    ]

    if HAS_TABULATE:
        table = tabulate(
            summary_rows,
            headers=["Klient", "Transakcji", "Kwota"],
            tablefmt="plain",
            colalign=("left", "right", "right"),
        )
        for line in table.splitlines():
            out("  " + line)
    else:
        out(f"  {'Klient':<35}  {'Transakcji':>10}  {'Kwota':>16}")
        out("  " + "─" * 65)
        for r in summary_rows:
            out(f"  {r[0]:<35}  {r[1]:>10}  {r[2]:>16}")

    out("  " + "═" * 65)
    total_all = df["amount"].sum()
    total_count = len(df)
    clients_count = df["name"].nunique()
    out(f"  {'RAZEM':<35}  {total_count:>10}  {fmt_amount(total_all) + ' PLN':>16}")
    out()
    out(f"  Unikalnych klientów: {clients_count}")
    out("═" * 64)
    _print_firms(firms, out)


def _print_firms(firms: pd.DataFrame | None, out) -> None:
    """Blok „Firmy (poza raportem)”: kto i ile, bez szczegółów wpłat."""
    if firms is None or firms.empty:
        return
    summary = (
        firms.groupby("name")
        .agg(n=("amount", "count"), s=("amount", "sum"))
        .sort_values("s", ascending=False)
        .reset_index()
    )
    out()
    out()
    out("═" * 64)
    out("  FIRMY (POZA RAPORTEM)")
    out("═" * 64)
    for _, row in summary.iterrows():
        out(f"  {row['name']:<35}  {int(row['n']):>10}  "
            f"{fmt_amount(row['s']) + ' PLN':>16}")
    out()
    out("  Oznaczone jako firma. Zmienisz to w menu „Firmy”.")
    out("═" * 64)


# ─── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Suma Wpłat — wpłaty od klientów indywidualnych z wyciągu mBank",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Przykłady:
  python parser.py --file lista_operacji_260301_260331.csv
  python parser.py --file statement.csv --output report.txt
  python parser.py --file statement.csv --encoding cp1250
""",
    )
    parser.add_argument(
        "--file", "-f",
        required=True,
        help="Ścieżka do pliku CSV z wyciągiem mBank",
    )
    parser.add_argument(
        "--encoding", "-e",
        default="utf-8-sig",
        help="Kodowanie pliku (domyślnie: utf-8-sig). "
             "Dla starszych wyciągów spróbuj cp1250.",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Zapisz raport do pliku tekstowego (opcjonalnie)",
    )

    args = parser.parse_args()

    # Sprawdzamy plik
    if not Path(args.file).exists():
        print(f"Błąd: plik nie znaleziony: {args.file}", file=sys.stderr)
        sys.exit(1)

    # Analiza
    try:
        df = analyze(args.file, args.encoding)
    except Exception as exc:
        print(f"Błąd podczas odczytu pliku: {exc}", file=sys.stderr)
        sys.exit(1)

    # Wynik
    output_lines: list = []
    print_report(df, output_lines)

    # Zapis do pliku
    if args.output:
        try:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write("\n".join(output_lines))
            print(f"\nRaport zapisany: {args.output}")
        except OSError as exc:
            print(f"Nie udało się zapisać pliku: {exc}", file=sys.stderr)


if __name__ == "__main__":
    # Przełączamy stdout/stderr na UTF-8 dla terminala Windows
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    main()
