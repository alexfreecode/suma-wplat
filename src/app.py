"""
app.py — nakładka GUI programu „Suma Wpłat”
Uruchomienie: python app.py  (lub pythonw app.py — bez konsoli)
"""

import json
import os
import re
import shutil
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from datetime import date, datetime
from pathlib import Path
from tkinter import (filedialog, font as tkfont, messagebox, scrolledtext,
                     simpledialog, ttk)

# Importujemy logikę analizy i generowania faktur
from difflib import SequenceMatcher

from parser import analyze, print_report, print_reconciliation_report
from saldeo_export import (VAT_BASIS, generate_saldeo_xlsx,
                           _parse_date)
from contractor_check import (load_saldeo_contractors, check_clients,
                              _normalize as _normalize_txt, SIMILARITY_THRESHOLD)

# ─── Konfiguracja (zachowywana między uruchomieniami) ─────────────────────────


# Nazwa programu — używana w tytułach okien i jako nazwa folderu ustawień
APP_NAME        = "Suma Wpłat"
# Poprzednia nazwa; potrzebna wyłącznie do przeniesienia starych ustawień
LEGACY_APP_NAME = "mBank Analyzer"


def _migrate_legacy_config(base: Path, new_dir: Path) -> None:
    """Przenosi ustawienia z folderu poprzedniej nazwy programu.

    Program nazywał się wcześniej „mBank Analyzer” i trzymał config.json
    w %APPDATA%\\mBank Analyzer. Po zmianie nazwy folder jest inny — bez tego
    kroku dotychczasowi użytkownicy straciliby dane sprzedawcy, słownik usług
    i ścieżkę do bazy kontrahentów, a program przywitałby ich oknem
    „pierwsze uruchomienie”.

    Kopiujemy, a nie przenosimy: gdyby ktoś wrócił do starej wersji,
    jego ustawienia nadal tam będą.
    """
    new_cfg = new_dir / "config.json"
    if new_cfg.exists():
        return                                   # nowe ustawienia już są
    old_cfg = base / LEGACY_APP_NAME / "config.json"
    if not old_cfg.is_file():
        return                                   # nie ma czego przenosić
    try:
        shutil.copy2(old_cfg, new_cfg)
    except Exception:
        pass                                     # brak ustawień nie blokuje startu


def _config_dir() -> Path:
    """Folder dla config.json.

    Przy uruchomieniu zbudowanego .exe (PyInstaller --onefile) __file__
    wskazuje na folder tymczasowy, usuwany po zamknięciu — nie można tam
    przechowywać ustawień. Dlatego dla „zamrożonej” aplikacji używamy
    folderu profilu użytkownika (%APPDATA%), a przy uruchomieniu zwykłego
    .py — folderu obok skryptu (wygodne przy programowaniu).
    """
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA", Path.home()))
        d = base / APP_NAME
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        _migrate_legacy_config(base, d)
        return d
    return Path(__file__).parent


_CONFIG_PATH = _config_dir() / "config.json"

def _load_config() -> dict:
    try:
        return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_config(data: dict) -> None:
    try:
        _CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                encoding="utf-8")
    except Exception:
        pass


# ─── Stałe ────────────────────────────────────────────────────────────────────

# Wersja programu. JEDYNE miejsce, w którym się ją podaje — instalator
# dostaje ją z build_all.ps1, żeby numery nie mogły się rozjechać.
# Trzy człony (major.minor.patch): zostawia miejsce na poprawkę bez
# udawania, że to nowa funkcjonalność — a poprawka jest prawdopodobna,
# bo obsługę PKO pisaliśmy bez dostępu do prawdziwych wpłat klientów.
APP_VERSION = "1.3.0"

# Numer wersji w pasku tytułu: użytkownik pisząc „nie działa” zwykle nie wie,
# co ma zainstalowane, a tutaj widzi to bez szukania. Sama nazwa nie mówi,
# czego program dotyczy, więc zostaje też krótki opis — bez nazwy banku,
# bo obsługiwane są dwa.
TITLE     = f"{APP_NAME} {APP_VERSION} — wpłaty od klientów z wyciągu bankowego"

# Adresy używane w oknie „O programie”. Repozytorium może kiedyś zmienić
# nazwę; GitHub trzyma wtedy przekierowanie, ale lepiej poprawić tutaj.
GITHUB_REPO   = "alexfreecode/suma-wplat"
RELEASES_API  = f"https://api.github.com/repos/{GITHUB_REPO}/releases"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"
DONATE_URL    = "https://revolut.me/oleksa49b"

# Strona programu — czysty adres, BEZ parametrów kampanii.
# Były tu wcześniej znaczniki Matomo (skąd przyszedł odwiedzający i z jakiej
# wersji), ale to zły interes: program obiecuje, że niczego nie wysyła,
# a użytkownik zobaczyłby w pasku przeglądarki adres z doklejonym śledzeniem.
# Wygląda jak złamanie obietnicy, choć technicznie nią nie jest. Przy tym
# sama liczba nic nie zmienia w decyzjach: kto ma program, nie jest nowym
# odwiedzającym.
WEBSITE_URL = "https://sumawplat.pl"

FONT_MONO = ("Courier New", 10)
FONT_UI   = ("Segoe UI", 10)
FONT_BTN  = ("Segoe UI", 10)
WIN_W, WIN_H = 860, 680
PAD = 10

# Podpowiedź pokazywana w pustym polu wyniku, dopóki nie ma analizy.
# Puste pole wygląda dokładnie tak samo jak wynik analizy, która nikogo nie
# znalazła — bez tego tekstu program przy pierwszym uruchomieniu sprawia
# wrażenie, że już coś policzył.
HINT_TITLE = "Podpowiedź"
# Tekst jest krótki nie bez powodu: im mniej wierszy, tym większą czcionką
# mieści się w polu wyniku, a podpowiedź ma być widoczna z drugiego końca
# biurka. Wyleciało „co potem” i odesłanie do pomocy — menu i tak jest na
# wierzchu. Zostały kroki i ostrzeżenie o formacie PKO, czyli jedyna rzecz,
# która realnie ratuje użytkownika przed nieudanym pierwszym podejściem.
# Bez pustych wierszy między krokami: każdy usunięty wiersz to większa
# czcionka. Na tym samym polu ten układ daje Arial 21 zamiast 16 i wypełnia
# szerokość okna, a numery kroków i tak trzymają strukturę.
HINT_TEXT = """
1.  Pobierz z banku listę operacji za miesiąc
    mBank: Historia → Eksportuj listę → format CSV
    PKO BP: Historia → Zrealizowane → „Pobierz zestawienie”
        → format XLS, nie CSV
2.  Kliknij „Wybierz” u góry i wskaż pobrany plik
3.  Kliknij „▶ Uruchom analizę”

Tutaj pojawi się wynik: kto, kiedy i ile wpłacił.
"""


# ─── Motywy raportów (ciemne / jasne tło) ─────────────────────────────────────

REPORT_THEMES = {
    "dark": {
        "bg": "#1e1e1e", "fg": "#d4d4d4", "insert": "white",
        "naglowek": "#ffffff", "klienci": "#5ecf5e", "wplywy": "#58b6f0",
        "wydatki": "#f0705e", "suma": "#ffd75e", "linia": "#5a5a5a",
        # Podpowiedź w pustym polu — przygaszona, żeby nie udawała wyniku
        "podpowiedz": "#8a8a8a", "podpowiedz_akcent": "#c8c8c8",
    },
    "light": {
        "bg": "#ffffff", "fg": "#1e1e1e", "insert": "black",
        "naglowek": "#000000", "klienci": "#187a18", "wplywy": "#0b62a4",
        "wydatki": "#c0392b", "suma": "#9a6b00", "linia": "#9a9a9a",
        "podpowiedz": "#8a8a8a", "podpowiedz_akcent": "#333333",
    },
}


def _report_theme() -> dict:
    """Aktualny motyw raportów wg config.json (domyślnie ciemny)."""
    name = _load_config().get("report_theme", "dark")
    return REPORT_THEMES.get(name, REPORT_THEMES["dark"])


def _koloruj_raport(text: tk.Text, report_text: str, theme: dict) -> None:
    """Kolorowe wyróżnienie nagłówków i sum — wspólne dla raportu głównego
    i okna „Kontrola kompletności". Sekcje mają własne kolory, a sumy
    od razu rzucają się w oczy."""
    bold = (FONT_MONO[0], FONT_MONO[1], "bold")
    text.tag_configure("naglowek", foreground=theme["naglowek"], font=bold)
    text.tag_configure("klienci",  foreground=theme["klienci"],  font=bold)
    text.tag_configure("wplywy",   foreground=theme["wplywy"],   font=bold)
    text.tag_configure("wydatki",  foreground=theme["wydatki"],  font=bold)
    text.tag_configure("suma",     foreground=theme["suma"],     font=bold)
    text.tag_configure("linia",    foreground=theme["linia"])

    for i, line in enumerate(report_text.splitlines(), start=1):
        s = line.strip()
        if not s:
            continue
        if s.startswith(("KONTROLA KOMPLETNOŚCI", "TABELA ZBIORCZA",
                         "Unikalnych klientów")):
            tag = "naglowek"
        elif s.startswith(("KLIENT:", "KLIENCI")):
            tag = "klienci"
        elif s.startswith("POZOSTAŁE WPŁYWY"):
            tag = "wplywy"
        elif s.startswith("WYDATKI"):
            tag = "wydatki"
        elif s.startswith(("SUMA KONTROLNA", "RAZEM")):
            tag = "suma"
        elif set(s) <= {"═", "─", "=", "-"}:
            tag = "linia"
        else:
            continue
        text.tag_add(tag, f"{i}.0", f"{i}.end")


def secondary_btn(parent, text: str, command, width: int | None = None) -> tk.Button:
    """Jednolity styl WSZYSTKICH przycisków w programie: płaski, szary,
    kursor „rączka”.

    Wcześniej akcje główne były kolorowe (zielona „Generuj”, niebieska
    „Uruchom analizę”, fioletowa „Faktury Saldeo”). Kolor nie niósł żadnej
    informacji, a okna wyglądały jak zlepek. Akcję główną wyróżnia teraz
    piktogram (▶, ✓) i miejsce w oknie, nie kolor."""
    kw = dict(
        text=text, font=FONT_BTN, command=command,
        relief=tk.FLAT, bg="#e4e4e4", activebackground="#d0d0d0",
        cursor="hand2", padx=10, pady=3, bd=0,
        highlightthickness=1, highlightbackground="#c0c0c0",
    )
    if width is not None:
        kw["width"] = width
    return tk.Button(parent, **kw)


def _services_catalog(cfg: dict) -> list[str]:
    """Słownik usług do wyboru na liście klientów: zapisane usługi plus
    usługa główna z danych sprzedawcy (żeby zawsze była na liście)."""
    catalog = list(cfg.get("services_catalog", []))
    main = (cfg.get("service_name") or "").strip()
    if main and main not in catalog:
        catalog.insert(0, main)
    return catalog


# Próg podobieństwa dla nazw usług. Celowo wyższy niż przy kontrahentach
# (0,84): nazwy usług są dłuższe i różnią się często jednym kluczowym słowem.
# Pomiar na realnych przykładach: warianty tej samej usługi dają 0,91–1,00
# („Konsultacje" / „Konsultacja" = 0,909), a różne usługi najwyżej 0,85
# („Korepetycje z niemieckiego" / „z angielskiego" = 0,846).
SERVICE_SIMILARITY_THRESHOLD = 0.90


def _similar_service(name: str, catalog: list[str]) -> str | None:
    """Szuka w słowniku usługi podobnej do wpisanej — chroni przed mnożeniem
    wariantów tej samej usługi („Konsultacja" / „konsultacje" / „Konsultacia").

    Korzysta z tej samej normalizacji, co porównywanie kontrahentów
    (contractor_check), ale z własnym, ostrzejszym progiem.
    """
    norm = _normalize_txt(name)
    best, best_ratio = None, 0.0
    for item in catalog:
        if _normalize_txt(item) == norm:
            return item                       # to samo, tylko inna pisownia
        ratio = SequenceMatcher(None, norm, _normalize_txt(item)).ratio()
        if ratio > best_ratio:
            best, best_ratio = item, ratio
    return best if best_ratio >= SERVICE_SIMILARITY_THRESHOLD else None


def _klienci(n: int) -> str:
    """Polska odmiana rzeczownika po liczbie — dopełniacz pasuje do „u N
    klientów" przy każdej liczbie poza jedynką."""
    return "1 klienta" if n == 1 else f"{n} klientów"


class AutocompleteCombobox(ttk.Combobox):
    """Pole usługi z podpowiedziami: w miarę pisania pod polem pojawia się
    lista pozycji słownika pasujących do wpisanego tekstu (najpierw te
    zaczynające się od wpisanego tekstu, potem pozostałe zawierające go).

    Świadomie NIE dopisuje tekstu za użytkownika. Autouzupełnianie „w locie"
    potrafi po cichu podmienić wpisywaną nazwę — piszesz „Sesja indywidualna",
    a pole samo robi z tego „Sesja grupowa" i łatwo tego nie zauważyć.

    Podpowiedzi pokazujemy we własnym okienku, a nie w rozwijanej liście
    Comboboxa: ta po otwarciu przechwytuje klawiaturę i dalsze pisanie
    trafiałoby do niej zamiast do pola.
    """

    # Klawisze, po których nie ma sensu przeliczać podpowiedzi — obsługuje je
    # nawigacja po liście albo sam Combobox
    _NAV_KEYS = frozenset((
        "Up", "Down", "Left", "Right", "Return", "Escape", "Tab",
        "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
    ))

    def __init__(self, master, catalog: list[str], **kw):
        super().__init__(master, **kw)
        self._catalog = list(catalog)
        self["values"] = self._catalog
        self._popup: tk.Toplevel | None = None
        self._listbox: tk.Listbox | None = None

        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<FocusOut>",   self._on_focus_out)
        self.bind("<Escape>",     lambda e: self.hide_popup())
        self.bind("<Down>",       self._on_down)
        self.bind("<Up>",         self._on_up)
        self.bind("<Return>",     self._on_return)
        self.bind("<<ComboboxSelected>>", lambda e: self.hide_popup())

    # ── Podpowiedzi ────────────────────────────────────────────────────────

    def _on_key_release(self, event):
        if event.keysym in self._NAV_KEYS:
            return
        typed = self.get().strip().lower()
        if not typed:
            self.hide_popup()
            return
        starts  = [s for s in self._catalog if s.lower().startswith(typed)]
        rest    = [s for s in self._catalog
                   if typed in s.lower() and s not in starts]
        matches = starts + rest
        # Gdy wpisano dokładnie to, co jest w słowniku, podpowiedź nic nie wnosi
        if not matches or (len(matches) == 1 and matches[0].lower() == typed):
            self.hide_popup()
            return
        self._show_popup(matches)

    def _show_popup(self, matches: list[str]):
        if self._popup is None:
            self._popup = tk.Toplevel(self)
            self._popup.overrideredirect(True)     # bez ramki i paska tytułu
            self._popup.attributes("-topmost", True)
            self._listbox = tk.Listbox(
                self._popup, font=FONT_UI, activestyle="none", bd=0,
                highlightthickness=1, highlightbackground="#7a7a7a",
                selectbackground="#0078d7", selectforeground="white",
                exportselection=False,
            )
            self._listbox.pack(fill=tk.BOTH, expand=True)
            self._listbox.bind("<Button-1>", self._on_click)

        lb = self._listbox
        lb.delete(0, tk.END)
        for item in matches:
            lb.insert(tk.END, item)
        lb.configure(height=min(len(matches), 6))

        self.update_idletasks()
        self._popup.geometry(
            f"{self.winfo_width()}x{lb.winfo_reqheight()}"
            f"+{self.winfo_rootx()}+{self.winfo_rooty() + self.winfo_height()}"
        )
        self._popup.deiconify()
        self._popup.lift()

    def hide_popup(self):
        try:
            if self._popup is not None:
                self._popup.withdraw()
        except tk.TclError:
            pass                                   # okno już zniknęło

    def _popup_visible(self) -> bool:
        try:
            return self._popup is not None and self._popup.winfo_ismapped()
        except tk.TclError:
            return False

    def _on_focus_out(self, event):
        # Kliknięcie w podpowiedź także zabiera fokus polu, więc chowamy listę
        # z drobnym opóźnieniem — inaczej zniknęłaby przed obsługą kliknięcia
        self.after(150, self.hide_popup)

    def _accept(self, value: str):
        self.set(value)
        self.icursor(tk.END)
        self.hide_popup()

    # ── Nawigacja klawiaturą ───────────────────────────────────────────────

    def _on_click(self, event):
        self._accept(self._listbox.get(self._listbox.nearest(event.y)))
        return "break"

    def _on_down(self, event):
        if not self._popup_visible():
            return                                 # zwykłe rozwinięcie listy
        return self._move(+1)

    def _on_up(self, event):
        if not self._popup_visible():
            return
        return self._move(-1)

    def _move(self, step: int):
        lb  = self._listbox
        cur = lb.curselection()
        idx = 0 if not cur else max(0, min(cur[0] + step, lb.size() - 1))
        lb.selection_clear(0, tk.END)
        lb.selection_set(idx)
        lb.see(idx)
        return "break"

    def _on_return(self, event):
        if not self._popup_visible():
            return
        cur = self._listbox.curselection()
        if cur:
            self._accept(self._listbox.get(cur[0]))
        else:
            self.hide_popup()
        return "break"


def _default_output_dir(source_path: str = "") -> Path:
    """Folder proponowany dla pliku z fakturami: ten sam, w którym leży
    wczytany wyciąg (zwykle „Pobrane”). Gdy ścieżki wyciągu nie znamy —
    „Dokumenty”, a w ostateczności folder domowy.

    Katalog roboczy programu odpada: przy starcie ze skrótu jest nim folder
    instalacji w AppData, czyli miejsce, w którym użytkownik pliku nie znajdzie.
    """
    if source_path:
        parent = Path(source_path).expanduser().parent
        if parent.is_dir():
            return parent
    for candidate in (Path.home() / "Documents", Path.home() / "Dokumenty"):
        if candidate.is_dir():
            return candidate
    return Path.home()


def _default_output_name(df) -> str:
    """Nazwa proponowana dla pliku z fakturami — z miesiącem wpłat, żeby
    kolejne miesiące nie nadpisywały się nawzajem."""
    try:
        # Kolumna „date” to tekst („2026-04-05” albo „05.04.2026”), więc
        # korzystamy z tego samego parsera, co generator faktur
        months = sorted({_parse_date(d).strftime("%Y-%m") for d in df["date"]})
    except Exception:
        months = []
    if len(months) == 1:
        return f"faktury_saldeo_{months[0]}.xlsx"
    if len(months) > 1:
        return f"faktury_saldeo_{months[0]}_{months[-1]}.xlsx"
    return "faktury_saldeo.xlsx"


# Do ilu członów uzupełniamy numer wersji przy porównywaniu
_VERSION_PARTS = 4


def _parse_version(text: str) -> tuple[int, ...]:
    """„v1.3” albo „1.3.1” → (1, 3, 0, 0) / (1, 3, 1, 0).

    Człony nieliczbowe pomijamy, żeby nietypowy tag nie wywalił porównania.

    Krótsze numery dopełniamy zerami, bo inaczej „1.3” wypadłoby MNIEJSZE
    od „1.3.0” (krótsza krotka sortuje się wcześniej) i program pokazywałby
    aktualizację do tej samej wersji.
    """
    czesci = []
    for kawalek in str(text).strip().lstrip("vV").split("."):
        cyfry = "".join(c for c in kawalek if c.isdigit())
        if not cyfry:
            break
        czesci.append(int(cyfry))
    czesci = czesci[:_VERSION_PARTS]
    return tuple(czesci + [0] * (_VERSION_PARTS - len(czesci))) if czesci else ()


def _fetch_releases(timeout: int = 8) -> list[dict]:
    """Pobiera listę wydań z GitHuba.

    Wywoływane WYŁĄCZNIE po kliknięciu użytkownika w oknie „O programie”.
    Program nie odzywa się do sieci sam z siebie — to świadoma decyzja:
    czyta wyciąg bankowy i nie wysyła niczego, dopóki nikt o to nie poprosi.
    Wysyłamy samo zapytanie GET, bez żadnych danych o użytkowniku.
    """
    request = urllib.request.Request(
        RELEASES_API + "?per_page=20",
        headers={
            "Accept": "application/vnd.github+json",
            # GitHub odrzuca zapytania bez User-Agent.
            # Nazwa BEZ polskich znaków: nagłówki HTTP muszą dać się zapisać
            # w latin-1, a „ł” z „Suma Wpłat” wywala całe zapytanie.
            "User-Agent": f"SumaWplat/{APP_VERSION}",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as odpowiedz:
        return json.loads(odpowiedz.read().decode("utf-8"))


# Nagłówek sekcji ze zmianami w opisie wydania na GitHubie
_NAGLOWEK_ZMIAN = re.compile(r"^(#{1,6})\s*co\s+nowego", re.IGNORECASE)
_NAGLOWEK       = re.compile(r"^(#{1,6})\s+")

# Wiersz, który w sekcji „Co nowego” pełni rolę podtytułu: nagłówek Markdown
# albo wiersz w całości pogrubiony („**Odświeżony interfejs:**”).
_PODTYTUL = re.compile(r"^\s*(?:#{1,6}\s+(.+?)|\*\*(.+?)\*\*)\s*:?\s*$")

# Ile wydań pokazujemy w całości i ile wierszy z każdego. Kto aktualizuje
# po roku, dostałby inaczej ścianę tekstu sklejoną z kilku opisów naraz.
# Reszta jest na stronie wydań i tam po nią odsyłamy.
MAX_WYDAN_W_OKNIE = 4
MAX_WIERSZY_NA_WYDANIE = 18


def _uprosc_markdown(tekst: str) -> str:
    """Zdejmuje najczęstsze znaczniki Markdown.

    Okno pokazuje zwykły tekst, więc gwiazdki i kratki tylko przeszkadzają —
    „**Odświeżony interfejs:**” ma się czytać jak zdanie, a nie jak kod.
    """
    linie = []
    for linia in tekst.split("\n"):
        linia = _NAGLOWEK.sub("", linia.rstrip())              # ## Tytuł → Tytuł
        goła = linia.strip()
        if len(goła) >= 3 and set(goła) in ({"-"}, {"="}, {"*"}, {"_"}):
            continue                                           # linia pozioma
        linia = re.sub(r"^\s*>\s?", "", linia)                 # cytat
        linia = re.sub(r"^(\s*)[-*+]\s+", r"\1  • ", linia)     # punktor
        linia = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1: \2", linia)  # odnośnik
        linia = re.sub(r"\*\*(.+?)\*\*", r"\1", linia)          # pogrubienie
        linia = re.sub(r"(?<!\w)\*(.+?)\*(?!\w)", r"\1", linia) # kursywa
        linia = re.sub(r"`([^`]+)`", r"\1", linia)              # kod
        linie.append(linia)

    # Najwyżej jedna pusta linia z rzędu
    wynik = []
    for linia in linie:
        if not linia.strip() and wynik and not wynik[-1].strip():
            continue
        wynik.append(linia)
    return "\n".join(wynik).strip()


def _streszczenie_zmian(sekcja: str) -> str:
    """Ze zdań w „Co nowego” wyciąga sam szkielet: podtytuły.

    Opis wydania na GitHubie ma się czytać jak tekst, a nie jak lista
    telegraficznych haseł, więc nie chcemy pisać go dwa razy: raz dla ludzi,
    raz dla programu. Zamiast tego program bierze z gotowego opisu same
    podtytuły — one odpowiadają na jedyne pytanie, jakie ma ktoś patrzący
    w okno aktualizacji: „czy jest tu coś dla mnie”. Po szczegóły klika
    w stronę wydania.

    Gdy podtytułów nie ma (starsze wydania, krótkie opisy), zwracamy pusty
    tekst i pokazujemy sekcję w całości.
    """
    podtytuly = []
    for linia in sekcja.split("\n")[1:]:        # pierwszy wiersz to „Co nowego”
        m = _PODTYTUL.match(linia)
        if m:
            # Dwukropek bywa i w środku pogrubienia („**Interfejs:**”)
            tytul = (m.group(1) or m.group(2) or "").strip().rstrip(":").strip()
            if tytul:
                podtytuly.append("  • " + tytul)
    # Wystarczy jeden: wydanie z jedną zmianą ma jeden podtytuł i to jest
    # dokładnie ta linia, którą chcemy pokazać. Ryzyko pomyłki jest małe,
    # bo za podtytuł uznajemy tylko wiersz pogrubiony W CAŁOŚCI albo
    # nagłówek Markdown, a nie pogrubienie w środku zdania.
    return "\n".join(podtytuly)


def _wyciag_zmian(body: str | None) -> str:
    """Wyciąga z opisu wydania sekcję „Co nowego”.

    Szukamy po NAGŁÓWKU, a nie po kolejności bloków. Dzięki temu opis wydania
    na GitHubie może być ułożony dowolnie — najpierw pobieranie, potem zmiany
    albo odwrotnie — a okno aktualizacji i tak pokaże to, co trzeba. I nie
    trzeba przepisywać opisów wydań, które już są opublikowane.

    Sekcja kończy się na następnym nagłówku tego samego lub wyższego rzędu.
    Gdy takiej sekcji nie ma (na przykład w pierwszym wydaniu), pokazujemy
    całość opisu — lepiej za dużo niż nic.
    """
    tekst = (body or "").strip()
    if not tekst:
        return ""

    linie = tekst.split("\n")
    start, poziom = None, 0
    for i, linia in enumerate(linie):
        m = _NAGLOWEK_ZMIAN.match(linia.strip())
        if m:
            start, poziom = i, len(m.group(1))
            break

    if start is not None:
        wybrane = [linie[start]]
        for linia in linie[start + 1:]:
            m = _NAGLOWEK.match(linia.strip())
            if m and len(m.group(1)) <= poziom:
                break                     # kolejna sekcja tego samego rzędu
            wybrane.append(linia)
        sekcja = "\n".join(wybrane)
        streszczenie = _streszczenie_zmian(sekcja)
        if streszczenie:
            return streszczenie
        tekst = sekcja

    return _uprosc_markdown(tekst)


def _icon_path() -> str | None:
    """Ścieżka do app.ico: w zbudowanym .exe — rozpakowana przez PyInstaller
    do folderu tymczasowego (_MEIPASS), przy uruchomieniu ze źródeł —
    assets/ obok src/. None, gdy pliku nie ma (ikona jest opcjonalna)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    p = base / "assets" / "app.ico"
    return str(p) if p.exists() else None

# ─── Tekst pomocy (wyświetlany w oknie „Pomoc”) ───────────────────────────────

HELP_TEXT = """\
POMOC — Suma Wpłat
══════════════════════════════════════════════════════════════════════════════

1. DO CZEGO SŁUŻY TEN PROGRAM?
─────────────────────────────────────────────────────────────────────────────
Program wczytuje listę operacji wyeksportowaną z bankowości internetowej,
rozpoznaje wpłaty od poszczególnych klientów, przygotowuje czytelny raport
zbiorczy oraz — w razie potrzeby — plik gotowy do zaimportowania jako faktury
do Saldeo Smart.

Obsługiwane banki i formaty:

  • mBank — „Historia” → „Eksportuj listę”, format CSV;
  • PKO BP — „Historia” → „Zrealizowane”, na dole listy „Pobierz zestawienie”,
    format XLS (nie CSV — patrz uwaga niżej).

Bank rozpoznawany jest po zawartości pliku, więc nie trzeba nic przestawiać
w ustawieniach; kto ma konta w obu bankach, po prostu wskazuje raz jeden plik,
raz drugi.

Uwaga o PKO: wybierz XLS, a nie CSV. W pliku XLS nazwa i adres nadawcy stoją
w osobnych kolumnach, a w CSV są wymieszane w jednym polu opisu i dane bywają
przez to niepewne. Jeśli wskażesz plik CSV z PKO, program o tym przypomni.

Typowy cykl pracy: co miesiąc pobierasz wyciąg z bankowości internetowej →
wskazujesz go w programie → uruchamiasz analizę → (opcjonalnie) generujesz
na jej podstawie faktury do importu w Saldeo.


2. PODSTAWOWA ANALIZA LISTY OPERACJI
─────────────────────────────────────────────────────────────────────────────
  1. Kliknij „Wybierz” przy polu „Wyciąg bankowy” i wskaż wyeksportowaną
     z bankowości internetowej listę operacji (mBank — .csv, PKO — .xls).
  2. Jeśli chcesz dodatkowo zapisać wynik do pliku tekstowego — wskaż jego
     ścieżkę w polu „Plik wyjściowy TXT (opcjonalnie)” (krok ten można pominąć).
  3. Sprawdź „Kodowanie pliku” — zwykle właściwe jest „utf-8-sig”; jeżeli polskie
     litery (ą, ę, ś, ż...) wyświetlają się jako „krzaki”, zmień na „cp1250”
     i uruchom analizę ponownie. Ustawienie dotyczy wyłącznie plików CSV
     z mBanku; pliki XLS z PKO niosą kodowanie w sobie i to pole ich
     nie dotyczy.
  4. Kliknij „▶ Uruchom analizę”. W oknie poniżej pojawi się raport: lista
     wpłat pogrupowana według klientów wraz z sumami miesięcznymi i łączną
     tabelą zbiorczą.


3. KONTROLA KOMPLETNOŚCI WYCIĄGU
─────────────────────────────────────────────────────────────────────────────
Przycisk „✓ Kontrola kompletności” tworzy osobny raport, który dzieli
WSZYSTKIE operacje z wyciągu na grupy:

  • KLIENCI — wpłaty od osób prywatnych (te same, które trafiają do głównego
    raportu i do faktur),
  • POZOSTAŁE WPŁYWY — inne przychody (od firm, zwroty, odsetki itp.),
  • WYDATKI — wszystkie obciążenia konta,
  • BLOKADY KARTOWE — tylko w wyciągach PKO i tylko wtedy, gdy jakieś są;
    szczegóły niżej.

Powtarzające się operacje (np. comiesięczne prowizje bankowe) są zwijane do
jednego wiersza z liczbą wystąpień i łączną kwotą. Raport kończy SUMA
KONTROLNA wszystkich operacji — dzięki temu przygotowując dokumenty dla
księgowości łatwo sprawdzisz, że żadna operacja nie została pominięta,
a kwoty zgadzają się z wyciągiem co do grosza.

Raport można zapisać do pliku tekstowego („Zapisz jako”) albo skopiować
w całości do schowka („Kopiuj do schowka”) — np. do wklejenia w e-mailu.

Wskazówka: przycisk działa niezależnie od głównej analizy — wystarczy, że
wskazany jest plik z wyciągiem.

Blokady kartowe to kwoty zablokowane przez autoryzację karty, jeszcze nie
zaksięgowane — w wyciągu PKO nie mają nawet daty operacji. Program pokazuje
je osobną kategorią i NIE wlicza do sum kontrolnych: ta sama płatność wróci
później na wyciąg jako zwykła operacja kartą i policzyłaby się drugi raz.


4. GENEROWANIE FAKTUR DLA SALDEO SMART
─────────────────────────────────────────────────────────────────────────────
Po pomyślnie zakończonej analizie aktywuje się przycisk „Faktury Saldeo”.
Otwiera on okno, w którym można:

  • zaznaczyć / odznaczyć klientów, dla których mają zostać wystawione faktury
    (przyciski „Zaznacz wszystko” / „Odznacz wszystko” ułatwiają pracę przy
    dłuższych listach — Twój wybór jest zapamiętywany do następnego razu),
  • ustawić usługę osobno dla każdego klienta — w kolumnie „Usługa na
    fakturze” obok nazwiska. Podstawia się tam usługa główna z „Dane
    sprzedawcy”, ale można wybrać inną ze słownika (strzałka rozwija całą
    listę, a pisanie zawęża podpowiedzi) albo wpisać zupełnie nową.
    Wybór zapamiętywany jest DLA TEGO KLIENTA i podstawi się następnym razem —
    patrz punkt 6 poniżej. Przy odznaczonym kliencie pole jest wyszarzone:
    skoro faktury dla niego nie wystawiamy, usługa i tak nigdzie nie trafi,
  • ustawić numer początkowy faktury („Lp.”),
  • wybrać „Podstawę zastosowania stawki ZW” (a113, a43, a82, du, iz — zgodnie
    z tym, jak rozliczasz zwolnienie z VAT z urzędem skarbowym),
  • zdecydować, czy oznaczyć faktury jako zapłacone (pole „Zapłacono”).
    Domyślnie WYŁĄCZONE — faktury trafiają do Saldeo jako nieopłacone, dzięki
    czemu możesz je później uzgodnić z wpłatami z wyciągu. Włącz tę opcję tylko,
    jeśli chcesz od razu zaznaczyć je jako opłacone,
  • wskazać plik wynikowy .xlsx — ścieżka jest już wypełniona: to folder
    wczytanego wyciągu (zwykle „Pobrane”) i nazwa z miesiącem wpłat, np.
    „faktury_saldeo_2026-04.xlsx”. Można ją zmienić ręcznie albo przyciskiem
    „Zapisz”. Jeśli wpiszesz samą nazwę, bez folderu, plik trafi tam, gdzie
    leży wyciąg,
  • opcjonalnie wskazać bazę kontrahentów Saldeo, by uniknąć duplikatów —
    patrz punkt 7 poniżej.

Po kliknięciu „Generuj” program tworzy plik Excel gotowy do zaimportowania:
w Saldeo Smart → Faktury → Importuj z pliku.

Uwaga: pole „Data dostawy” Saldeo nie przyjmuje zakresu dat — jeśli klient
płacił kilka razy w miesiącu, program wpisuje datę jego OSTATNIEJ wpłaty w tym
miesiącu. Jeżeli klient w jednym miesiącu zapłacił różne kwoty, program zapisze
to jako jedną pozycję „1 szt. × suma wpłat” i wypisze stosowne ostrzeżenie —
taką fakturę warto sprawdzić ręcznie po imporcie.


5. DANE SPRZEDAWCY
─────────────────────────────────────────────────────────────────────────────
Pozycja „Dane sprzedawcy” w pasku menu u góry okna otwiera ustawienia,
w których można wprowadzić lub poprawić w dowolnej chwili:

  • imię / nazwę sprzedawcy — pojawia się na fakturze jako „Wystawca faktury”,
  • numer konta bankowego — pole „Konto bankowe” na fakturze,
  • nazwę usługi — pole „Nazwa towaru” na fakturze (np. „Konsultacja
    psychologiczna”),
  • tło raportów (ciemne lub jasne) — dotyczy pola „Wynik analizy” oraz okna
    „Kontrola kompletności”; zmiana jest widoczna od razu.

Dane te zapisywane są lokalnie na Twoim komputerze i wykorzystywane przy
każdym kolejnym generowaniu faktur — nie trzeba wpisywać ich za każdym razem.
Nikt poza Tobą nie ma do nich dostępu — nie są nigdzie wysyłane.


6. SŁOWNIK USŁUG
─────────────────────────────────────────────────────────────────────────────
Jedna nazwa usługi na wszystkie faktury wystarcza tylko wtedy, gdy robisz dla
wszystkich to samo. Dlatego usługę wybiera się osobno przy każdym kliencie
(punkt 4), a wszystkie użyte nazwy zbierają się w słowniku usług.

Pozycja „Słownik usług” w pasku menu otwiera okno, w którym widać całą listę.
Przy każdej usłudze pokazane jest, do ilu klientów jest przypisana. Można tu:

  • dodać usługę — wpisz nazwę w polu na dole i kliknij „Dodaj”,
  • zmienić nazwę — zaznacz pozycję i kliknij „Zmień nazwę” (albo kliknij ją
    dwukrotnie). Poprawka trafia od razu do wszystkich klientów, którzy tę
    usługę mają przypisaną,
  • usunąć usługę — zaznacz i kliknij „Usuń”. Jeśli jest komuś przypisana,
    program najpierw pokaże listę tych klientów i ostrzeże, że dostaną
    usługę główną.

Usługi głównej (pogrubionej) nie da się stąd usunąć ani przemianować — należy
do danych sprzedawcy i zmienia się ją w menu „Dane sprzedawcy”.

Przy wpisywaniu nowej nazwy program porównuje ją z tym, co już jest w słowniku,
i pyta, jeśli trafi na coś bardzo podobnego („Konsultacja” / „Konsultacje”).
Chroni to słownik przed mnożeniem wariantów tej samej usługi.


7. JAK UNIKNĄĆ DUPLIKATÓW KONTRAHENTÓW W SALDEO
─────────────────────────────────────────────────────────────────────────────
To opcjonalna, ale zalecana funkcja. Chroni przed sytuacją, w której w bazie
kontrahentów Saldeo z czasem pojawia się wiele kart dla tej samej osoby —
np. raz zapisanej jako „Jan Kowalski”, innym razem jako „Kowalski Jan”, a
jeszcze innym razem z drobną literówką. Powoduje to rozdrobnienie historii
płatności klienta na kilka kart i utrudnia później porządki w Saldeo.

Jak z tego skorzystać — krok po kroku:

  a) Zaloguj się do Saldeo Smart, przejdź do sekcji „Kontrahenci” i wyeksportuj
     listę kontrahentów do pliku CSV (zwykle przycisk „Eksportuj”).

  b) Wróć do naszego programu — w oknie „Faktury Saldeo”, w polu „Baza
     kontrahentów Saldeo — plik CSV (opcjonalnie)”, kliknij „Wczytaj”
     i wskaż plik pobrany w poprzednim kroku. Ścieżka do niego zostanie
     zapamiętana — przy kolejnych generowaniach wystarczy ją od czasu do
     czasu odświeżyć nową wersją wyeksportowanej listy.

     Program pilnuje świeżości tej bazy: jeżeli wskazany plik nie pochodzi
     z dzisiaj, przed sprawdzeniem zapyta, czy kontynuować ze starszym
     plikiem, czy przerwać i najpierw pobrać świeży eksport z Saldeo.

  c) Kliknij „Generuj”. Program porówna nazwy klientów z listy operacji z bazą
     kontrahentów Saldeo i obsłuży trzy sytuacje:

       • DOKŁADNE DOPASOWANIE — nic się nie dzieje, faktura trafi prosto do
         istniejącej karty kontrahenta;

       • PODOBNA, ALE NIE IDENTYCZNA NAZWA — otworzy się dodatkowe okno
         „Podobni kontrahenci w bazie Saldeo”. Dla każdej takiej pary zobaczysz
         nazwę z listy operacji i podobną nazwę już istniejącą w Saldeo, a obok —
         pole wyboru „To ten sam kontrahent — użyj nazwy z bazy Saldeo”:
            – zaznacz je, jeśli to ta sama osoba — wówczas faktura zostanie
              zapisana pod nazwą DOKŁADNIE TAKĄ, jak w bazie Saldeo, dzięki
              czemu przy imporcie trafi do istniejącej karty (bez duplikatu);
            – zostaw niezaznaczone, jeśli to inna osoba — nazwa pozostanie
              taka, jak na liście operacji (Saldeo utworzy dla niej osobną kartę);
         przycisk „Anuluj” w tym oknie przerywa całe generowanie faktur, jeśli
         wolisz najpierw wyjaśnić wątpliwości;

       • ZUPEŁNIE NOWY KLIENT — program tylko informuje, że dla takiej osoby
         Saldeo utworzy nową kartę kontrahenta. To normalne i oczekiwane przy
         pojawieniu się nowego klienta.

  d) Po zapisaniu pliku zobaczysz podsumowanie wszystkich tych ustaleń w
     oknie „Gotowe (z uwagami)”, np.:
        ✓  „Kowalski Jan” zostanie zapisane w fakturze jako
           „JAN KOWALSKI” (zgodnie z bazą Saldeo)...
        ℹ  Nowy kontrahent: „Anna Nowak” — zostanie utworzona dla niego
           nowa karta...

Jeśli nie wskażesz pliku z bazą kontrahentów — program po prostu pominie ten
krok i będzie działał tak, jak dotychczas.


8. NAJCZĘSTSZE PYTANIA
─────────────────────────────────────────────────────────────────────────────
P: Program pokazuje klienta jako „NIEZNANY” — co to znaczy?
O: Nie udało się rozpoznać nazwiska nadawcy w opisie operacji. Zdarza się to
   przy nietypowych formatach przelewów. Taką pozycję warto sprawdzić i w razie
   potrzeby poprawić ręcznie po imporcie do Saldeo.

P: Dlaczego w polu „Data dostawy” jest tylko jedna data, choć klient płacił
   kilka razy w miesiącu?
O: Saldeo nie przyjmuje zakresu dat w tym polu — program celowo wpisuje datę
   ostatniej wpłaty z danego miesiąca. To zgodne z przyjętą praktyką pracy.

P: Czy moje dane (imię, numer konta) są gdzieś wysyłane?
O: Nie. Wyciąg, dane sprzedawcy i wszystko, co program liczy, zostaje
   wyłącznie na Twoim komputerze.

P: Czy program łączy się z internetem?
O: Sam z siebie nigdy — ani przy starcie, ani w tle. Jedyne połączenie
   powstaje wtedy, gdy sam(a) klikniesz „Sprawdź aktualizacje” w oknie
   „O programie”: program pyta wtedy GitHub o numer najnowszej wersji.
   Nie wysyła przy tym niczego o Tobie ani o Twoim wyciągu. Przyciski
   otwierające stronę programu czy Revolut uruchamiają zwykłą przeglądarkę
   — to ona łączy się z siecią, nie program.

P: Program otworzył się w innym rozmiarze niż ostatnio — dlaczego?
O: Rozmiar i położenie okna są zapamiętywane przy zamykaniu. Jeśli okno
   otworzyło się pośrodku ekranu w domyślnym rozmiarze, to znaczy, że
   zapamiętane położenie wypadłoby poza obecny ekran — na przykład po
   odłączeniu drugiego monitora. Program wtedy celowo wraca na środek,
   żeby okno nie zniknęło poza krawędzią.


9. O PROGRAMIE: WERSJA, AKTUALIZACJE, WSPARCIE AUTORA
─────────────────────────────────────────────────────────────────────────────
Pozycja „O programie” w pasku menu pokazuje numer zainstalowanej wersji
(widać go też w pasku tytułu okna głównego), prowadzi na stronę programu
sumawplat.pl, pozwala sprawdzić aktualizacje i zawiera podziękowanie
dla autora.

Przycisk ze stroną otwiera zwykłą przeglądarkę — to przeglądarka łączy się
z internetem, nie program.

AKTUALIZACJE

Program NIE łączy się z internetem sam z siebie — ani przy starcie, ani
w tle. Czyta wyciąg bankowy i nie wysyła nigdzie niczego. Dopiero kliknięcie
„Sprawdź aktualizacje” wysyła jedno zapytanie do GitHuba o listę wydań.
Nie idą z nim żadne dane o Tobie ani o Twoim wyciągu.

W odpowiedzi zobaczysz swoją wersję, najnowszą dostępną oraz listę zmian
między nimi. Program niczego nie pobiera i nie instaluje sam — jeśli
zdecydujesz się zaktualizować, przycisk otworzy stronę wydań w przeglądarce,
a resztą kierujesz Ty.

Aktualizacja instaluje się po prostu na wierzchu poprzedniej wersji.
Ustawienia (dane sprzedawcy, słownik usług) zostają zachowane.

Jeśli komputer nie ma dostępu do internetu albo połączenia blokuje sieć
firmowa, program po prostu powie, że nie udało się sprawdzić, i będzie
działać dalej bez żadnej różnicy.

WSPARCIE AUTORA (całkowicie dobrowolne)

Program jest darmowy i takim pozostanie — to w żaden sposób się nie zmieni.
Jeśli jednak zaoszczędził Ci czasu i miał(a)byś ochotę w jakiś sposób
podziękować autorowi, w oknie „O programie” jest przycisk „Postaw kawę”.
Otwiera Revolut, gdzie można wysłać dowolną kwotę, choćby symboliczną:

    revolut.me/oleksa49b        (RevTag: @oleksa49b)

To czysto symboliczny gest — funkcje programu w żaden sposób od niego nie
zależą.
"""


# ─── Kreator pierwszego uruchomienia (dane sprzedawcy) ────────────────────────

class SellerSetupDialog(tk.Toplevel):
    """
    Pokazywany raz — przy pierwszym uruchomieniu programu (gdy w config.json
    nie ma jeszcze danych sprzedawcy). Pyta o imię/nazwę sprzedawcy, numer
    rachunku bankowego i nazwę usługi — te dane wstawiane są do generowanych
    faktur Saldeo. Wpisane wartości zapisywane są lokalnie w config.json
    i nigdy więcej nie są pytane.

    Dzięki temu w kodzie programu nie ma ani jednej zaszytej na stałe danej
    osobistej — program można swobodnie przekazywać innym użytkownikom.
    """

    def __init__(self, parent: tk.Tk, first_run: bool = True):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._first_run = first_run

        self.title("Konfiguracja — pierwsze uruchomienie" if first_run
                   else "Dane sprzedawcy")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        outer = tk.Frame(self, bg="#f0f0f0", padx=18, pady=16)
        outer.pack(fill=tk.BOTH, expand=True)

        if first_run:
            intro = ("Witaj! Zanim zaczniesz, podaj swoje dane sprzedawcy —\n"
                     "będą automatycznie wpisywane na fakturach Saldeo.\n"
                     "Dane zapiszą się lokalnie na tym komputerze i nie będą\n"
                     "już pytane przy kolejnych uruchomieniach.")
        else:
            intro = ("Tutaj możesz poprawić swoje dane sprzedawcy —\n"
                     "np. jeśli pomyliłeś się przy pierwszym uruchomieniu.\n"
                     "Zmiany zostaną zapisane lokalnie na tym komputerze.")

        tk.Label(
            outer, text=intro,
            font=FONT_UI, bg="#f0f0f0", justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 14))

        cfg = _load_config()

        def _field(label_text, default=""):
            row = tk.Frame(outer, bg="#f0f0f0")
            row.pack(fill=tk.X, pady=(0, 8))
            tk.Label(row, text=label_text, font=FONT_UI, bg="#f0f0f0",
                     width=24, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar(value=default)
            tk.Entry(row, textvariable=var, font=FONT_UI, width=40).pack(
                side=tk.LEFT, fill=tk.X, expand=True)
            return var

        self.name_var    = _field("Imię / nazwa sprzedawcy:",
                                  cfg.get("seller_name", ""))
        self.account_var = _field("Numer konta bankowego:",
                                  cfg.get("seller_account", ""))
        self.service_var = _field("Nazwa usługi (na fakturze):",
                                  cfg.get("service_name", "Usługa"))

        # ── Wygląd: tło raportów (ciemne / jasne) ──
        theme_row = tk.Frame(outer, bg="#f0f0f0")
        theme_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(theme_row, text="Tło raportów:", font=FONT_UI, bg="#f0f0f0",
                 width=24, anchor="w").pack(side=tk.LEFT)
        self.theme_var = tk.StringVar(value=cfg.get("report_theme", "dark"))
        tk.Radiobutton(theme_row, text="ciemne", value="dark",
                       variable=self.theme_var, bg="#f0f0f0",
                       font=FONT_UI,
                       command=self._on_theme_change).pack(side=tk.LEFT)
        tk.Radiobutton(theme_row, text="jasne", value="light",
                       variable=self.theme_var, bg="#f0f0f0",
                       font=FONT_UI,
                       command=self._on_theme_change).pack(side=tk.LEFT, padx=(8, 0))

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(10, 0))

        secondary_btn(btn_row,
                      "Zapisz i kontynuuj" if first_run else "Zapisz",
                      self._save).pack(side=tk.LEFT)

        if not first_run:
            secondary_btn(btn_row, "Anuluj", self.destroy,
                          width=10).pack(side=tk.LEFT, padx=(10, 0))

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.wait_window()

    def _on_theme_change(self):
        """Zmiana tła raportów działa NATYCHMIAST (podgląd na żywo) i od razu
        jest zapisywana — bez tego przełącznik sprawiał wrażenie, że nie
        reaguje (efekt było widać dopiero przy następnym raporcie)."""
        cfg = _load_config()
        cfg["report_theme"] = self.theme_var.get()
        _save_config(cfg)
        parent = self.master
        if hasattr(parent, "apply_report_theme"):
            parent.apply_report_theme()

    def _save(self):
        name    = self.name_var.get().strip()
        account = self.account_var.get().strip()
        service = self.service_var.get().strip()

        if not name or not account:
            messagebox.showerror(
                "Błąd",
                "Podaj przynajmniej nazwę sprzedawcy i numer konta —\n"
                "te dane są niezbędne do wygenerowania faktur.",
                parent=self,
            )
            return

        cfg = _load_config()
        cfg["seller_name"]    = name
        cfg["seller_account"] = account
        cfg["service_name"]   = service or "Usługa"
        cfg["report_theme"]   = self.theme_var.get()
        _save_config(cfg)

        self.destroy()

    def _on_close(self):
        # Nie zamykamy użytkownika na siłę — ale bez danych sprzedawcy
        # generowanie faktur będzie po prostu wstawiać puste wartości,
        # dopóki ich nie uzupełni (kreator można otworzyć ponownie,
        # restartując program po usunięciu seller_name/seller_account
        # z config.json).
        self.destroy()


# ─── Okno „Podobni kontrahenci w bazie Saldeo” ────────────────────────────────

class ContractorMatchDialog(tk.Toplevel):
    """
    Modalne okno rozstrzygania „podejrzanie podobnych” dopasowań do bazy
    kontrahentów Saldeo.

    Dla każdej pary (nazwa z wyciągu ↔ podobna nazwa z bazy Saldeo) użytkownik
    decyduje: to ta sama osoba czy nie.
      • Jeśli „tak” — w fakturze zostanie użyta nazwa Z BAZY SALDEO, aby
        Saldeo przy imporcie podpięło fakturę pod istniejącą kartę kontrahenta
        (a nie utworzyło zduplikowanej).
      • Jeśli „nie” — nazwa pozostanie jak w wyciągu (powstanie nowa karta).

    Wynik — self.result:
      dict {nazwa_z_wyciągu: nazwa_z_bazy_Saldeo} tylko dla potwierdzonych par,
      albo None, gdy użytkownik kliknął „Anuluj” (generowanie należy przerwać).
    """

    def __init__(self, parent: tk.Tk, matches: list[dict]):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("Podobni kontrahenci w bazie Saldeo")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self._matches = matches
        self._vars: list[tk.BooleanVar] = []
        self.result: dict[str, str] | None = None

        self._build_ui()

        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _build_ui(self):
        outer = tk.Frame(self, bg="#f0f0f0", padx=16, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(outer,
                 text=("Niektóre nazwy klientów z listy operacji są bardzo podobne do\n"
                       "kontrahentów już istniejących w bazie Saldeo. Zaznacz, jeśli\n"
                       "to ta sama osoba — wtedy faktura zostanie wystawiona pod\n"
                       "nazwą z bazy Saldeo, co pozwoli uniknąć powstania\n"
                       "zduplikowanej karty kontrahenta."),
                 font=FONT_UI, bg="#f0f0f0", justify="left",
                 anchor="w").pack(anchor="w", pady=(0, 12))

        list_frame = tk.Frame(outer, bg="#f0f0f0")
        list_frame.pack(fill=tk.BOTH, expand=True)

        for m in self._matches:
            var = tk.BooleanVar(value=False)
            self._vars.append(var)

            row = tk.Frame(list_frame, bg="white", relief=tk.GROOVE, bd=1)
            row.pack(fill=tk.X, pady=3)

            tk.Label(row, text=f"Z listy operacji: „{m['name']}”",
                     font=FONT_UI, bg="white", anchor="w").pack(fill=tk.X, padx=8, pady=(6, 0))
            tk.Label(row, text=f"W bazie Saldeo:  „{m['matched']}”",
                     font=FONT_UI, bg="white", fg="#0078d4", anchor="w").pack(fill=tk.X, padx=8)
            tk.Checkbutton(row, variable=var, bg="white", anchor="w",
                           font=FONT_UI,
                           text="To ten sam kontrahent — użyj nazwy z bazy Saldeo"
                           ).pack(fill=tk.X, padx=6, pady=(0, 6))

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(14, 0))

        secondary_btn(btn_row, "Kontynuuj generowanie",
                      self._confirm).pack(side=tk.LEFT)

        secondary_btn(btn_row, "Anuluj", self._cancel,
                      width=10).pack(side=tk.LEFT, padx=(10, 0))

    def _confirm(self):
        self.result = {
            m["name"]: m["matched"]
            for m, var in zip(self._matches, self._vars)
            if var.get()
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ─── Okno „Pomoc” ─────────────────────────────────────────────────────────────

class HelpDialog(tk.Toplevel):
    """Okno pomocy — pokazuje HELP_TEXT w przewijanym polu tekstowym."""

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.transient(parent)

        self.title(f"Pomoc — {APP_NAME}")
        self.configure(bg="#f0f0f0")
        self.geometry("760x620")
        self.minsize(480, 360)

        outer = tk.Frame(self, bg="#f0f0f0", padx=12, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        text = scrolledtext.ScrolledText(
            outer, font=FONT_MONO, wrap=tk.WORD,
            bg="white", fg="#1e1e1e",
            relief=tk.SUNKEN, bd=1,
        )
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", HELP_TEXT)
        text.configure(state=tk.DISABLED)

        secondary_btn(outer, "Zamknij", self.destroy,
                      width=12).pack(pady=(10, 0))

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")


class AboutDialog(tk.Toplevel):
    """Okno „O programie”: wersja, ręczne sprawdzenie aktualizacji, podziękowanie.

    Sprawdzanie aktualizacji jest CELOWO ręczne. Poza tym oknem program nie
    łączy się z siecią w ogóle i to jest jego mocna strona: czyta wyciąg
    bankowy i nie wysyła nigdzie niczego. Automatyczne odpytywanie serwera
    przy każdym starcie ten argument by osłabiło, a przy okazji zostawiało
    ślad w logach (adres IP, częstotliwość uruchomień). Tutaj nic się nie
    dzieje, dopóki użytkownik sam nie kliknie.

    Program niczego też nie pobiera ani nie instaluje sam: pokazuje, co się
    zmieniło, i otwiera stronę wydania w przeglądarce. Automatyczna
    aktualizacja niepodpisanego pliku .exe to dokładnie to zachowanie,
    na które reaguje SmartScreen.
    """

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("O programie")
        self.configure(bg="#f0f0f0")
        self.resizable(False, False)

        self._sprawdzanie = False

        outer = tk.Frame(self, bg="#f0f0f0", padx=18, pady=16)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── Nagłówek ──
        tk.Label(outer, text=APP_NAME, font=("Segoe UI", 16, "bold"),
                 bg="#f0f0f0").pack(anchor="w")
        tk.Label(outer, text=f"Wersja {APP_VERSION}", font=FONT_UI,
                 bg="#f0f0f0", fg="#444444").pack(anchor="w")
        tk.Label(outer,
                 text="Wpłaty od klientów z wyciągu bankowego,\n"
                      "gotowe do faktur w Saldeo Smart.",
                 font=FONT_UI, bg="#f0f0f0", fg="#666666",
                 justify=tk.LEFT).pack(anchor="w", pady=(6, 0))

        secondary_btn(outer, "Strona programu (sumawplat.pl)",
                      lambda: webbrowser.open(WEBSITE_URL)).pack(anchor="w",
                                                                 pady=(10, 0))

        tk.Frame(outer, bg="#d0d0d0", height=1).pack(fill=tk.X, pady=12)

        # ── Aktualizacje ──
        tk.Label(outer, text="Aktualizacje", font=("Segoe UI", 10, "bold"),
                 bg="#f0f0f0").pack(anchor="w")
        tk.Label(outer,
                 text="Program nie łączy się z internetem sam z siebie.\n"
                      "Kliknij poniżej, aby jednorazowo zapytać GitHub\n"
                      "o najnowszą wersję. Nic nie pobiera się automatycznie.",
                 font=("Segoe UI", 9), bg="#f0f0f0", fg="#666666",
                 justify=tk.LEFT).pack(anchor="w", pady=(2, 8))

        self._btn_sprawdz = secondary_btn(outer, "Sprawdź aktualizacje",
                                         self._sprawdz)
        self._btn_sprawdz.pack(anchor="w")

        self._wynik = tk.Label(outer, text="", font=FONT_UI, bg="#f0f0f0",
                               justify=tk.LEFT, wraplength=430)
        self._wynik.pack(anchor="w", pady=(8, 0))

        # Ramka na listę zmian — pojawia się dopiero, gdy jest co pokazać
        self._zmiany_frame = tk.Frame(outer, bg="#f0f0f0")
        self._zmiany = scrolledtext.ScrolledText(
            self._zmiany_frame, font=("Segoe UI", 9), wrap=tk.WORD,
            height=9, width=54, bg="white", fg="#1e1e1e",
            relief=tk.SUNKEN, bd=1,
        )
        self._zmiany.pack(fill=tk.BOTH, expand=True)

        self._btn_pobierz = secondary_btn(
            outer, "Pobierz najnowszą wersję",
            lambda: webbrowser.open(RELEASES_PAGE))

        tk.Frame(outer, bg="#d0d0d0", height=1).pack(fill=tk.X, pady=12)

        # ── Wsparcie autora ──
        tk.Label(outer, text="Wsparcie autora", font=("Segoe UI", 10, "bold"),
                 bg="#f0f0f0").pack(anchor="w")
        tk.Label(outer,
                 text="Program jest darmowy i taki pozostanie.\n"
                      "Jeśli oszczędził Ci czasu, możesz podziękować dowolną\n"
                      "kwotą — choćby symboliczną, „na kawę”. Funkcje programu\n"
                      "w żaden sposób od tego nie zależą.",
                 font=("Segoe UI", 9), bg="#f0f0f0", fg="#666666",
                 justify=tk.LEFT).pack(anchor="w", pady=(2, 8))

        secondary_btn(outer, "☕  Postaw kawę (Revolut)",
                      lambda: webbrowser.open(DONATE_URL)).pack(anchor="w")

        # ── Zamknięcie ──
        secondary_btn(outer, "Zamknij", self.destroy,
                      width=12).pack(anchor="e", pady=(14, 0))

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.wait_window()

    # ── Sprawdzanie aktualizacji ───────────────────────────────────────────

    def _sprawdz(self):
        if self._sprawdzanie:
            return
        self._sprawdzanie = True
        self._btn_sprawdz.configure(state=tk.DISABLED)
        self._wynik.configure(text="Sprawdzam…", fg="#444444")
        self._zmiany_frame.pack_forget()
        self._btn_pobierz.pack_forget()

        # Sieć w osobnym wątku — inaczej okno zamarza do końca zapytania
        threading.Thread(target=self._pobierz_w_tle, daemon=True).start()

    def _pobierz_w_tle(self):
        try:
            wydania = _fetch_releases()
            blad = None
        except urllib.error.URLError as exc:
            wydania, blad = None, getattr(exc, "reason", exc)
        except Exception as exc:                    # timeout, zły JSON itd.
            wydania, blad = None, exc
        # Do widżetów wracamy wyłącznie w wątku głównym
        try:
            if self.winfo_exists():
                self.after(0, lambda: self._pokaz(wydania, blad))
        except tk.TclError:
            pass                                    # okno zamknięto w międzyczasie

    def _pokaz(self, wydania, blad):
        self._sprawdzanie = False
        try:
            self._btn_sprawdz.configure(state=tk.NORMAL)
        except tk.TclError:
            return

        if blad is not None:
            self._wynik.configure(
                fg="#a80000",
                text="Nie udało się sprawdzić aktualizacji.\n"
                     f"({blad})\n\n"
                     "Sprawdź połączenie z internetem albo zajrzyj na stronę\n"
                     "wydań ręcznie — program działa dalej normalnie.",
            )
            self._btn_pobierz.pack(anchor="w", pady=(8, 0))
            return

        moja = _parse_version(APP_VERSION)
        # Wersje robocze i zapowiedzi pomijamy — użytkownikowi pokazujemy
        # tylko to, co naprawdę wydane
        gotowe = [w for w in (wydania or [])
                  if not w.get("draft") and not w.get("prerelease")]
        nowsze = [w for w in gotowe if _parse_version(w.get("tag_name", "")) > moja]

        if not nowsze:
            self._wynik.configure(
                fg="#107c10",
                text=f"Masz najnowszą wersję ({APP_VERSION}).",
            )
            return

        # Od najnowszej do najstarszej
        nowsze.sort(key=lambda w: _parse_version(w.get("tag_name", "")), reverse=True)
        najnowsza = nowsze[0].get("tag_name", "").lstrip("vV")

        self._wynik.configure(
            fg="#0a5b0a",
            text=f"Masz wersję {APP_VERSION}, najnowsza to {najnowsza}.\n"
                 f"Co się zmieniło od Twojej wersji:",
        )

        tekst = []
        for w in nowsze[:MAX_WYDAN_W_OKNIE]:
            tag = w.get("tag_name", "").lstrip("vV")
            nazwa = (w.get("name") or "").strip()
            naglowek = f"Wersja {tag}" + (f" — {nazwa}" if nazwa and nazwa != tag else "")
            tekst.append(naglowek)
            tekst.append("─" * len(naglowek))
            opis = _wyciag_zmian(w.get("body")) or "(brak opisu zmian)"
            wiersze = opis.split("\n")
            if len(wiersze) > MAX_WIERSZY_NA_WYDANIE:
                wiersze = wiersze[:MAX_WIERSZY_NA_WYDANIE]
                wiersze.append("   […] dalszy ciąg na stronie wydań")
            tekst.extend(wiersze)
            tekst.append("")

        # Starsze wydania tylko wymieniamy z nazwy
        pominiete = nowsze[MAX_WYDAN_W_OKNIE:]
        if pominiete:
            numery = ", ".join(w.get("tag_name", "").lstrip("vV")
                               for w in pominiete)
            tekst.append(f"Wcześniejsze wersje ({numery}) opisane są "
                         "na stronie wydań.")

        self._zmiany.configure(state=tk.NORMAL)
        self._zmiany.delete("1.0", tk.END)
        self._zmiany.insert("1.0", "\n".join(tekst).strip())
        self._zmiany.configure(state=tk.DISABLED)

        self._zmiany_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self._btn_pobierz.pack(anchor="w", pady=(8, 0))


# ─── Okno „Kontrola kompletności wyciągu” ─────────────────────────────────────

class ReconciliationDialog(tk.Toplevel):
    """Okno osobnego raportu „Kontrola kompletności” —
    podział WSZYSTKICH operacji wyciągu na Klienci / Pozostałe wpływy / Wydatki."""

    def __init__(self, parent: tk.Tk, report_text: str, default_dir: str, default_name: str):
        super().__init__(parent)
        # UWAGA: celowo BEZ self.transient(parent) — okna „transient" nie mają
        # w Windows standardowych przycisków minimalizuj/maksymalizuj, a raport
        # bywa szeroki i użytkownik chce móc zmaksymalizować okno jednym kliknięciem.

        self._report_text = report_text
        self._default_dir = default_dir
        self._default_name = default_name

        self.title("Kontrola kompletności wyciągu")
        self.configure(bg="#f0f0f0")
        # Szerokość dobrana do najdłuższych wierszy raportu (~105 znaków mono)
        self.geometry("1000x640")
        self.minsize(560, 400)

        outer = tk.Frame(self, bg="#f0f0f0", padx=12, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        theme = _report_theme()
        text = scrolledtext.ScrolledText(
            outer, font=FONT_MONO, wrap=tk.NONE,
            bg=theme["bg"], fg=theme["fg"],
            insertbackground=theme["insert"],
            relief=tk.SUNKEN, bd=1,
        )
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", report_text)
        _koloruj_raport(text, report_text, theme)
        text.configure(state=tk.DISABLED)
        self._text = text   # referencja do przemalowania przy zmianie motywu

        # Poziomy pasek przewijania — raport ma długie wiersze (wrap=NONE),
        # bez niego tekst byłby ucięty bez możliwości przewinięcia myszą
        h_scroll = tk.Scrollbar(outer, orient=tk.HORIZONTAL, command=text.xview)
        h_scroll.pack(fill=tk.X)
        text.configure(xscrollcommand=h_scroll.set)

        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(pady=(10, 0))

        secondary_btn(btn_row, "Zapisz jako", self._save,
                      width=14).pack(side=tk.LEFT, padx=(0, 8))
        self._copy_btn = secondary_btn(btn_row, "Kopiuj do schowka",
                                       self._copy, width=17)
        self._copy_btn.pack(side=tk.LEFT, padx=(0, 8))
        secondary_btn(btn_row, "Zamknij", self.destroy,
                      width=12).pack(side=tk.LEFT)

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

    def apply_theme(self, theme: dict) -> None:
        """Przemalowuje raport na nowy motyw — wywoływane z okna głównego,
        gdy użytkownik przełączy tło w ustawieniach, a to okno jest otwarte."""
        self._text.configure(bg=theme["bg"], fg=theme["fg"],
                             insertbackground=theme["insert"])
        self._text.configure(state=tk.NORMAL)
        _koloruj_raport(self._text, self._report_text, theme)
        self._text.configure(state=tk.DISABLED)

    def _copy(self):
        self.clipboard_clear()
        self.clipboard_append(self._report_text)
        # Krótkie potwierdzenie na samym przycisku (bez wyskakującego okna)
        self._copy_btn.configure(text="Skopiowano ✓")
        self.after(1500, lambda: self._copy_btn.configure(text="Kopiuj do schowka"))

    def _save(self):
        path = filedialog.asksaveasfilename(
            title="Zapisz raport jako...",
            defaultextension=".txt",
            filetypes=[("Pliki tekstowe", "*.txt"), ("Wszystkie pliki", "*.*")],
            initialdir=self._default_dir,
            initialfile=self._default_name,
        )
        if not path:
            return
        try:
            Path(path).write_text(self._report_text, encoding="utf-8")
        except OSError as exc:
            messagebox.showwarning("Nie udało się zapisać pliku", str(exc))


# ─── Okno „Generowanie faktur Saldeo” ─────────────────────────────────────────

class ServicesDialog(tk.Toplevel):
    """Słownik usług wpisywanych na fakturach („Nazwa towaru").

    Bez tego okna słownik dało się tylko zapełniać — usługa wpisana w oknie
    faktur zostawała w nim na zawsze, razem z literówkami i niedokończonymi
    nazwami. Tutaj można go obejrzeć, dodać pozycję, poprawić nazwę i usunąć
    zbędną.

    Zmiany zapisywane są od razu po każdej operacji, dlatego okno ma tylko
    przycisk „Zamknij” — nie ma stanu, który można by zgubić.
    """

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("Słownik usług")
        self.configure(bg="#f0f0f0")
        self.geometry("580x430")
        self.minsize(460, 320)

        outer = tk.Frame(self, bg="#f0f0f0", padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(outer,
                 text="Usługi do wyboru przy klientach w oknie „Faktury Saldeo”:",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        # -- Tabela --
        table = tk.Frame(outer, bg="#f0f0f0")
        table.pack(fill=tk.BOTH, expand=True, pady=(4, 6))

        self._tree = ttk.Treeview(table, columns=("usluga", "uzycie"),
                                  show="headings", selectmode="browse")
        self._tree.heading("usluga", text="Usługa")
        self._tree.heading("uzycie", text="Przypisana do")
        self._tree.column("usluga", width=340, anchor="w")
        self._tree.column("uzycie", width=140, anchor="w", stretch=False)
        # Usługa główna wyróżniona — stąd jej nie ruszamy, bo należy
        # do danych sprzedawcy
        self._tree.tag_configure("main", font=("Segoe UI", 10, "bold"))
        self._tree.bind("<Double-1>", lambda e: self._rename())

        tsb = tk.Scrollbar(table, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=tsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tsb.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Label(outer,
                 text="Pogrubiona pozycja to usługa główna z „Dane sprzedawcy”.",
                 font=("Segoe UI", 8), bg="#f0f0f0", fg="#666666",
                 justify=tk.LEFT).pack(anchor="w", pady=(0, 8))

        # -- Dodawanie --
        add_row = tk.Frame(outer, bg="#f0f0f0")
        add_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(add_row, text="Nowa usługa:", font=FONT_UI,
                 bg="#f0f0f0").pack(side=tk.LEFT)
        self._new_var = tk.StringVar()
        entry = tk.Entry(add_row, textvariable=self._new_var, font=FONT_UI)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 6))
        entry.bind("<Return>", lambda e: self._add())
        secondary_btn(add_row, "Dodaj", self._add, width=10).pack(side=tk.LEFT)

        # -- Przyciski --
        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(fill=tk.X)
        secondary_btn(btn_row, "Zmień nazwę", self._rename,
                      width=14).pack(side=tk.LEFT)
        secondary_btn(btn_row, "Usuń", self._delete,
                      width=10).pack(side=tk.LEFT, padx=(6, 0))
        secondary_btn(btn_row, "Zamknij", self.destroy,
                      width=12).pack(side=tk.RIGHT)

        self._refresh()

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        entry.focus_set()
        self.wait_window()

    # -- Dane ---------------------------------------------------------------

    def _refresh(self, select: str | None = None):
        """Przerysowuje tabelę na podstawie konfiguracji — po każdej zmianie."""
        cfg        = _load_config()
        self._main = (cfg.get("service_name") or "").strip()
        assigned   = cfg.get("client_services", {})

        self._tree.delete(*self._tree.get_children())
        for name in _services_catalog(cfg):
            used = sum(1 for v in assigned.values() if v == name)
            iid  = self._tree.insert(
                "", tk.END,
                values=(name, f"u {_klienci(used)}" if used else "—"),
                tags=("main",) if name == self._main else (),
            )
            if name == select:
                self._tree.selection_set(iid)
                self._tree.see(iid)

    def _selected(self) -> str | None:
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("Nie wybrano usługi",
                                "Zaznacz najpierw pozycję na liście.",
                                parent=self)
            return None
        return self._tree.item(sel[0], "values")[0]

    def _block_main(self, name: str) -> bool:
        """Usługi głównej stąd nie ruszamy: należy do danych sprzedawcy
        i tak czy owak wróciłaby do słownika przy następnym otwarciu."""
        if name != self._main:
            return False
        messagebox.showinfo(
            "Usługa główna",
            f"„{name}” to usługa główna, podstawiana nowym klientom.\n\n"
            "Można ją zmienić w menu „Dane sprzedawcy”, w polu\n"
            "„Nazwa usługi (na fakturze)”.",
            parent=self)
        return True

    # -- Operacje -----------------------------------------------------------

    def _add(self):
        name = self._new_var.get().strip()
        if not name:
            messagebox.showerror("Pusta nazwa",
                                 "Wpisz nazwę usługi, którą chcesz dodać.",
                                 parent=self)
            return

        cfg     = _load_config()
        catalog = _services_catalog(cfg)
        if any(name.lower() == item.lower() for item in catalog):
            messagebox.showinfo("Usługa już jest w słowniku",
                                f"„{name}” jest już na liście.", parent=self)
            return

        # Ta sama ochrona przed wariantami tej samej usługi, co przy
        # wpisywaniu nazwy w oknie faktur
        similar = _similar_service(name, catalog)
        if similar and not messagebox.askyesno(
            "Podobna usługa już istnieje",
            f"W słowniku jest już usługa:\n    „{similar}”\n\n"
            f"Wpisano:\n    „{name}”\n\n"
            "Dodać mimo to jako osobną pozycję?",
            parent=self,
        ):
            return

        cfg["services_catalog"] = sorted(set(catalog) | {name})
        _save_config(cfg)
        self._new_var.set("")
        self._refresh(select=name)

    def _rename(self):
        name = self._selected()
        if name is None or self._block_main(name):
            return

        new = simpledialog.askstring("Zmiana nazwy usługi",
                                     "Nowa nazwa usługi:",
                                     initialvalue=name, parent=self)
        if new is None:
            return
        new = new.strip()
        if not new or new == name:
            return

        cfg     = _load_config()
        catalog = _services_catalog(cfg)
        if any(new.lower() == item.lower() and item != name for item in catalog):
            messagebox.showerror("Nazwa zajęta",
                                 f"Usługa „{new}” jest już w słowniku.",
                                 parent=self)
            return

        # Razem z nazwą poprawiamy przypisania klientów — inaczej wskazywałyby
        # na pozycję, której już nie ma
        cfg["services_catalog"] = sorted(
            {new if item == name else item for item in catalog})
        cfg["client_services"]  = {k: (new if v == name else v)
                                   for k, v in cfg.get("client_services",
                                                       {}).items()}
        _save_config(cfg)
        self._refresh(select=new)

    def _delete(self):
        name = self._selected()
        if name is None or self._block_main(name):
            return

        cfg      = _load_config()
        assigned = cfg.get("client_services", {})
        used     = [k for k, v in assigned.items() if v == name]

        if used:
            lista = "\n".join(f"    \u2022 {k}" for k in sorted(used)[:10])
            if len(used) > 10:
                lista += "\n    \u2022 \u2026"
            if not messagebox.askyesno(
                "Usługa jest w użyciu",
                f"Usługa „{name}” jest przypisana do {_klienci(len(used))}:\n"
                f"{lista}\n\n"
                "Po usunięciu ci klienci dostaną usługę główną:\n"
                f"    „{self._main}”\n\nUsunąć?",
                icon="warning", parent=self,
            ):
                return
            for k in used:
                assigned.pop(k, None)
            cfg["client_services"] = assigned
        elif not messagebox.askyesno("Usunięcie usługi",
                                     f"Usunąć „{name}” ze słownika?",
                                     parent=self):
            return

        cfg["services_catalog"] = sorted(
            item for item in _services_catalog(cfg) if item != name)
        _save_config(cfg)
        self._refresh()


class SaldeoDialog(tk.Toplevel):
    """Modalne okno konfiguracji i generowania pliku Excel importu Saldeo."""

    def __init__(self, parent: tk.Tk, df, source_path: str = ""):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.title("Generowanie faktur Saldeo")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        self._df = df
        self._client_vars: dict[str, tk.BooleanVar] = {}
        self._service_vars: dict[str, tk.StringVar] = {}   # usługa per klient
        self._service_boxes: dict[str, AutocompleteCombobox] = {}

        # Folder i nazwa proponowane dla pliku wynikowego. Bez tego pole było
        # puste, a wpisana sama nazwa lądowała w katalogu roboczym programu
        # (przy uruchomieniu ze skrótu — w folderze instalacji), gdzie nikt
        # by jej nie szukał i gdzie znika przy odinstalowaniu.
        self._out_dir  = _default_output_dir(source_path)
        self._out_name = _default_output_name(df)
        self._result_path: str | None = None

        self._build_ui()

        # Wyśrodkowanie względem okna rodzica
        self.update_idletasks()
        pw = parent.winfo_x() + parent.winfo_width()  // 2
        ph = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"+{pw - self.winfo_width()//2}+{ph - self.winfo_height()//2}")

        self.wait_window()

    # ── Budowa interfejsu ──────────────────────────────────────────────────────

    def _build_ui(self):
        outer = tk.Frame(self, bg="#f0f0f0", padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── Lista klientów ──
        tk.Label(outer, text="Klienci do uwzględnienia w fakturach:",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        list_frame = tk.Frame(outer, bg="white", relief=tk.SUNKEN, bd=1)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 10))

        canvas    = tk.Canvas(list_frame, bg="white", highlightthickness=0, height=200)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        inner     = tk.Frame(canvas, bg="white")

        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Podpowiedzi usług wyświetlają się w osobnym okienku nad listą, więc
        # przy przewijaniu zostałyby „w powietrzu” — chowamy je razem z ruchem
        def _scroll(*args):
            for box in self._service_boxes.values():
                box.hide_popup()
            canvas.yview(*args)

        scrollbar.configure(command=_scroll)

        canvas.pack(side=tk.LEFT,  fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Wypełniamy listę klientów; wcześniej odznaczonych przywracamy z konfiguracji
        cfg_list = _load_config()
        saved_excluded = set(cfg_list.get("excluded_clients", []))
        saved_services = cfg_list.get("client_services", {})
        default_service = cfg_list.get("service_name", "Usługa")
        catalog = _services_catalog(cfg_list)

        # Nagłówek kolumn — bez niego nie widać, po co jest pole obok nazwiska
        head = tk.Frame(inner, bg="white")
        head.pack(fill=tk.X, padx=6, pady=(2, 4))
        tk.Label(head, text="Klient", font=("Segoe UI", 8, "bold"),
                 bg="white", fg="#666666", width=26, anchor="w").pack(side=tk.LEFT)
        tk.Label(head, text="Usługa na fakturze", font=("Segoe UI", 8, "bold"),
                 bg="white", fg="#666666", anchor="w").pack(side=tk.LEFT)

        all_clients = sorted(self._df["name"].unique())
        for client in all_clients:
            var = tk.BooleanVar(value=client not in saved_excluded)
            self._client_vars[client] = var

            row = tk.Frame(inner, bg="white")
            row.pack(fill=tk.X, padx=6, pady=1)

            tk.Checkbutton(row, text=client, variable=var,
                           bg="white", font=FONT_UI, anchor="w", width=26,
                           command=lambda c=client: self._sync_service(c),
                           ).pack(side=tk.LEFT)

            # Usługa dla tego klienta: zapamiętany wybór albo usługa główna.
            # Pole jest edytowalne — można wpisać nową usługę (po wygenerowaniu
            # trafi do słownika) albo wybrać istniejącą: strzałka rozwija cały
            # słownik, a pisanie zawęża podpowiedzi do pasujących pozycji.
            # Szerokości dobrane tak, żeby strzałka mieściła się w oknie —
            # wcześniej wystawała poza listę i nie dało się jej kliknąć.
            svar = tk.StringVar(value=saved_services.get(client, default_service))
            self._service_vars[client] = svar
            box = AutocompleteCombobox(row, catalog, textvariable=svar,
                                       font=FONT_UI, width=26)
            box.pack(side=tk.LEFT, padx=(6, 0))
            self._service_boxes[client] = box
            self._sync_service(client)     # odznaczony klient → pole wyszarzone

        # Przyciski „Zaznacz wszystko / Odznacz wszystko”
        btn_row = tk.Frame(outer, bg="#f0f0f0")
        btn_row.pack(fill=tk.X, pady=(0, 10))
        secondary_btn(btn_row, "Zaznacz wszystko",
                      lambda: self._toggle_all(True), width=14).pack(side=tk.LEFT)
        secondary_btn(btn_row, "Odznacz wszystko",
                      lambda: self._toggle_all(False), width=16).pack(side=tk.LEFT,
                                                                      padx=(6, 0))

        # ── Początkowy numer faktury ──
        num_row = tk.Frame(outer, bg="#f0f0f0")
        num_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(num_row, text="Początkowy numer faktury (Lp.):",
                 font=FONT_UI, bg="#f0f0f0").pack(side=tk.LEFT)
        self.inv_num_var = tk.IntVar(value=1)
        tk.Spinbox(num_row, from_=1, to=9999,
                   textvariable=self.inv_num_var,
                   width=6, font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Podstawa zastosowania stawki ZW ──
        zw_row = tk.Frame(outer, bg="#f0f0f0")
        zw_row.pack(fill=tk.X, pady=(0, 10))

        tk.Label(zw_row, text="Podstawa zastosowania stawki ZW:",
                 font=FONT_UI, bg="#f0f0f0").pack(side=tk.LEFT)

        cfg = _load_config()
        saved_basis = cfg.get("vat_basis", VAT_BASIS)
        self.vat_basis_var = tk.StringVar(value=saved_basis)
        ttk.Combobox(zw_row, textvariable=self.vat_basis_var,
                     values=["a113", "a43", "a82", "du", "iz"],
                     state="readonly", width=8,
                     font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Oznaczenie faktur jako zapłacone (domyślnie WYŁĄCZONE) ──
        self.mark_paid_var = tk.BooleanVar(value=cfg.get("mark_paid", False))
        paid_row = tk.Frame(outer, bg="#f0f0f0")
        paid_row.pack(fill=tk.X, pady=(0, 2))
        tk.Checkbutton(
            paid_row,
            text="Oznacz faktury jako zapłacone (wypełnij kolumnę „Zapłacono”)",
            variable=self.mark_paid_var,
            bg="#f0f0f0", font=FONT_UI, anchor="w",
        ).pack(anchor="w")
        tk.Label(outer,
                 text=("Domyślnie wyłączone — faktury importują się jako nieopłacone,\n"
                       "dzięki czemu można je później uzgodnić z wpłatami z wyciągu."),
                 font=("Segoe UI", 8), fg="#666666", bg="#f0f0f0",
                 justify="left").pack(anchor="w", pady=(0, 10))

        # ── Porównanie z bazą kontrahentów Saldeo (opcjonalnie) ──
        tk.Label(outer, text="Baza kontrahentów Saldeo — plik CSV (opcjonalnie):",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        ref_row = tk.Frame(outer, bg="#f0f0f0")
        ref_row.pack(fill=tk.X, pady=(4, 2))

        self.contractors_csv_var = tk.StringVar(
            value=cfg.get("saldeo_contractors_csv", ""))
        tk.Entry(ref_row, textvariable=self.contractors_csv_var,
                 font=FONT_UI, width=46).pack(side=tk.LEFT, fill=tk.X, expand=True)
        secondary_btn(ref_row, "Wczytaj", self._browse_contractors_csv,
                      width=12).pack(side=tk.LEFT, padx=(6, 0))

        tk.Label(outer,
                 text=("Wskaż plik wyeksportowany z Saldeo (Kontrahenci → Eksportuj),\n"
                       "a program ostrzeże o nowych kontrahentach i podobnych nazwach —\n"
                       "by uniknąć powstawania zduplikowanych kart w bazie Saldeo."),
                 font=("Segoe UI", 8), fg="#666666", bg="#f0f0f0",
                 justify="left").pack(anchor="w", pady=(0, 10))

        # ── Plik wyjściowy ──
        tk.Label(outer, text="Plik wyjściowy Excel (.xlsx):",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        out_row = tk.Frame(outer, bg="#f0f0f0")
        out_row.pack(fill=tk.X, pady=(4, 10))

        self.out_var = tk.StringVar(value=str(self._out_dir / self._out_name))
        tk.Entry(out_row, textvariable=self.out_var,
                 font=FONT_UI, width=52).pack(side=tk.LEFT, fill=tk.X, expand=True)
        secondary_btn(out_row, "Zapisz", self._browse_output,
                      width=12).pack(side=tk.LEFT, padx=(6, 0))

        # ── Przyciski akcji ──
        action_row = tk.Frame(outer, bg="#f0f0f0")
        action_row.pack(pady=(4, 0))

        secondary_btn(action_row, "Generuj",
                      self._generate).pack(side=tk.LEFT)

        secondary_btn(action_row, "Anuluj", self.destroy,
                      width=10).pack(side=tk.LEFT, padx=(10, 0))

    # ── Obsługa zdarzeń ────────────────────────────────────────────────────────

    def _toggle_all(self, state: bool):
        for client, var in self._client_vars.items():
            var.set(state)
            self._sync_service(client)

    def _sync_service(self, client: str):
        """Usługa ma sens tylko dla klienta, dla którego wystawiamy fakturę —
        przy odznaczonym polu wyboru wyszarzamy listę usług, żeby nie kusiła
        do ustawień, które i tak nie trafią do pliku."""
        box = self._service_boxes.get(client)
        if box is None:
            return
        if self._client_vars[client].get():
            box.configure(state="normal")
        else:
            box.hide_popup()
            box.configure(state="disabled")

    def _browse_output(self):
        current = Path(self.out_var.get().strip() or self._out_name)
        path = filedialog.asksaveasfilename(
            title="Zapisz plik importu Saldeo",
            defaultextension=".xlsx",
            filetypes=[("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*")],
            initialdir=str(current.parent if current.is_absolute()
                           else self._out_dir),
            initialfile=current.name,
        )
        if path:
            self.out_var.set(path)

    def _browse_contractors_csv(self):
        path = filedialog.askopenfilename(
            title="Wskaż plik wyeksportowany z Saldeo (lista kontrahentów, CSV)",
            filetypes=[("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")],
        )
        if path:
            self.contractors_csv_var.set(path)

    def _generate(self):
        out_path = self.out_var.get().strip()
        if not out_path:
            messagebox.showerror("Błąd", "Podaj ścieżkę zapisu pliku.",
                                 parent=self)
            return
        # Dodajemy .xlsx, gdy nie podano rozszerzenia
        if not Path(out_path).suffix:
            out_path += ".xlsx"
        # Sama nazwa bez folderu trafiłaby do katalogu roboczego programu —
        # przy starcie ze skrótu jest nim folder instalacji w AppData.
        # Zapisujemy wtedy tam, gdzie leży wczytany wyciąg.
        if not Path(out_path).is_absolute():
            out_path = str(self._out_dir / out_path)
        self.out_var.set(out_path)

        excluded  = {name for name, var in self._client_vars.items() if not var.get()}
        start_num = self.inv_num_var.get()
        vat_basis = self.vat_basis_var.get().strip()
        contractors_csv = self.contractors_csv_var.get().strip()
        mark_paid = self.mark_paid_var.get()

        cfg = _load_config()

        # ── Usługi wybrane przy klientach ──────────────────────────────────
        # Wpisaną ręcznie usługę, której nie ma w słowniku, dopisujemy — ale
        # najpierw sprawdzamy, czy to nie inna pisownia już istniejącej pozycji
        # (inaczej słownik zapełniłby się wariantami tej samej usługi).
        catalog = _services_catalog(cfg)
        selected_services: dict[str, str] = {}
        for client, svar in self._service_vars.items():
            if client in excluded:
                continue
            service = svar.get().strip()
            if not service:
                continue
            if service not in catalog:
                similar = _similar_service(service, catalog)
                if similar and messagebox.askyesno(
                    "Podobna usługa już istnieje",
                    f"W słowniku jest już usługa:\n    „{similar}”\n\n"
                    f"Wpisano:\n    „{service}”\n\n"
                    "Czy chodziło o tę istniejącą usługę?\n"
                    "„Nie” doda wpisaną jako nową pozycję słownika.",
                    parent=self,
                ):
                    service = similar
                    svar.set(similar)
                else:
                    catalog.append(service)
            selected_services[client] = service

        # Zapisujemy ustawienia na następne uruchomienie
        cfg["vat_basis"]               = vat_basis
        cfg["excluded_clients"]        = sorted(excluded)   # odznaczone pola
        cfg["saldeo_contractors_csv"]  = contractors_csv
        cfg["mark_paid"]               = mark_paid
        cfg["services_catalog"]        = sorted(set(catalog))
        # Usługi zapamiętujemy per klient — przy kolejnym generowaniu podstawią
        # się same; wpisy dla klientów spoza tego wyciągu zostawiamy nietknięte
        cfg["client_services"]         = {**cfg.get("client_services", {}),
                                          **selected_services}
        _save_config(cfg)

        # ── Świeżość bazy kontrahentów: baza w Saldeo zmienia się na bieżąco,
        #    więc sprawdzanie duplikatów na starym eksporcie może być mylące.
        #    Datą „eksportu" jest data modyfikacji pliku (moment pobrania). ──
        if contractors_csv:
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(contractors_csv))
            except OSError:
                mtime = None
            if mtime is not None and mtime.date() != date.today():
                if not messagebox.askyesno(
                    "Nieaktualna baza kontrahentów?",
                    "Wskazany plik z bazą kontrahentów Saldeo pochodzi z:\n"
                    f"    {mtime.strftime('%d.%m.%Y %H:%M')}"
                    f"    (dziś jest {date.today().strftime('%d.%m.%Y')})\n\n"
                    "Od tego czasu baza w Saldeo mogła się zmienić i sprawdzenie\n"
                    "duplikatów może być niedokładne.\n\n"
                    "Kontynuować ze starszym plikiem?\n\n"
                    "„Nie” przerwie generowanie — wtedy wyeksportuj świeżą listę\n"
                    "(Saldeo → Kontrahenci → Eksportuj) i wskaż nowy plik.",
                    icon="warning", parent=self,
                ):
                    return

        # ── Porównanie z bazą kontrahentów Saldeo (gdy wskazano plik) ──
        contractor_warnings: list[str] = []
        name_overrides: dict[str, str] = {}
        if contractors_csv:
            try:
                contractors = load_saldeo_contractors(contractors_csv)
                if not contractors:
                    contractor_warnings.append(
                        "⚠ Nie udało się odczytać kontrahentów ze wskazanego pliku CSV — "
                        "upewnij się, że to plik wyeksportowany z Saldeo "
                        "(Kontrahenci → Eksportuj)."
                    )
                else:
                    included = sorted(
                        name for name, var in self._client_vars.items() if var.get()
                    )
                    check_results = check_clients(included, contractors)
                    similar = [r for r in check_results if r["status"] == "similar"]
                    new     = [r for r in check_results if r["status"] == "new"]
                    exact   = [r for r in check_results if r["status"] == "exact"]

                    # Podsumowanie pokazujemy ZAWSZE, także gdy wszystko się
                    # zgadza. Bez tego przy samych trafieniach okno milczało
                    # i nie było wiadomo, czy porównanie w ogóle się wykonało.
                    contractor_warnings.append(
                        f"Sprawdzono {len(check_results)} klientów z bazą Saldeo "
                        f"({len(contractors)} kontrahentów): "
                        f"{len(exact)} dokładnych, {len(similar)} podobnych, "
                        f"{len(new)} nowych."
                    )

                    # Podejrzanie podobne nazwy — decyzję podejmuje użytkownik:
                    # „czy to ten sam kontrahent?” → jeśli tak, w fakturze
                    # zostanie użyta nazwa z bazy Saldeo (jak w bazie),
                    # by nie tworzyć zduplikowanej karty.
                    if similar:
                        dlg = ContractorMatchDialog(self, similar)
                        if dlg.result is None:
                            return  # użytkownik kliknął „Anuluj” — przerywamy generowanie
                        name_overrides = dlg.result

                    for r in similar:
                        if r["name"] in name_overrides:
                            contractor_warnings.append(
                                f"✓ „{r['name']}” zostanie zapisane w fakturze jako "
                                f"„{name_overrides[r['name']]}” — zgodnie z bazą Saldeo, "
                                "by uniknąć powstania zduplikowanej karty kontrahenta."
                            )
                        else:
                            contractor_warnings.append(
                                f"⚠ „{r['name']}” jest podobne do kontrahenta "
                                f"„{r['matched']}” w bazie Saldeo, ale wskazano, że to "
                                "inna osoba — zostanie utworzona osobna karta kontrahenta."
                            )

                    for r in new:
                        contractor_warnings.append(
                            f"ℹ Nowy kontrahent: „{r['name']}” — nie znaleziono go "
                            "w bazie kontrahentów Saldeo. Przy imporcie zostanie "
                            "utworzona dla niego nowa karta."
                        )
            except Exception as exc:
                contractor_warnings.append(
                    f"⚠ Błąd podczas wczytywania bazy kontrahentów Saldeo: {exc}"
                )

        try:
            _, warnings = generate_saldeo_xlsx(
                self._df,
                out_path,
                start_inv_num=start_num,
                exclude_names=excluded,
                name_overrides=name_overrides or None,
                vat_basis=vat_basis or None,
                seller_name=cfg.get("seller_name") or None,
                seller_account=cfg.get("seller_account") or None,
                service_name=cfg.get("service_name") or None,
                mark_paid=mark_paid,
                # Klucze muszą być nazwami PO podmianie z bazy Saldeo —
                # inaczej przemianowani klienci zgubiliby swoją usługę
                client_services={name_overrides.get(k, k): v
                                 for k, v in selected_services.items()},
            )
        except Exception as exc:
            messagebox.showerror("Błąd podczas generowania", str(exc), parent=self)
            return

        all_warnings = contractor_warnings + warnings

        # Podsumowanie porównania z bazą nie jest ostrzeżeniem, tylko informacją.
        # Ostrzeżenie (żółty wykrzyknik) pokazujemy dopiero, gdy jest się czym
        # zająć: duplikaty, nowi kontrahenci, błędy odczytu, różne kwoty wpłat.
        needs_attention = any(
            w.lstrip().startswith(("⚠", "ℹ", "✓")) for w in all_warnings
        )

        msg = f"Plik zapisany:\n{out_path}"
        if all_warnings:
            msg += "\n\nSzczegóły:\n" + "\n".join(all_warnings)
        if needs_attention:
            messagebox.showwarning("Gotowe (z uwagami)", msg, parent=self)
        else:
            messagebox.showinfo("Gotowe", msg, parent=self)

        self._result_path = out_path
        self.destroy()


# ─── Okno główne ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITLE)
        self.resizable(True, True)
        self.minsize(640, 480)

        # Ikona programu w pasku tytułu i na pasku zadań; „default=" sprawia,
        # że dziedziczą ją także wszystkie okna potomne (dialogi)
        ico = _icon_path()
        if ico:
            try:
                self.iconbitmap(default=ico)
            except Exception:
                pass

        self._df = None            # wynik ostatniej analizy
        self._last_report = ""     # tekst ostatniego raportu (do przemalowania motywu)
        # Ostatnia geometria w stanie „normalnym”. Zapamiętujemy ją na bieżąco,
        # bo po zmaksymalizowaniu okna winfo_geometry() zwraca już rozmiar
        # pełnoekranowy i nie da się z niego odtworzyć tego sprzed maksymalizacji.
        self._normal_geometry = ""
        self._hint_job = None      # zaplanowane przeliczenie podpowiedzi

        self._build_ui()

        # Geometrię ustawiamy PO zbudowaniu interfejsu. Zrobiona wcześniej
        # rozjeżdżała się o wysokość paska menu: okno prosiło o 600 px, a po
        # dołożeniu menu raportowało 580. Zapisane 580 przy następnym starcie
        # dawało 560 i okno kurczyło się z każdym uruchomieniem.
        self._restore_geometry()

        # Kreator pierwszego uruchomienia: gdy dane sprzedawcy nie są jeszcze
        # zapisane w config.json — pytamy o nie raz, przed rozpoczęciem pracy.
        cfg = _load_config()
        if not cfg.get("seller_name") or not cfg.get("seller_account"):
            self.update_idletasks()
            SellerSetupDialog(self, first_run=True)

    # ── Budowa interfejsu ──────────────────────────────────────────────────────

    def _build_ui(self):
        self.configure(bg="#f0f0f0")

        # ── Pasek menu (u góry okna) — ustawienia i pomoc bez ikon,
        #    żeby rząd przycisków zawierał wyłącznie akcje analizy ──
        menubar = tk.Menu(self)
        menubar.add_command(label="Dane sprzedawcy",
                            command=self._open_seller_settings)
        menubar.add_command(label="Słownik usług",
                            command=self._open_services)
        menubar.add_command(label="Pomoc", command=self._open_help)
        menubar.add_command(label="O programie", command=self._open_about)
        self.config(menu=menubar)

        # ── Panel sterowania (góra) ──
        ctrl = tk.Frame(self, bg="#f0f0f0", padx=PAD, pady=PAD)
        ctrl.pack(fill=tk.X)

        # Plik wejściowy
        tk.Label(ctrl, text="Wyciąg bankowy:", font=FONT_UI,
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=(0, 2))

        self.input_var = tk.StringVar()
        tk.Entry(ctrl, textvariable=self.input_var,
                 font=FONT_UI, width=68).grid(row=1, column=0,
                                              sticky="ew", padx=(0, 6))
        secondary_btn(ctrl, "Wybierz", self._browse_input,
                      width=12).grid(row=1, column=1)

        # Plik wyjściowy TXT
        tk.Label(ctrl, text="Plik wyjściowy TXT (opcjonalnie):", font=FONT_UI,
                 bg="#f0f0f0").grid(row=2, column=0, sticky="w", pady=(PAD, 2))

        self.output_var = tk.StringVar()
        tk.Entry(ctrl, textvariable=self.output_var,
                 font=FONT_UI, width=68).grid(row=3, column=0,
                                              sticky="ew", padx=(0, 6))
        secondary_btn(ctrl, "Zapisz", self._browse_output,
                      width=12).grid(row=3, column=1)

        # Kodowanie
        enc_row = tk.Frame(ctrl, bg="#f0f0f0")
        enc_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(PAD, 0))

        tk.Label(enc_row, text="Kodowanie pliku:", font=FONT_UI,
                 bg="#f0f0f0").pack(side=tk.LEFT)
        self.enc_var = tk.StringVar(value="utf-8-sig")
        ttk.Combobox(enc_row, textvariable=self.enc_var,
                     values=["utf-8-sig", "cp1250"],
                     state="readonly", width=12,
                     font=FONT_UI).pack(side=tk.LEFT, padx=(8, 0))

        # ── Przyciski akcji ──
        btn_row = tk.Frame(ctrl, bg="#f0f0f0")
        btn_row.grid(row=5, column=0, columnspan=2, pady=(PAD + 4, 0))

        secondary_btn(btn_row, "▶  Uruchom analizę",
                      self._run).pack(side=tk.LEFT)

        self._saldeo_btn = secondary_btn(btn_row, "Faktury Saldeo",
                                         self._open_saldeo_dialog)
        self._saldeo_btn.configure(state=tk.DISABLED)
        self._saldeo_btn.pack(side=tk.LEFT, padx=(12, 0))

        secondary_btn(btn_row, "  ✓ Kontrola kompletności  ",
                      self._open_reconciliation).pack(side=tk.LEFT, padx=(12, 0))

        ctrl.columnconfigure(0, weight=1)

        # ── Separator ──
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=PAD, pady=2)

        # ── Pole tekstowe wyniku ──
        result_frame = tk.Frame(self, bg="#f0f0f0")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=PAD, pady=(0, 4))

        tk.Label(result_frame, text="Wynik analizy:",
                 font=FONT_UI, bg="#f0f0f0").pack(anchor="w")

        theme = _report_theme()
        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            font=FONT_MONO,
            bg=theme["bg"], fg=theme["fg"],
            insertbackground=theme["insert"],
            wrap=tk.NONE,
            state=tk.DISABLED,
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self._show_hint()

        # Poziomy pasek przewijania
        h_scroll = tk.Scrollbar(result_frame, orient=tk.HORIZONTAL,
                                 command=self.result_text.xview)
        h_scroll.pack(fill=tk.X)
        self.result_text.configure(xscrollcommand=h_scroll.set)

        # ── Pasek stanu ──
        self.status_var = tk.StringVar(
            value="☕ Spodobał się program? → menu „O programie”")
        tk.Label(self, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg="#e0e0e0",
                 anchor="w", padx=PAD, pady=3,
                 relief=tk.SUNKEN).pack(fill=tk.X, side=tk.BOTTOM)

    # ── Obsługa przycisków ─────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Wybierz wyciąg bankowy (mBank — CSV, PKO — XLS)",
            filetypes=[
                ("Wyciągi bankowe", "*.csv *.xls *.xlsx"),
                ("mBank — lista operacji (CSV)", "*.csv"),
                ("PKO — zestawienie operacji (XLS)", "*.xls *.xlsx"),
                ("Wszystkie pliki", "*.*"),
            ],
        )
        if not path:
            return
        self.input_var.set(path)
        p = Path(path)
        self.output_var.set(str(p.with_suffix(".txt")))
        self.status_var.set(f"Wybrano plik: {p.name}")

    def _browse_output(self):
        initial_dir  = ""
        initial_file = ""
        if self.input_var.get():
            p = Path(self.input_var.get())
            initial_dir  = str(p.parent)
            initial_file = p.stem + "_raport.txt"

        path = filedialog.asksaveasfilename(
            title="Zapisz raport jako...",
            defaultextension=".txt",
            filetypes=[("Pliki tekstowe", "*.txt"), ("Wszystkie pliki", "*.*")],
            initialdir=initial_dir,
            initialfile=initial_file,
        )
        if path:
            self.output_var.set(path)

    def _run(self):
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showerror("Błąd", "Wybierz plik z wyciągiem bankowym.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("Błąd", f"Plik nie znaleziony:\n{input_path}")
            return

        encoding    = self.enc_var.get()
        output_path = self.output_var.get().strip()

        self.status_var.set("Analizuję...")
        self.update_idletasks()

        try:
            df = analyze(input_path, encoding)
        except Exception as exc:
            messagebox.showerror("Błąd podczas odczytu pliku", str(exc))
            self.status_var.set("Błąd — zobacz komunikat powyżej")
            return

        output_lines: list = []
        try:
            print_report(df, output_lines)
        except Exception as exc:
            messagebox.showerror("Błąd podczas tworzenia raportu", str(exc))
            self.status_var.set("Błąd — zobacz komunikat powyżej")
            return

        # Wyświetlamy w polu tekstowym (motyw mógł się zmienić w ustawieniach,
        # więc stosujemy go przy każdym wyświetleniu raportu)
        theme = _report_theme()
        report_text = "\n".join(output_lines)
        self._last_report = report_text
        self.result_text.configure(state=tk.NORMAL,
                                   bg=theme["bg"], fg=theme["fg"],
                                   insertbackground=theme["insert"])
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, report_text)
        _koloruj_raport(self.result_text, report_text, theme)
        self.result_text.configure(state=tk.DISABLED)
        self.result_text.see("1.0")

        # Zapis do pliku
        saved_msg = ""
        if output_path:
            try:
                Path(output_path).write_text(report_text, encoding="utf-8")
                saved_msg = f"  •  Zapisano: {Path(output_path).name}"
            except OSError as exc:
                messagebox.showwarning("Nie udało się zapisać pliku", str(exc))

        # Zapamiętujemy wynik i odblokowujemy przycisk Saldeo
        self._df = df
        self._saldeo_btn.configure(
            state=tk.NORMAL if not df.empty else tk.DISABLED
        )

        # Pasek stanu
        if df.empty:
            self.status_var.set("Nie znaleziono płatności od osób fizycznych")
        else:
            clients = df["name"].nunique()
            txns    = len(df)
            total   = df["amount"].sum()
            self.status_var.set(
                f"Gotowe  •  {clients} klientów  •  "
                f"{txns} transakcji  •  "
                f"{total:,.2f} PLN".replace(",", " ").replace(".", ",")
                + saved_msg
            )

    def _open_reconciliation(self):
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showerror("Błąd", "Wybierz plik z wyciągiem bankowym.")
            return
        if not Path(input_path).exists():
            messagebox.showerror("Błąd", f"Plik nie znaleziony:\n{input_path}")
            return

        encoding = self.enc_var.get()

        output_lines: list = []
        try:
            print_reconciliation_report(input_path, encoding, output_lines)
        except Exception as exc:
            messagebox.showerror("Błąd podczas tworzenia raportu", str(exc))
            return

        p = Path(input_path)
        ReconciliationDialog(
            self,
            "\n".join(output_lines),
            default_dir=str(p.parent),
            default_name=p.stem + "_kontrola.txt",
        )

    def _open_saldeo_dialog(self):
        if self._df is None or self._df.empty:
            messagebox.showinfo("Brak danych",
                                "Najpierw uruchom analizę listy operacji.")
            return
        SaldeoDialog(self, self._df, self.input_var.get().strip())

    # ── Rozmiar i położenie okna między uruchomieniami ─────────────────────

    def _restore_geometry(self):
        """Przywraca rozmiar i położenie okna z poprzedniego uruchomienia.

        Zapisane położenie sprawdzamy względem obecnego pulpitu: jeśli ktoś
        odłączył drugi monitor, zapamiętane współrzędne wskazywałyby poza
        ekran i okno otworzyłoby się niewidoczne. W takim wypadku wracamy
        na środek.
        """
        cfg = _load_config()
        zapisana = cfg.get("window_geometry") or ""
        zmaksymalizowane = bool(cfg.get("window_maximized"))

        szer, wys, x, y = WIN_W, WIN_H, None, None
        m = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)([+-]\d+)", zapisana.strip())
        if m:
            szer, wys = int(m.group(1)), int(m.group(2))
            x, y = int(m.group(3)), int(m.group(4))

        self.update_idletasks()
        ekran_szer, ekran_wys = self.winfo_screenwidth(), self.winfo_screenheight()

        # Rozmiar nie może przekroczyć ekranu ani zejść poniżej minimum okna
        szer = max(640, min(szer, ekran_szer))
        wys  = max(480, min(wys,  ekran_wys))

        # Czy okno zmieściłoby się na widocznym pulpicie? Wymagamy, żeby pasek
        # tytułu był w zasięgu myszy — inaczej okna nie da się przesunąć.
        widoczne = (x is not None and y is not None
                    and x + szer > 60 and x < ekran_szer - 60
                    and 0 <= y < ekran_wys - 40)
        if not widoczne:
            x = (ekran_szer - szer) // 2
            y = (ekran_wys - wys) // 2

        self.geometry(f"{szer}x{wys}+{x}+{y}")
        self._normal_geometry = f"{szer}x{wys}+{x}+{y}"

        if zmaksymalizowane:
            try:
                self.state("zoomed")
            except tk.TclError:
                pass               # nie każdy system to obsługuje

        self.bind("<Configure>", self._remember_geometry)
        # Zaplanowane przeliczenie podpowiedzi trzeba odwołać przy zamykaniu:
        # inaczej Tk próbuje wywołać metodę okna, którego już nie ma, i wypisuje
        # „invalid command name”. W oknie bez konsoli tego nie widać, ale to
        # nadal błąd, który zaśmieca każde zamknięcie.
        self.bind("<Destroy>", self._cancel_hint_job, add="+")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _cancel_hint_job(self, event=None):
        if event is not None and event.widget is not self:
            return
        if self._hint_job is not None:
            try:
                self.after_cancel(self._hint_job)
            except Exception:
                pass
            self._hint_job = None

    def _remember_geometry(self, event=None):
        """Zapisuje w pamięci geometrię, ale tylko gdy okno jest „normalne”.

        Bez tego filtra po zmaksymalizowaniu zapamiętalibyśmy rozmiar pełnego
        ekranu jako zwykły i po przywróceniu okno zostałoby olbrzymie.
        """
        # <Configure> przychodzi także od każdego widżetu w środku — nas
        # interesuje wyłącznie samo okno
        if event is not None and event.widget is not self:
            return
        try:
            if self.state() == "normal":
                self._normal_geometry = self.winfo_geometry()
        except tk.TclError:
            pass

        # Podpowiedź przeliczamy dopiero tutaj. Zaraz po ustawieniu geometrii
        # pole wyniku nie ma jeszcze docelowego rozmiaru (raportuje wartość
        # sprzed przebudowy) i czcionka wychodziła za mała. Przy okazji
        # podpowiedź dopasowuje się, gdy ktoś rozciągnie okno myszą.
        # Odkładamy o chwilę, żeby przy ciągnięciu ramki nie liczyć w kółko.
        if not self._last_report:
            if self._hint_job is not None:
                try:
                    self.after_cancel(self._hint_job)
                except Exception:
                    pass
            self._hint_job = self.after(120, self._show_hint_if_empty)

    def _show_hint_if_empty(self):
        """Przerysowuje podpowiedź, o ile okno wciąż istnieje i nie ma raportu."""
        self._hint_job = None
        try:
            if self.winfo_exists() and not self._last_report:
                self._show_hint()
        except tk.TclError:
            pass

    def _on_close(self):
        """Zapisuje geometrię do config.json i zamyka program."""
        self._cancel_hint_job()
        cfg = _load_config()
        cfg["window_geometry"] = self._normal_geometry or self.winfo_geometry()
        try:
            cfg["window_maximized"] = self.state() == "zoomed"
        except tk.TclError:
            cfg["window_maximized"] = False
        _save_config(cfg)
        self.destroy()

    # Zakres, w którym dobieramy wielkość podpowiedzi
    HINT_MIN, HINT_MAX = 11, 26
    # O tyle punktów nagłówek jest większy od treści
    HINT_TITLE_BONUS = 3
    # Odstęp pod nagłówkiem i między wierszami treści (piksele)
    HINT_GAP, HINT_SPACING = 10, 3

    def _hint_height(self, rozmiar: int, wierszy: int) -> int:
        """Wysokość całej podpowiedzi w pikselach przy danym rozmiarze.

        Liczona osobno dla nagłówka i treści, bo nagłówek jest większy
        i ma pod sobą odstęp — wspólny wzór zaniżałby wynik i ostatni
        wiersz wychodziłby poza pole.
        """
        naglowek = tkfont.Font(family="Arial", size=rozmiar + self.HINT_TITLE_BONUS,
                               weight="bold")
        tresc = tkfont.Font(family="Arial", size=rozmiar)
        return (naglowek.metrics("linespace") + self.HINT_GAP
                + wierszy * (tresc.metrics("linespace") + self.HINT_SPACING))

    def _hint_font_size(self, linie: list[str]) -> int:
        """Największy rozmiar Arialu, przy którym podpowiedź mieści się
        w polu wyniku i w pionie, i w poziomie.

        Liczymy, zamiast wpisywać na sztywno: przy innym skalowaniu ekranu
        albo po zmianie tekstu sztywna liczba albo urwałaby wiersze, albo
        zostawiła podpowiedź niepotrzebnie małą.

        Rozmiar dobieramy RAZ, pod okno w rozmiarze domyślnym. Po
        zmaksymalizowaniu okna tekst celowo nie rośnie dalej — podpowiedź
        ma być czytelna, a nie zajmować cały ekran.
        """
        self.result_text.update_idletasks()
        szer = self.result_text.winfo_width()
        wys  = self.result_text.winfo_height()
        if szer < 100 or wys < 100:        # okno jeszcze nieułożone
            szer, wys = WIN_W - 40, WIN_H - 270

        najdluzsza = max(linie, key=len)
        wybrany = self.HINT_MIN
        for rozmiar in range(self.HINT_MIN, self.HINT_MAX + 1):
            f = tkfont.Font(family="Arial", size=rozmiar)
            if (f.measure(najdluzsza) <= szer - 30
                    and self._hint_height(rozmiar, len(linie)) <= wys - 16):
                wybrany = rozmiar
            else:
                break
        return wybrany

    def _show_hint(self):
        """Wypisuje podpowiedź w pustym polu wyniku.

        Tekst NIE trafia do self._last_report — inaczej pojechałby do
        zapisywanego pliku i do schowka, i wracałby przy przełączeniu motywu
        już po analizie. Czcionka proporcjonalna, nie monospace: to zdania
        do przeczytania, a nie tabela z liczbami.
        """
        theme = _report_theme()
        tresc = ["   " + l for l in HINT_TEXT.strip("\n").split("\n")]
        rozmiar = self._hint_font_size(tresc)

        self.result_text.configure(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        # Nagłówek w pierwszym wierszu — bez pustego wiersza nad nim,
        # żeby nie tracić wysokości, z której rośnie czcionka
        self.result_text.insert("1.0", HINT_TITLE + "\n" + "\n".join(tresc))

        self.result_text.tag_add("podpowiedz", "2.0", tk.END)
        self.result_text.tag_configure(
            "podpowiedz", font=("Arial", rozmiar),
            foreground=theme["podpowiedz"], spacing1=self.HINT_SPACING,
        )
        # Nagłówek: wyśrodkowany, mocniejszy, z odstępem pod spodem
        self.result_text.tag_add("podpowiedz_tytul", "1.0", "1.end")
        self.result_text.tag_configure(
            "podpowiedz_tytul",
            font=("Arial", rozmiar + self.HINT_TITLE_BONUS, "bold"),
            foreground=theme["podpowiedz_akcent"],
            justify=tk.CENTER, spacing3=self.HINT_GAP,
        )
        self.result_text.configure(state=tk.DISABLED)

    def apply_report_theme(self):
        """Przemalowuje pole wyniku na aktualny motyw — wywoływane od razu
        po przełączeniu tła w ustawieniach (natychmiastowy podgląd),
        bez czekania na kolejną analizę."""
        theme = _report_theme()
        self.result_text.configure(bg=theme["bg"], fg=theme["fg"],
                                   insertbackground=theme["insert"])
        if self._last_report:
            self.result_text.configure(state=tk.NORMAL)
            _koloruj_raport(self.result_text, self._last_report, theme)
            self.result_text.configure(state=tk.DISABLED)
        else:
            self._show_hint()          # jeszcze nie było analizy
        # Przemaluj także otwarte okna „Kontrola kompletności" — bez tego
        # zostawałyby w starym motywie i wyglądało to na błąd
        for child in self.winfo_children():
            if isinstance(child, ReconciliationDialog) and child.winfo_exists():
                child.apply_theme(theme)

    def _open_seller_settings(self):
        """Otwiera okno edycji danych sprzedawcy —
        dostępne w każdej chwili, nie tylko przy pierwszym uruchomieniu."""
        SellerSetupDialog(self, first_run=False)

    def _open_services(self):
        """Otwiera słownik usług — podgląd i porządki w liście usług,
        które podpowiadają się przy klientach w oknie faktur."""
        ServicesDialog(self)

    def _open_help(self):
        """Otwiera okno wbudowanej pomocy użytkownika."""
        HelpDialog(self)

    def _open_about(self):
        """Otwiera okno „O programie” — wersja, aktualizacje, podziękowanie."""
        AboutDialog(self)


# ─── Punkt wejścia ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
