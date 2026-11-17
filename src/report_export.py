"""
report_export.py — raport z analizy wyciągu w pliku Excel.

TXT wystarcza, żeby przeczytać wynik, ale do dalszej pracy (filtr po kliencie,
suma zaznaczenia, przesłanie do biura rachunkowego) potrzebny jest arkusz.
Plik ma osobne karty:

  * „Podsumowanie”  — klient, liczba wpłat, suma; na dole RAZEM,
  * „Wpłaty”        — każda wpłata klienta osobnym wierszem (data, kwota, tytuł),
  * „Firmy”         — wpłaty od firm oznaczonych przez użytkownika (gdy są),
  * „Kontrola”      — CAŁY wyciąg z kategorią każdej operacji i sumą kontrolną,
                      ta sama treść co okno „Kontrola kompletności”.

Kwoty są liczbami, daty datami: w Excelu da się je sumować i sortować,
a nie tylko oglądać.
"""
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from parser import (KATEGORIA_BLOKADY, _collapse_desc, categorize_transaction,
                    load_transactions)

FORMAT_KWOTY = '#,##0.00 "zł";-#,##0.00 "zł"'
FORMAT_DATY = "yyyy-mm-dd"

_NAGLOWEK_FONT = Font(bold=True)
_NAGLOWEK_FILL = PatternFill("solid", fgColor="BDD7EE")
_RAZEM_FONT = Font(bold=True)
_RAZEM_FILL = PatternFill("solid", fgColor="FFF2CC")


def _data(value):
    """„2026-05-14” → data Excela. Pusta albo nietypowa zostaje tekstem."""
    text = str(value or "").strip()
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return text


def _arkusz(wb, tytul: str, naglowki: list[str], szerokosci: list[int]):
    ws = wb.create_sheet(tytul)
    ws.append(naglowki)
    for cell in ws[1]:
        cell.font = _NAGLOWEK_FONT
        cell.fill = _NAGLOWEK_FILL
        cell.alignment = Alignment(vertical="center")
    for i, w in enumerate(szerokosci, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return ws


def _wiersz_razem(ws, etykieta: str, kolumna_kwoty: int, kolumna_liczby: int | None = None):
    """Dopisuje pod tabelą wiersz RAZEM ze wzorem SUMA, żeby suma zgadzała się
    także po ręcznej poprawce w Excelu."""
    ostatni = ws.max_row
    wiersz = [""] * ws.max_column
    wiersz[0] = etykieta
    ws.append(wiersz)
    r = ws.max_row
    lit = get_column_letter(kolumna_kwoty)
    ws.cell(r, kolumna_kwoty, f"=SUM({lit}2:{lit}{ostatni})").number_format = FORMAT_KWOTY
    if kolumna_liczby:
        lit_n = get_column_letter(kolumna_liczby)
        ws.cell(r, kolumna_liczby, f"=SUM({lit_n}2:{lit_n}{ostatni})")
    for cell in ws[r]:
        cell.font = _RAZEM_FONT
        cell.fill = _RAZEM_FILL


def _wplaty(ws, df: pd.DataFrame):
    df = df.copy()
    df["_d"] = df["date"].astype(str)
    for _, row in df.sort_values(["_d", "name"]).iterrows():
        ws.append([_data(row["date"]), row["name"], float(row["amount"]),
                   str(row.get("title") or "")])
        ws.cell(ws.max_row, 1).number_format = FORMAT_DATY
        ws.cell(ws.max_row, 3).number_format = FORMAT_KWOTY
    if len(df):
        ws.auto_filter.ref = f"A1:D{ws.max_row}"
        _wiersz_razem(ws, "RAZEM", 3)


def export_report_xlsx(output_path: str, clients: pd.DataFrame, firms: pd.DataFrame | None,
                       file_path: str, encoding: str, companies: list | None = None) -> str:
    """Zapisuje raport do pliku .xlsx i zwraca jego ścieżkę.

    clients / firms — wynik analyze_split(); file_path, encoding, companies —
    to samo, czym była robiona analiza (potrzebne do karty „Kontrola”).
    """
    wb = Workbook()
    wb.remove(wb.active)

    # ── Podsumowanie ──
    ws = _arkusz(wb, "Podsumowanie", ["Klient", "Wpłat", "Kwota"], [38, 10, 18])
    if not clients.empty:
        summary = (clients.groupby("name")
                   .agg(n=("amount", "count"), s=("amount", "sum"))
                   .sort_values("s", ascending=False))
        for name, row in summary.iterrows():
            ws.append([name, int(row["n"]), float(row["s"])])
            ws.cell(ws.max_row, 3).number_format = FORMAT_KWOTY
        _wiersz_razem(ws, "RAZEM", 3, kolumna_liczby=2)
    else:
        ws.append(["Nie znaleziono płatności od osób fizycznych."])

    # ── Wpłaty ──
    ws = _arkusz(wb, "Wpłaty", ["Data", "Klient", "Kwota", "Tytuł"], [12, 34, 16, 60])
    _wplaty(ws, clients)

    # ── Firmy ──
    if firms is not None and not firms.empty:
        ws = _arkusz(wb, "Firmy", ["Data", "Nadawca", "Kwota", "Tytuł"], [12, 34, 16, 60])
        _wplaty(ws, firms)

    # ── Kontrola ──
    # Układ jak w oknie „Kontrola kompletności”: podsumowanie na górze, potem
    # osobne sekcje Klienci / Pozostałe wpływy / Wydatki (i Blokady kartowe),
    # każda z sumą, na końcu suma kontrolna z sum sekcji. Firmy oznaczone przez
    # użytkownika są, jak w oknie, w „Pozostałe wpływy”.
    ops = load_transactions(file_path, encoding, companies)
    ops = ops.assign(kat=[categorize_transaction(a, bool(c), bool(h))
                          for a, c, h in zip(ops["amount"], ops["is_client"], ops["is_hold"])])
    ws = wb.create_sheet("Kontrola")
    # Pierwsza kolumna szeroka: poza datami mieści podpisy sekcji i sum
    for i, w in enumerate([40, 16, 90], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    bez_blokad = ops[~ops["is_hold"]]
    wplywy = bez_blokad[bez_blokad["amount"] > 0]

    def linia(etykieta, wartosc=None, pogrubiona=False, fill=None):
        ws.append([etykieta, wartosc])
        r = ws.max_row
        if wartosc is not None:
            ws.cell(r, 2).number_format = FORMAT_KWOTY if not isinstance(wartosc, int) else "0"
        if pogrubiona or fill:
            for cell in ws[r]:
                if pogrubiona:
                    cell.font = _RAZEM_FONT
                if fill:
                    cell.fill = fill
        return r

    linia("KONTROLA KOMPLETNOŚCI WYCIĄGU", pogrubiona=True)
    linia("Wszystkich operacji", int(len(bez_blokad)))
    linia(f"Suma wpływów ({len(wplywy)} operacji)", float(wplywy["amount"].sum()))
    linia("Suma wszystkich operacji", float(bez_blokad["amount"].sum()))
    blokady = ops[ops["is_hold"]]
    if len(blokady):
        linia(f"Blokady kartowe, poza sumą ({len(blokady)} operacji)",
              float(blokady["amount"].sum()))

    kategorie = ["Klienci", "Pozostałe wpływy", "Wydatki"]
    if len(blokady):
        kategorie.append(KATEGORIA_BLOKADY)
    sumy = {}
    for kat in kategorie:
        sub = ops[ops["kat"] == kat].copy()
        sub["_d"] = sub["date"].astype(str)
        ws.append([])
        ws.append([f"{kat.upper()}: {len(sub)} operacji", "", ""])
        for cell in ws[ws.max_row]:
            cell.font = _NAGLOWEK_FONT
            cell.fill = _NAGLOWEK_FILL
        ws.append(["Data", "Kwota", "Opis operacji"])
        for cell in ws[ws.max_row]:
            cell.font = _NAGLOWEK_FONT
        start = ws.max_row + 1
        for _, row in sub.sort_values("_d").iterrows():
            ws.append([_data(row["date"]), float(row["amount"]), _collapse_desc(row["desc"])])
            ws.cell(ws.max_row, 1).number_format = FORMAT_DATY
            ws.cell(ws.max_row, 2).number_format = FORMAT_KWOTY
        koniec = ws.max_row
        ws.append([f"Razem {kat.lower()}",
                   f"=SUM(B{start}:B{koniec})" if koniec >= start else 0, ""])
        r = ws.max_row
        ws.cell(r, 2).number_format = FORMAT_KWOTY
        for cell in ws[r]:
            cell.font = _RAZEM_FONT
            cell.fill = _RAZEM_FILL
        sumy[kat] = f"B{r}"

    ws.append([])
    kontrolne = [sumy[k] for k in ("Klienci", "Pozostałe wpływy", "Wydatki")]
    ws.append(["SUMA KONTROLNA (wszystkie operacje)", "=" + "+".join(kontrolne),
               "bez blokad kartowych" if len(blokady) else ""])
    r = ws.max_row
    ws.cell(r, 2).number_format = FORMAT_KWOTY
    for cell in ws[r]:
        cell.font = _RAZEM_FONT
        cell.fill = _RAZEM_FILL

    wb.save(output_path)
    _wlacz_formaty_liczb(output_path)
    return output_path


def _wlacz_formaty_liczb(path: str) -> None:
    """Dopisuje do stylów pliku flagę applyNumberFormat="1".

    openpyxl 3.1 zapisuje styl komórki z numerem formatu, ale tej flagi nie
    serializuje wcale (nie ma jej w CellStyle.__attrs__). Nowy Excel bez niej
    format ignoruje: data pokazuje się jako 46297, kwota bez „zł” (zgłosił
    właściciel 17.11.2026, potwierdzone przez Excel COM: format „Ogólny”).
    Poprawiamy gotowy plik: w xl/styles.xml każdemu <xf> z numFmtId różnym
    od 0 dopisujemy flagę."""
    import re
    import shutil
    import tempfile
    import zipfile

    def fix(xml: str) -> str:
        return re.sub(
            r'<xf (?=[^>]*numFmtId="(?!0")\d+")(?![^>]*applyNumberFormat)',
            '<xf applyNumberFormat="1" ', xml)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx",
                                      dir=str(__import__("pathlib").Path(path).parent))
    tmp.close()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = fix(data.decode("utf-8")).encode("utf-8")
            dst.writestr(item, data)
    shutil.move(tmp.name, path)
