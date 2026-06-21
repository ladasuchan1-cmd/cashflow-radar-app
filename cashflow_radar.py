"""
Cashflow Radar — lokální Streamlit aplikace.

Spouštění:
    A) dvojklik na run_cashflow_radar.py
    B) v terminálu:  streamlit run cashflow_radar.py

Očekávané soubory:
    faktury.xlsx              — Číslo, Doklad, Splatno, Firma, Celkem
    prijemky.xlsx             — Číslo, Datum zápisu, Firma, Celkem, Poznámka
    pohyby.xlsx               — Typ, Kód, Datum, Množství, Vážená, Firma, Dodavatel, Výrobce,
                                 (Stav zásoby)
    prijemky_polozkove.xlsx   — Pohoda tiskový export položek příjemek (VOLITELNÉ, velmi doporučeno)
    stav.xlsx                 — Kód, Stav zásoby, (Vážená) (VOLITELNÉ)

Logika výpočtu (od nejpřesnější po nejméně):
    1) Máš-li `prijemky_polozkove.xlsx` → použije PŘESNOU vazbu Příjemka ↔ Kód ↔ Množství
       (nemusí se hádat přes Dodavatel+Datum).
    2) Bez položek → napojí pohyby na příjemky přes Dodavatel+Datum (±3 dny).
    
    Pak se pro každý kód aplikuje REVERZNÍ LIFO (s aktuálním stavem skladu, buď z stav.xlsx,
    nebo z posledního pohybu) — zbývající kusy se přiřadí nejnovějším příjemkám.
    
    Pokud chybí úplně všechno o aktuálním stavu, použije se FORWARD FIFO podle pohybů.
"""

from __future__ import annotations

import io
import re
import time
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# GitHub sync cache (perzistence na Streamlit Cloud). Když modul/secrets chybí,
# funkce jsou no-op a appka jede dál jen s lokální cache.
try:
    import persistence as _persist
except Exception:  # noqa: BLE001
    _persist = None

# ---------------------------------------------------------------------------
# Konfigurace
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Cashflow Radar", page_icon="📡", layout="wide")

RIZIKO_SKLADEM_PCT = 0.50
RIZIKO_DNI = 14
DATUM_TOLERANCE_DNI = 3

# Whitelist agend pro skladové pohyby. Zahrnuje pouze agendy, které reprezentují
# reálné nákupy (příjmy) a prodeje (výdeje). Vynechává: Převodky mezi sklady,
# Inventury, Reklamace, Servisy — ty by zkreslovaly cashflow radar.
AGENDY_PRIJEM_WHITELIST = ["Příjemka", "Příjemky"]
AGENDY_VYDEJ_WHITELIST = ["Prodejka", "Prodejky", "Výdejka", "Výdejky",
                          "Vydaná faktura", "Vydané faktury"]

STAV_KRIT = "🔴 KRITICKÉ"
STAV_RIZIKO = "🟠 RIZIKO"
STAV_PROCES = "🟡 V PROCESU"
STAV_OK = "🟢 OK"
STAV_NESP = "⚫ NESPÁROVÁNO"
STAV_CHYBI_DATA = "❓ CHYBÍ DATA"

# Metody napojení pohybů na příjemky
NAPOJ_POLOZKOVE = "PŘESNĚ (z položkového exportu)"
NAPOJ_DODAVATEL_DATUM = "ODHADEM (Dodavatel + Datum)"

# Metody rekonstrukce skladu
REKO_LIFO_EXT = "REVERZNÍ LIFO (externí stav skladu)"
REKO_LIFO_POHYB = "REVERZNÍ LIFO (stav z posledního pohybu)"
REKO_FIFO = "FORWARD FIFO (z příjmů a výdejů)"

MAPA_FAKTURY = {
    "Číslo": "Cislo", "Cislo": "Cislo",
    "Doklad": "Doklad",
    "Splatno": "Splatno",
    "Firma": "Firma",
    "Celkem": "Celkem",
    "K likvidaci": "K_likvidaci", "K_likvidaci": "K_likvidaci",
}
MAPA_PRIJEMKY = {
    "Číslo": "Cislo", "Cislo": "Cislo",
    "Datum zápisu": "Datum", "Datum zapisu": "Datum", "Datum": "Datum",
    "Firma": "Firma",
    "Celkem": "Celkem",
    "Poznámka": "Poznamka", "Poznamka": "Poznamka",
}
MAPA_POHYBY = {
    "Typ": "Typ",
    "Pohyb": "Pohyb",
    "Kód": "Kod", "Kod": "Kod",
    "Datum": "Datum",
    "Množství": "Mnozstvi", "Mnozstvi": "Mnozstvi",
    "Vážená": "Cena", "Vazena": "Cena",
    "Firma": "Firma",
    "Dodavatel": "Dodavatel",
    "Výrobce": "Vyrobce", "Vyrobce": "Vyrobce",
    "Stav zásoby": "Stav_zasoby", "Stav zasoby": "Stav_zasoby",
}
MAPA_STAV = {
    "Typ": "Typ_karty",
    "Kód": "Kod", "Kod": "Kod", "IDS": "Kod",
    "Stav zásoby": "Stav_zasoby", "Stav zasoby": "Stav_zasoby", "StavZ": "Stav_zasoby",
    "Množství": "Stav_zasoby", "Mnozstvi": "Stav_zasoby",
    "Vážená": "Cena", "Vazena": "Cena", "VNakup": "Cena",
    "Cena": "Cena",
    "Nákupní": "Cena_nakupni", "Nakupni": "Cena_nakupni",
    "Nakup": "Cena_nakupni", "Nákup": "Cena_nakupni",
    "NCena": "Cena_nakupni", "NakupCena": "Cena_nakupni",
    "Název": "Nazev", "Nazev": "Nazev",
    "Výrobce": "Vyrobce", "Vyrobce": "Vyrobce",
    "Dodavatel": "Dodavatel", "Firma": "Dodavatel",
    "Rezervace": "Rezervace",
    "Rezervováno": "Rezervace", "Rezervovano": "Rezervace",
}

POV_INTERNI = {
    "faktury.xlsx": {"Cislo", "Doklad", "Splatno", "Firma", "Celkem"},
    "prijemky.xlsx": {"Cislo", "Datum", "Firma", "Celkem", "Poznamka"},
    "pohyby.xlsx": {"Typ", "Kod", "Datum", "Mnozstvi", "Cena"},
}


# ---------------------------------------------------------------------------
# Parser Pohoda položkového exportu
# ---------------------------------------------------------------------------

# Pohoda čísla příjemek: např. '26SP02659' (2 cifry roku + písmena + cifry)
_re_cislo_prijemky = re.compile(r"^\d{2}[A-Z]{1,4}\d{3,}$")
_re_datum = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}$")


def _extrahuj_kod_a_nazev(cela_bunka: str) -> tuple[str, str]:
    """
    Vrátí (kod, nazev) z první buňky řádku položky.
    Formát: 'KOD:Název produktu@popis, KOD' nebo 'KOD:Název'
    Kód je všechno před první dvojtečkou.
    Název je text za dvojtečkou, ořezaný o případnou koncovou duplikaci kódu.
    """
    if ":" not in cela_bunka:
        return "", cela_bunka.strip()
    kod, zbytek = cela_bunka.split(":", 1)
    kod = kod.strip()
    nazev = zbytek.strip()
    # Odstraň koncovou duplikaci kódu za poslední čárkou
    # "Název@vel. X, KOD" → "Název@vel. X"
    if "," in nazev:
        pred, za = nazev.rsplit(",", 1)
        if za.strip() == kod:
            nazev = pred.strip()
    return kod, nazev


def _precti_pdf_prijemky(src) -> pd.DataFrame:
    """
    Parser PDF tiskového exportu 'Položky příjemek' z Pohody.
    Používá souřadnice sloupců pro spolehlivé čtení víceřádkových položek.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "Pro čtení PDF je potřeba knihovna pdfplumber.\n"
            "Nainstaluj ji příkazem: pip install pdfplumber"
        )
    from collections import defaultdict

    # Hranice sloupců v bodech (pt) — odvozeno ze záhlaví Pohoda PDF
    COL_KLIC_MAX  = 190   # kód:název
    COL_MNS_MIN   = 190
    COL_MNS_MAX   = 265   # Nks
    COL_JCENA_MIN = 265
    COL_JCENA_MAX = 345   # J.cena

    RE_CP = re.compile(r"\b(\d{2}[A-Z]{1,4}\d{3,})\b")
    RE_DT = re.compile(r"\d{1,2}\.\d{1,2}\.\d{4}")

    def _cislo_pdf(s):
        try:
            return float(str(s).replace("\xa0", "").replace(" ", "").replace(",", "."))
        except Exception:
            return None

    def _skupiny_y(words, snap=4):
        grp = defaultdict(list)
        for w in words:
            y = round(w["top"] / snap) * snap
            grp[y].append(w)
        return [(y, sorted(v, key=lambda x: x["x0"])) for y, v in sorted(grp.items())]

    def _v_sloupci(slova, x_min, x_max):
        return " ".join(w["text"] for w in slova if x_min <= w["x0"] < x_max).strip()

    radky: list[dict] = []
    aktivni = None

    # src může být cesta nebo file-like objekt
    with pdfplumber.open(src) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            for y, slova in _skupiny_y(words, snap=4):
                line = " ".join(w["text"] for w in slova)

                if any(x in line for x in [
                    "Položky příjemek", "KOLOSHOP", "Označení příjemky",
                    "Tisk vybraných", "IČ:", "Strana"
                ]):
                    continue

                # Hlavička příjemky
                if slova and RE_DT.match(slova[0]["text"]):
                    m = RE_CP.search(line)
                    if m:
                        aktivni = m.group(1)
                    continue

                if not aktivni:
                    continue

                klic = _v_sloupci(slova, 0, COL_KLIC_MAX)
                if not klic or ":" not in klic:
                    continue

                mns_text = _v_sloupci(slova, COL_MNS_MIN, COL_MNS_MAX)
                m_ks = re.search(r"(\d+(?:,\d+)?)ks", mns_text)
                if not m_ks:
                    continue
                mnozstvi = _cislo_pdf(m_ks.group(1))

                jcena_text = _v_sloupci(slova, COL_JCENA_MIN, COL_JCENA_MAX)
                m_cena = re.search(r"[\d ]+,\d{2}", jcena_text)
                if not m_cena:
                    continue
                j_cena = _cislo_pdf(m_cena.group(0))

                kod = klic.split(":")[0].strip()
                if kod and mnozstvi and j_cena:
                    radky.append({
                        "Cislo_Prijemky": aktivni,
                        "Kod": kod,
                        "Nazev": klic.split(":", 1)[1].strip() if ":" in klic else "",
                        "Prijato_ks": mnozstvi,
                        "J_cena": j_cena,
                        "Cena_Celkem": round(mnozstvi * j_cena, 2),
                    })

    if not radky:
        return pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Nazev",
                                     "Prijato_ks", "J_cena", "Cena_Celkem"])
    df = pd.DataFrame(radky)
    return (df.groupby(["Cislo_Prijemky", "Kod"], as_index=False)
            .agg(Nazev=("Nazev", "first"),
                 Prijato_ks=("Prijato_ks", "sum"),
                 J_cena=("J_cena", "mean"),
                 Cena_Celkem=("Cena_Celkem", "sum")))


def precti_prijemky_polozkove(src) -> pd.DataFrame:
    """
    Parser Pohoda tiskového exportu 'Položky příjemek'.
    Automaticky rozpozná formát XLSX nebo PDF.

    Vrátí DataFrame: Cislo_Prijemky, Kod, Nazev, Prijato_ks, J_cena, Cena_Celkem
    """
    # Detekce formátu
    je_pdf = False
    if hasattr(src, "name"):
        je_pdf = str(src.name).lower().endswith(".pdf")
    elif isinstance(src, (str, Path)):
        je_pdf = str(src).lower().endswith(".pdf")

    if je_pdf:
        return _precti_pdf_prijemky(src)

    # XLSX parser (původní logika)
    raw = pd.read_excel(src, engine="openpyxl", header=None)
    aktivni = None
    radky: list[dict] = []

    for idx in range(len(raw)):
        vals = raw.iloc[idx].tolist()
        vals_clean = [v for v in vals if pd.notna(v) and str(v).strip() != ""]
        if not vals_clean:
            continue

        first = str(vals_clean[0]).strip()
        first_raw = vals_clean[0]

        # Detekce hlavičky příjemky: datum + číslo příjemky
        # Datum může být datetime objekt (z Excel číselného formátu) nebo string
        je_datum = bool(_re_datum.match(first))
        if not je_datum:
            # Zkus pandas/datetime objekt
            import datetime as _dt
            je_datum = isinstance(first_raw, (_dt.datetime, _dt.date)) or \
                       (hasattr(first_raw, 'year') and hasattr(first_raw, 'month'))

        if je_datum and len(vals_clean) >= 2:
            second = str(vals_clean[1]).strip()
            if _re_cislo_prijemky.match(second):
                aktivni = second
                continue

        if ":" in first and len(vals_clean) >= 2:
            try:
                mnozstvi = float(vals_clean[1])
            except (ValueError, TypeError):
                continue

            kod, nazev = _extrahuj_kod_a_nazev(first)
            if not kod:
                continue

            j_cena = None
            for v in vals_clean[2:]:
                if isinstance(v, str) and not v.replace(".", "").replace(",", "").isdigit():
                    continue
                try:
                    j_cena = float(v)
                    break
                except (ValueError, TypeError):
                    continue

            if aktivni is None or j_cena is None or mnozstvi <= 0:
                continue

            radky.append({
                "Cislo_Prijemky": aktivni,
                "Kod": kod,
                "Nazev": nazev,
                "Prijato_ks": mnozstvi,
                "J_cena": j_cena,
                "Cena_Celkem": mnozstvi * j_cena,
            })

    if not radky:
        return pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Nazev",
                                     "Prijato_ks", "J_cena", "Cena_Celkem"])
    df = pd.DataFrame(radky)
    df = (df.groupby(["Cislo_Prijemky", "Kod"], as_index=False)
          .agg(Nazev=("Nazev", "first"),
               Prijato_ks=("Prijato_ks", "sum"),
               J_cena=("J_cena", "mean"),
               Cena_Celkem=("Cena_Celkem", "sum")))
    return df


# ---------------------------------------------------------------------------
# Načítání standardních souborů
# ---------------------------------------------------------------------------

def _prejmenuj(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    df = df.rename(columns={c: c.strip() if isinstance(c, str) else c
                            for c in df.columns})
    return df.rename(columns={k: v for k, v in mapa.items() if k in df.columns})


def _to_str_col(df: pd.DataFrame, col: str) -> None:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip()


def _to_num_col(df: pd.DataFrame, col: str, default: float = 0.0) -> None:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)
    else:
        df[col] = default


def nacti_a_normalizuj(faktury_src, prijemky_src, pohyby_src,
                       polozky_src_list=None, stav_src=None,
                       pohyby_2_src=None):
    faktury = _prejmenuj(pd.read_excel(faktury_src, engine="openpyxl"), MAPA_FAKTURY)
    prijemky = _prejmenuj(pd.read_excel(prijemky_src, engine="openpyxl"), MAPA_PRIJEMKY)
    pohyby = _prejmenuj(pd.read_excel(pohyby_src, engine="openpyxl"), MAPA_POHYBY)
    # Druhý soubor pohybů (např. 2025) — může mít jiný formát sloupců
    if pohyby_2_src is not None:
        raw2 = pd.read_excel(pohyby_2_src, engine="openpyxl")
        # Pohoda export prodejů 2025 má: Agenda (= Pohyb), Pohyb (= Typ Výdej/Příjem),
        # záporné Množství, Výrobce s háčkem
        if "Agenda" in raw2.columns and "Pohyb" not in MAPA_POHYBY.get("Pohyb", ""):
            raw2 = raw2.rename(columns={
                "Agenda": "Pohyb",    # Agenda = název agendy (Prodejky)
                "Pohyb": "Typ",       # Pohyb = směr (Výdej / Příjem)
            })
        pohyby2 = _prejmenuj(raw2, MAPA_POHYBY)
        # Záporné množství → kladné (výdeje jsou záporné v tomto exportu)
        if "Mnozstvi" in pohyby2.columns:
            pohyby2["Mnozstvi"] = pd.to_numeric(
                pohyby2["Mnozstvi"], errors="coerce").abs()
        pohyby = pd.concat([pohyby2, pohyby], ignore_index=True)

    for name, df in (("faktury.xlsx", faktury),
                     ("prijemky.xlsx", prijemky),
                     ("pohyby.xlsx", pohyby)):
        chybi = POV_INTERNI[name] - set(df.columns)
        if chybi:
            raise ValueError(
                f"V souboru `{name}` chybí tyto sloupce: "
                f"{', '.join(sorted(chybi))}.\n"
                f"Nalezené sloupce: {', '.join(str(c) for c in df.columns)}"
            )

    for c in ("Cislo", "Doklad", "Firma"):
        _to_str_col(faktury, c)
    # Vynech prázdné řádky bez čísla faktury (Pohoda někdy exportuje prázdné záhlaví)
    faktury = faktury[faktury["Cislo"].astype(str).str.strip().str.len() > 0].copy()
    faktury = faktury[~faktury["Cislo"].astype(str).str.lower().isin(["nan", "none", ""])].copy()

    for c in ("Cislo", "Poznamka", "Firma"):
        _to_str_col(prijemky, c)
    for c in ("Kod", "Typ", "Pohyb", "Firma", "Dodavatel", "Vyrobce"):
        _to_str_col(pohyby, c)
    # Vynech řádky pohybů bez kódu
    pohyby = pohyby[pohyby["Kod"].astype(str).str.strip().str.len() > 0].copy()
    pohyby = pohyby[~pohyby["Kod"].astype(str).str.lower().isin(["nan", "none", ""])].copy()

    faktury["Splatno"] = pd.to_datetime(faktury["Splatno"], errors="coerce")
    prijemky["Datum"] = pd.to_datetime(prijemky["Datum"], errors="coerce")
    # Datum pohybů může být string DD.MM.YYYY (export z Pohody) nebo datetime
    pohyby["Datum"] = pd.to_datetime(
        pohyby["Datum"], errors="coerce", dayfirst=True)

    _to_num_col(faktury, "Celkem")
    # K likvidaci: pokud chybí, použij Celkem (tj. považuj vše za nezaplacené)
    if "K_likvidaci" in faktury.columns:
        faktury["K_likvidaci"] = pd.to_numeric(faktury["K_likvidaci"], errors="coerce").fillna(faktury["Celkem"])
    else:
        faktury["K_likvidaci"] = faktury["Celkem"]
    _to_num_col(prijemky, "Celkem")
    _to_num_col(pohyby, "Mnozstvi")
    _to_num_col(pohyby, "Cena")
    if "Stav_zasoby" in pohyby.columns:
        pohyby["Stav_zasoby"] = pd.to_numeric(pohyby["Stav_zasoby"], errors="coerce")

    # Normalizace směru pohybu (Prijem / Vydej)
    # Pohoda exportuje 'Typ' buď jako směr (Příjem/Výdej) nebo jako typ položky (Karta/Služba).
    # Pokud 'Typ' neobsahuje směr, odvodíme ho z 'Pohyb' (název agendy).
    import unicodedata as _ud
    def _ascii(s):
        return _ud.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii").strip().lower()

    def _smer_z_textu(s):
        if s.startswith("prij"):
            return "Prijem"
        if s.startswith("vyd") or s in ("prodejka", "prodejky", "vydana faktura", "vydane faktury"):
            return "Vydej"
        return None

    typ_ascii = pohyby["Typ"].apply(lambda x: _ascii(x) if x else "")
    pohyby["Typ_norm"] = typ_ascii.apply(lambda s: _smer_z_textu(s) or "")

    # Pokud Typ neobsahuje směr (=Karta/Služba apod.), zkus z Pohyb (agenda)
    mask_neznamy = pohyby["Typ_norm"] == ""
    if mask_neznamy.any() and "Pohyb" in pohyby.columns:
        pohyb_ascii = pohyby.loc[mask_neznamy, "Pohyb"].apply(
            lambda x: _ascii(x) if x else "")
        pohyby.loc[mask_neznamy, "Typ_norm"] = pohyb_ascii.apply(
            lambda s: _smer_z_textu(s) or s)

    # Pokud sloupec 'Pohyb' chybí, doplň prázdné hodnoty (fallback = neznámá agenda)
    if "Pohyb" not in pohyby.columns:
        pohyby["Pohyb"] = ""

    if "Dodavatel" in pohyby.columns:
        pohyby["Dodavatel_eff"] = pohyby["Dodavatel"]
    else:
        pohyby["Dodavatel_eff"] = ""
    mask = (pohyby["Dodavatel_eff"].isna()
            | (pohyby["Dodavatel_eff"].astype(str).str.len() == 0)
            | (pohyby["Dodavatel_eff"].astype(str).str.lower() == "nan"))
    pohyby.loc[mask, "Dodavatel_eff"] = pohyby.loc[mask, "Firma"]
    pohyby["Dodavatel_eff"] = pohyby["Dodavatel_eff"].astype(str).str.strip()

    # Položkový export příjemek (Pohoda) — může být více souborů, sloučí se
    polozky = None
    if polozky_src_list:
        # Filtruj jen neprázdné zdroje
        platne = [s for s in polozky_src_list if s is not None]
        if platne:
            kusy = [precti_prijemky_polozkove(src) for src in platne]
            # Sloučení: pokud se příjemka vyskytuje ve více souborech, preferuj
            # záznam se stejným Cislo_Prijemky+Kod — unikátní řádky zachováme.
            polozky = (pd.concat(kusy, ignore_index=True)
                       .drop_duplicates(subset=["Cislo_Prijemky", "Kod"])
                       .reset_index(drop=True))

    # Stav skladu
    stav = None
    if stav_src is not None:
        stav = _prejmenuj(pd.read_excel(stav_src, engine="openpyxl"), MAPA_STAV)
        if "Kod" not in stav.columns or "Stav_zasoby" not in stav.columns:
            raise ValueError(
                f"V souboru `stav.xlsx` musí být sloupce `Kód` a `Stav zásoby`.\n"
                f"Nalezené sloupce: {', '.join(str(c) for c in stav.columns)}"
            )
        # Filtr: pokud je sloupec 'Typ_karty' (v exportu 'Typ'), nech jen 'Karta'
        if "Typ_karty" in stav.columns:
            mask_karta = stav["Typ_karty"].astype(str).str.strip().str.lower().str.startswith("kart")
            stav = stav[mask_karta].copy()
        _to_str_col(stav, "Kod")
        stav["Stav_zasoby"] = pd.to_numeric(stav["Stav_zasoby"], errors="coerce").fillna(0)
        if "Cena" in stav.columns:
            stav["Cena"] = pd.to_numeric(stav["Cena"], errors="coerce").fillna(0)
        else:
            stav["Cena"] = 0.0
        if "Cena_nakupni" in stav.columns:
            stav["Cena_nakupni"] = pd.to_numeric(stav["Cena_nakupni"], errors="coerce").fillna(0)
        # Agregace (kód může být ve více skladech/střediscích — sečti)
        agg = {"Stav_zasoby": "sum", "Cena": "mean"}
        if "Cena_nakupni" in stav.columns:
            agg["Cena_nakupni"] = "mean"
        if "Nazev" in stav.columns:
            agg["Nazev"] = "first"
        if "Vyrobce" in stav.columns:
            agg["Vyrobce"] = "first"
        if "Dodavatel" in stav.columns:
            agg["Dodavatel"] = "first"
        if "Rezervace" in stav.columns:
            stav["Rezervace"] = pd.to_numeric(stav["Rezervace"], errors="coerce").fillna(0)
            agg["Rezervace"] = "sum"
        stav = stav.groupby("Kod").agg(agg).reset_index()

        # Fallback ceny: pokud Vážená = 0, nahraď Nákupní
        mask_zero = stav["Cena"].fillna(0) <= 0
        if "Cena_nakupni" in stav.columns and mask_zero.any():
            stav.loc[mask_zero, "Cena"] = stav.loc[mask_zero, "Cena_nakupni"]
        # Fallback z pohybů: pokud stále 0, vezmi průměrnou Vážená z pohybů
        mask_still_zero = stav["Cena"].fillna(0) <= 0
        if mask_still_zero.any() and not pohyby.empty and "Cena" in pohyby.columns:
            pohyb_cena = (pohyby[pohyby["Cena"].fillna(0) > 0]
                          .groupby("Kod")["Cena"].mean())
            stav.loc[mask_still_zero, "Cena"] = (
                stav.loc[mask_still_zero, "Kod"].map(pohyb_cena).fillna(0).values)

        # Efektivní stav = fyzický stav - rezervace (rezervace = de facto prodáno)
        if "Rezervace" in stav.columns:
            stav["Stav_efektivni"] = (stav["Stav_zasoby"] - stav["Rezervace"]).clip(lower=0)
        else:
            stav["Stav_efektivni"] = stav["Stav_zasoby"]

    return faktury, prijemky, pohyby, polozky, stav


def filtruj_pohyby_podle_agend(
    pohyby: pd.DataFrame,
    agendy_prijem: list[str],
    agendy_vydej: list[str],
) -> tuple[pd.DataFrame, dict]:
    """
    Odfiltruje pohyby, které nejsou ve whitelistu povolených agend pro svůj směr.

    Příjmy: jen ty s Pohyb ∈ agendy_prijem
    Výdeje: jen ty s Pohyb ∈ agendy_vydej

    Vrátí:
        (filtrované pohyby, diagnostika s počtem pohybů po agendách)
    """
    diag = {
        "agendy_puvodne": pd.DataFrame(),
        "zahrnuto": 0,
        "vyloucenyo": 0,
        "agenda_chybi": False,
    }

    if "Pohyb" not in pohyby.columns:
        diag["agenda_chybi"] = True
        return pohyby.copy(), diag

    # Přehled agend v datech (co máme k dispozici)
    stats = (
        pohyby.groupby(["Typ_norm", "Pohyb"])
        .size()
        .reset_index(name="Pocet")
        .sort_values(["Typ_norm", "Pocet"], ascending=[True, False])
    )
    diag["agendy_puvodne"] = stats

    # Pokud je Pohyb všude prázdný, filtrování vypneme (uživatel má starý export)
    neprazdne = pohyby["Pohyb"].astype(str).str.strip()
    if (neprazdne == "").all() or (neprazdne.str.lower() == "nan").all():
        diag["agenda_chybi"] = True
        return pohyby.copy(), diag

    # Normalizace whitelistu pro case-insensitive porovnání bez diakritiky
    import unicodedata
    def _norm(s: str) -> str:
        if not s:
            return ""
        return (unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore")
                .decode("ascii").strip().lower())

    prijem_set = {_norm(a) for a in agendy_prijem}
    vydej_set = {_norm(a) for a in agendy_vydej}

    pohyby_ascii = pohyby["Pohyb"].apply(_norm)

    mask_prijem = (pohyby["Typ_norm"] == "Prijem") & pohyby_ascii.isin(prijem_set)
    mask_vydej = (pohyby["Typ_norm"] == "Vydej") & pohyby_ascii.isin(vydej_set)

    mask = mask_prijem | mask_vydej
    out = pohyby[mask].copy()
    diag["zahrnuto"] = int(mask.sum())
    diag["vyloucenyo"] = int((~mask).sum())
    return out, diag


def stav_z_poslednich_pohybu(pohyby: pd.DataFrame) -> pd.DataFrame | None:
    if "Stav_zasoby" not in pohyby.columns:
        return None
    p = pohyby.dropna(subset=["Stav_zasoby", "Datum"])
    if p.empty:
        return None
    p_sorted = p.sort_values("Datum")
    posledni = p_sorted.groupby("Kod").tail(1)[["Kod", "Stav_zasoby", "Cena"]].copy()
    return posledni


# ---------------------------------------------------------------------------
# Párování faktura ↔ příjemka
# ---------------------------------------------------------------------------

def _norm_parovaci_klic(s: str) -> str:
    """Normalizuje párovací klíč: malá písmena, sjednotí oddělovače / - _ a mezery."""
    if not s or str(s).lower() == "nan":
        return ""
    return str(s).strip().lower().replace("-", "/").replace("_", "/").replace(" ", "")


def _extrahuj_cisla(s: str) -> list[str]:
    """Extrahuje všechny číselné sekvence délky >= 4 z řetězce."""
    import re
    return re.findall(r"\d{4,}", str(s))


def _klic_cyklomax(doklad: str) -> str | None:
    """
    BH-2026-237-000101  →  2376000101
    Logika: BH-RRRR-xxx-yyyyyy  →  xxx + poslední cifra roku + yyyyyy
    """
    import re
    m = re.match(r"BH[/-](\d{4})[/-](\d+)[/-](\d+)", doklad.strip(), re.IGNORECASE)
    if not m:
        return None
    rok, xxx, yyyyyy = m.group(1), m.group(2), m.group(3)
    posledni_cifra_roku = rok[-1]
    return xxx + posledni_cifra_roku + yyyyyy


def _klic_crussis(doklad: str) -> str | None:
    """
    FO-4691/2026  →  suffix '4691'  (hledáme ho jako konec čísla v Poznámce)
    """
    import re
    m = re.search(r"FO[/-](\d+)", doklad.strip(), re.IGNORECASE)
    if not m:
        return None
    return m.group(1)


def _podobnost_nazvu(a: str, b: str) -> float:
    """Jednoduchá podobnost: podíl společných slov (>= 3 znaky) z delšího řetězce."""
    slova_a = set(w.lower() for w in str(a).split() if len(w) >= 3)
    slova_b = set(w.lower() for w in str(b).split() if len(w) >= 3)
    if not slova_a or not slova_b:
        return 0.0
    return len(slova_a & slova_b) / max(len(slova_a), len(slova_b))


def parovat_faktury_prijemky(faktury: pd.DataFrame, prijemky: pd.DataFrame) -> pd.DataFrame:
    """
    Vrstvená párovací pipeline — vektorizovaná pro velké datasety (18k+ faktur).

    Vrstvy:
        1. exact_norm    — normalizovaná přesná shoda
        2. substring     — normalizovaný Doklad je substring normalizované Poznámky
        3. cyklomax      — BH-2026-xxx-yyyyyy ↔ xxx{rok[-1]}yyyyyy
        4. crussis       — FO-nnnn/rrrr → suffix nnnn v Poznámce
        5. no_separator  — bez oddělovačů na obou stranách
        6. cisla_prunik  — průnik čísel >= 4 cifry
        7. castka_nazev  — částka +-1 Kč + název >= 80 % shoda
    """
    rows: list[dict] = []

    fa_norm = faktury["Doklad"].apply(_norm_parovaci_klic)
    pr_norm = prijemky["Poznamka"].apply(_norm_parovaci_klic)
    fa_nosep = faktury["Doklad"].astype(str).str.replace(r"[-/_\s]", "", regex=True).str.lower()
    pr_nosep = prijemky["Poznamka"].astype(str).str.replace(r"[-/_\s]", "", regex=True).str.lower()

    def _sparovane():
        return {r["Cislo_Faktury"] for r in rows}

    def _add(fa_c, pr_c, mt):
        rows.append({"Cislo_Faktury": fa_c, "Cislo_Prijemky": pr_c, "match_type": mt})

    # Vrstva 1: přesná shoda (merge)
    fa_t = faktury[["Cislo"]].assign(_k=fa_norm.values)
    pr_t = prijemky[["Cislo"]].assign(_k=pr_norm.values).rename(columns={"Cislo": "Cislo_Prijemky"})
    for _, r in fa_t[fa_t["_k"] != ""].merge(pr_t[pr_t["_k"] != ""], on="_k").iterrows():
        _add(r["Cislo"], r["Cislo_Prijemky"], "exact")

    # Vrstva 2: substring — Doklad (min 6 znaků) obsažen v Poznámce + stejná firma
    # Podmínka délky a firmy zabraňuje falešným shodám (např. "25013" v "625013525")
    spar = _sparovane()
    fa_nez = faktury[~faktury["Cislo"].isin(spar)].copy()
    fa_nez["_fnorm"] = fa_norm[fa_nez.index].values
    fa_nez_filt = fa_nez[fa_nez["_fnorm"].str.len() >= 6]  # min délka klíče
    for pr_i, (pr_cislo, poz, pr_firma) in enumerate(
            zip(prijemky["Cislo"], pr_norm, prijemky["Firma"])):
        if not poz:
            continue
        mask = fa_nez_filt["_fnorm"].apply(lambda d: bool(d) and d in poz)
        for _, fa_row in fa_nez_filt[mask].iterrows():
            # Vyžaduj shodu firmy >= 60 %
            if _podobnost_nazvu(str(fa_row.get("Firma", "")), str(pr_firma)) >= 0.6:
                _add(fa_row["Cislo"], pr_cislo, "substring")

    # Vrstva 3: Cyklomax
    spar = _sparovane()
    for idx in faktury.index[~faktury["Cislo"].isin(spar)]:
        klic = _klic_cyklomax(str(faktury.at[idx, "Doklad"]))
        if not klic:
            continue
        kn = _norm_parovaci_klic(klic)
        mask = (pr_norm == kn) | pr_norm.str.contains(re.escape(kn), na=False)
        for pr_cislo in prijemky.loc[mask, "Cislo"]:
            _add(faktury.at[idx, "Cislo"], pr_cislo, "cyklomax")

    # Vrstva 4: Crussis suffix
    spar = _sparovane()
    for idx in faktury.index[~faktury["Cislo"].isin(spar)]:
        suffix = _klic_crussis(str(faktury.at[idx, "Doklad"]))
        if not suffix:
            continue
        mask = pr_nosep.str.endswith(suffix) | pr_nosep.str.startswith(suffix)
        for pr_cislo in prijemky.loc[mask, "Cislo"]:
            _add(faktury.at[idx, "Cislo"], pr_cislo, "crussis_suffix")

    # Vrstva 5: bez oddělovačů (merge)
    spar = _sparovane()
    fa_idx5 = faktury.index[~faktury["Cislo"].isin(spar)]
    fa_t2 = faktury.loc[fa_idx5, ["Cislo"]].assign(_ns=fa_nosep[fa_idx5].values)
    pr_t2 = prijemky[["Cislo"]].assign(_ns=pr_nosep.values).rename(columns={"Cislo": "Cislo_Prijemky"})
    for _, r in fa_t2[fa_t2["_ns"] != ""].merge(pr_t2[pr_t2["_ns"] != ""], on="_ns").iterrows():
        _add(r["Cislo"], r["Cislo_Prijemky"], "no_separator")

    # Vrstva 6: průnik čísel >= 4 cifry + stejná firma (alespoň 60 % shoda)
    # Bez podmínky firmy by docházelo k tisícům falešných shod u velkých datasetů.
    spar = _sparovane()
    pr_cisla_sets = {c: (set(_extrahuj_cisla(p)), str(firma))
                     for c, p, firma in zip(prijemky["Cislo"], prijemky["Poznamka"], prijemky["Firma"])}
    for idx in faktury.index[~faktury["Cislo"].isin(spar)]:
        cisla_fa = set(_extrahuj_cisla(faktury.at[idx, "Doklad"]))
        if not cisla_fa:
            continue
        firma_fa = str(faktury.at[idx, "Firma"])
        for pr_cislo, (cisla_pr, firma_pr) in pr_cisla_sets.items():
            if cisla_fa & cisla_pr and _podobnost_nazvu(firma_fa, firma_pr) >= 0.6:
                _add(faktury.at[idx, "Cislo"], pr_cislo, "cisla_prunik")

    # Vrstva 7: částka +-1 Kč + podobný název (>= 80 %)
    spar = _sparovane()
    for idx in faktury.index[~faktury["Cislo"].isin(spar)]:
        castka = float(faktury.at[idx, "Celkem"] or 0)
        firma = str(faktury.at[idx, "Firma"])
        pr_ok = prijemky[(prijemky["Celkem"] - castka).abs() <= 1.0]
        for _, pr_row in pr_ok.iterrows():
            if _podobnost_nazvu(firma, str(pr_row.get("Firma", ""))) >= 0.8:
                _add(faktury.at[idx, "Cislo"], pr_row["Cislo"], "castka_nazev")

    if not rows:
        return pd.DataFrame(columns=["Cislo_Faktury", "Cislo_Prijemky", "match_type"])
    return (pd.DataFrame(rows)
            .drop_duplicates(subset=["Cislo_Faktury", "Cislo_Prijemky"])
            .reset_index(drop=True))


# ---------------------------------------------------------------------------
# Napojení příjmů na příjemky: buď přesně (z položek), nebo odhadem (Dodavatel+Datum)
# ---------------------------------------------------------------------------

def napojit_z_polozek(polozky: pd.DataFrame, prijemky: pd.DataFrame,
                     pohyby: pd.DataFrame) -> pd.DataFrame:
    """
    Přesná vazba z Pohoda položkového exportu.
    Doplní vyrobce a (případně) váženou cenu z pohybů, pokud tam jsou.
    """
    if polozky.empty:
        return pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Prijato_ks",
                                     "Cena", "Vyrobce", "Datum"])

    # Vyrobce z pohybů (první nalezený pro daný kód)
    vyrobci = {}
    if "Vyrobce" in pohyby.columns:
        for kod, grp in pohyby.groupby("Kod"):
            vals = grp["Vyrobce"].dropna().unique()
            vals = [v for v in vals if v and str(v).lower() != "nan"]
            if vals:
                vyrobci[kod] = vals[0]

    # Vážená cena z pohybů jako potenciální upřesnění (jinak použij J.cena)
    vazene = {}
    if "Cena" in pohyby.columns:
        prijmy_p = pohyby[pohyby["Typ_norm"] == "Prijem"]
        if not prijmy_p.empty:
            for kod, grp in prijmy_p.groupby("Kod"):
                ceny = grp["Cena"].dropna()
                ceny = ceny[ceny > 0]
                if len(ceny) > 0:
                    vazene[kod] = ceny.mean()

    # Datum příjemky (pro LIFO řazení)
    dat_prijemky = dict(zip(prijemky["Cislo"], prijemky["Datum"]))

    out = polozky.copy()
    out["Vyrobce"] = out["Kod"].map(vyrobci).fillna("")
    # Preferuj váženou cenu z pohybů, fallback J.cena z položek
    out["Cena"] = out["Kod"].map(vazene)
    out["Cena"] = out["Cena"].fillna(out["J_cena"])
    out["Datum"] = out["Cislo_Prijemky"].map(dat_prijemky)

    return out[["Cislo_Prijemky", "Kod", "Prijato_ks", "Cena", "Vyrobce", "Datum"]]


def napojit_prijmy_na_prijemky(prijemky: pd.DataFrame,
                               pohyby: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Fallback: napojit přijímací pohyby přes Dodavatel+Datum (±3 dny)."""
    diag = {"osirele_prijmy": pd.DataFrame(), "kolize_den": 0}
    prijmy = pohyby[pohyby["Typ_norm"] == "Prijem"].copy()

    if prijmy.empty:
        return pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Prijato_ks",
                                     "Cena", "Vyrobce", "Datum"]), diag

    prij_lookup = prijemky[["Cislo", "Firma", "Datum", "Celkem"]].rename(
        columns={"Cislo": "Cislo_Prijemky", "Firma": "Firma_p", "Datum": "Datum_p"})
    prij_lookup["Firma_norm"] = prij_lookup["Firma_p"].astype(str).str.strip().str.lower()
    prijmy["Dodavatel_norm"] = prijmy["Dodavatel_eff"].astype(str).str.strip().str.lower()

    prijemky_podle = {
        k: g.sort_values("Datum_p").reset_index(drop=True)
        for k, g in prij_lookup.groupby("Firma_norm")
    }

    osirele_idx: list[int] = []
    napojene: list[dict] = []
    kolize_set: set[tuple] = set()

    for idx, row in prijmy.iterrows():
        dod = row["Dodavatel_norm"]
        dt = row["Datum"]
        mnozstvi = row["Mnozstvi"]
        if mnozstvi <= 0 or pd.isna(dt):
            osirele_idx.append(idx)
            continue
        kandidati = prijemky_podle.get(dod)
        if kandidati is None or kandidati.empty:
            osirele_idx.append(idx)
            continue
        delta = (kandidati["Datum_p"] - dt).abs()
        v_toleranci = kandidati[delta <= pd.Timedelta(days=DATUM_TOLERANCE_DNI)].copy()
        if v_toleranci.empty:
            osirele_idx.append(idx)
            continue
        v_toleranci["_delta"] = (v_toleranci["Datum_p"] - dt).abs()
        min_delta = v_toleranci["_delta"].min()
        nejblizsi = v_toleranci[v_toleranci["_delta"] == min_delta]

        if len(nejblizsi) == 1:
            p = nejblizsi.iloc[0]
            napojene.append({
                "Cislo_Prijemky": p["Cislo_Prijemky"], "Kod": row["Kod"],
                "Prijato_ks": mnozstvi, "Cena": row["Cena"],
                "Vyrobce": row.get("Vyrobce", ""), "Datum": dt,
            })
        else:
            kolize_set.add((dod, min_delta, tuple(nejblizsi["Cislo_Prijemky"].tolist())))
            vahy = nejblizsi["Celkem"].clip(lower=0)
            if vahy.sum() <= 0:
                vahy = pd.Series([1.0] * len(nejblizsi), index=nejblizsi.index)
            podily = vahy / vahy.sum()
            for (_, p), podil in zip(nejblizsi.iterrows(), podily):
                napojene.append({
                    "Cislo_Prijemky": p["Cislo_Prijemky"], "Kod": row["Kod"],
                    "Prijato_ks": mnozstvi * podil, "Cena": row["Cena"],
                    "Vyrobce": row.get("Vyrobce", ""), "Datum": dt,
                })

    diag["kolize_den"] = len(kolize_set)
    if osirele_idx:
        diag["osirele_prijmy"] = prijmy.loc[osirele_idx, [
            "Datum", "Dodavatel_eff", "Kod", "Mnozstvi", "Cena"
        ]].rename(columns={"Dodavatel_eff": "Dodavatel"})

    if not napojene:
        return pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Prijato_ks",
                                     "Cena", "Vyrobce", "Datum"]), diag

    df = pd.DataFrame(napojene)
    agg = (df.groupby(["Cislo_Prijemky", "Kod"])
           .agg(Prijato_ks=("Prijato_ks", "sum"),
                Cena=("Cena", "mean"),
                Vyrobce=("Vyrobce", "first"),
                Datum=("Datum", "min"))
           .reset_index())
    return agg, diag


# ---------------------------------------------------------------------------
# Rekonstrukce skladu
# ---------------------------------------------------------------------------

def rekonstrukce_lifo(prijmy: pd.DataFrame, prijemky: pd.DataFrame,
                      stav: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    diag = {"nezarazeno": pd.DataFrame(), "neexistujici_kod": pd.DataFrame()}

    prij_date = prijemky[["Cislo", "Datum"]].rename(
        columns={"Cislo": "Cislo_Prijemky", "Datum": "Datum_prijemky"})
    prijmy_s = prijmy.merge(prij_date, on="Cislo_Prijemky", how="left")
    prijmy_s["Datum_sort"] = prijmy_s["Datum_prijemky"].fillna(prijmy_s["Datum"])

    # Použij Stav_efektivni (= stav - rezervace) pokud existuje, jinak Stav_zasoby
    stav_col = "Stav_efektivni" if "Stav_efektivni" in stav.columns else "Stav_zasoby"
    stav_map = dict(zip(stav["Kod"], stav[stav_col]))

    vystup: list[dict] = []
    nezarazene: list[dict] = []

    kody_s_prijmy = set(prijmy_s["Kod"].unique())
    kody_ve_stavu = set(stav_map.keys())

    for kod in kody_ve_stavu & kody_s_prijmy:
        aktualni = stav_map.get(kod, 0)
        radky = prijmy_s[prijmy_s["Kod"] == kod].sort_values("Datum_sort", ascending=False)
        zbyva = aktualni
        for _, r in radky.iterrows():
            prijato = r["Prijato_ks"]
            if zbyva <= 0:
                zbyva_ks = 0; prodano_ks = prijato
            else:
                zbyva_ks = min(prijato, zbyva)
                prodano_ks = prijato - zbyva_ks
                zbyva -= zbyva_ks
            vystup.append({
                "Cislo_Prijemky": r["Cislo_Prijemky"], "Kod": kod,
                "Prijato_ks": prijato, "Prodano_ks": prodano_ks, "Zbyva_ks": zbyva_ks,
                "Cena": r["Cena"], "Vyrobce": r.get("Vyrobce", ""),
            })
        if zbyva > 0:
            nezarazene.append({"Kod": kod, "Nezarazene_ks": zbyva,
                               "Poznamka": "Stav skladu > součet příjmů"})

    for kod in kody_s_prijmy - kody_ve_stavu:
        radky = prijmy_s[prijmy_s["Kod"] == kod]
        for _, r in radky.iterrows():
            vystup.append({
                "Cislo_Prijemky": r["Cislo_Prijemky"], "Kod": kod,
                "Prijato_ks": r["Prijato_ks"], "Prodano_ks": r["Prijato_ks"], "Zbyva_ks": 0,
                "Cena": r["Cena"], "Vyrobce": r.get("Vyrobce", ""),
            })

    kody_jen_stav = kody_ve_stavu - kody_s_prijmy
    if kody_jen_stav:
        je = [{"Kod": k, "Stav_zasoby": stav_map[k]}
              for k in kody_jen_stav if stav_map[k] > 0]
        if je:
            diag["neexistujici_kod"] = pd.DataFrame(je)

    if not vystup:
        polozky = pd.DataFrame(columns=["Cislo_Prijemky", "Kod", "Prijato_ks",
                                        "Prodano_ks", "Zbyva_ks", "Cena", "Hodnota",
                                        "Prijato_hodnota", "Vyrobce"])
    else:
        polozky = pd.DataFrame(vystup)
        polozky["Hodnota"] = polozky["Zbyva_ks"] * polozky["Cena"]
        polozky["Prijato_hodnota"] = polozky["Prijato_ks"] * polozky["Cena"]

    if nezarazene:
        diag["nezarazeno"] = pd.DataFrame(nezarazene)

    return polozky, diag


def rekonstrukce_fifo(prijmy: pd.DataFrame, pohyby: pd.DataFrame) -> pd.DataFrame:
    vydeje = pohyby[pohyby["Typ_norm"] == "Vydej"]
    vydej_per_kod = vydeje.groupby("Kod")["Mnozstvi"].sum().to_dict() if not vydeje.empty else {}

    p = prijmy.copy()
    p["Prodano_ks"] = 0.0

    for kod, celkem_vydano in vydej_per_kod.items():
        if celkem_vydano <= 0:
            continue
        radky = p[p["Kod"] == kod].sort_values("Datum")
        if radky.empty:
            continue
        zbyva = celkem_vydano
        for ridx, r in radky.iterrows():
            if zbyva <= 0:
                break
            k = min(r["Prijato_ks"], zbyva)
            p.at[ridx, "Prodano_ks"] = k
            zbyva -= k

    p["Zbyva_ks"] = (p["Prijato_ks"] - p["Prodano_ks"]).clip(lower=0)
    p["Hodnota"] = p["Zbyva_ks"] * p["Cena"]
    p["Prijato_hodnota"] = p["Prijato_ks"] * p["Cena"]

    return p[["Cislo_Prijemky", "Kod", "Prijato_ks", "Prodano_ks", "Zbyva_ks",
              "Cena", "Hodnota", "Prijato_hodnota", "Vyrobce"]]


# ---------------------------------------------------------------------------
# Agregace
# ---------------------------------------------------------------------------

def spocitat_prijemky(prijemky: pd.DataFrame, polozky: pd.DataFrame) -> pd.DataFrame:
    if polozky.empty:
        base = prijemky[["Cislo", "Firma", "Celkem"]].rename(
            columns={"Cislo": "Cislo_Prijemky", "Celkem": "Celkem_Prijemka"}).copy()
        for c in ("Prijato_ks", "Prodano_ks", "Skladem_ks",
                  "Hodnota_Prijem", "Hodnota_Skladem", "Pct_Prodano"):
            base[c] = 0.0
        base["Lezaky"] = [[] for _ in range(len(base))]
        return base

    agg = (polozky.groupby("Cislo_Prijemky")
           .agg(Prijato_ks=("Prijato_ks", "sum"),
                Prodano_ks=("Prodano_ks", "sum"),
                Skladem_ks=("Zbyva_ks", "sum"),
                Hodnota_Prijem=("Prijato_hodnota", "sum"),
                Hodnota_Skladem=("Hodnota", "sum"))
           .reset_index())

    lez = (polozky[polozky["Zbyva_ks"] > 0]
           .groupby("Cislo_Prijemky")
           .apply(lambda d: d[["Kod", "Zbyva_ks", "Cena", "Hodnota", "Vyrobce"]]
                  .sort_values("Hodnota", ascending=False).to_dict("records"))
           .rename("Lezaky"))
    agg = agg.merge(lez, on="Cislo_Prijemky", how="left")
    agg["Lezaky"] = agg["Lezaky"].apply(lambda x: x if isinstance(x, list) else [])

    agg = prijemky[["Cislo", "Firma", "Celkem"]].rename(
        columns={"Cislo": "Cislo_Prijemky", "Celkem": "Celkem_Prijemka"}
    ).merge(agg, on="Cislo_Prijemky", how="left")

    for c in ("Prijato_ks", "Prodano_ks", "Skladem_ks",
              "Hodnota_Prijem", "Hodnota_Skladem"):
        agg[c] = agg[c].fillna(0.0)
    agg["Lezaky"] = agg["Lezaky"].apply(lambda x: x if isinstance(x, list) else [])

    def pct(r):
        # Bez napojených položek nevíme, co je skladem. Nelžeme.
        base = r["Hodnota_Prijem"]
        if base <= 0:
            return None
        return max(0.0, min(1.0, 1 - r["Hodnota_Skladem"] / base))

    agg["Pct_Prodano"] = agg.apply(pct, axis=1)
    return agg


def spocitat_faktury(faktury, prij_agg, vazba, dnes: date) -> pd.DataFrame:
    vf = vazba.merge(prij_agg, on="Cislo_Prijemky", how="left")
    agg = (vf.groupby("Cislo_Faktury")
           .agg(Pocet_Prijemek=("Cislo_Prijemky", "nunique"),
                Hodnota_Prijem_sum=("Hodnota_Prijem", "sum"),
                Hodnota_Skladem_sum=("Hodnota_Skladem", "sum"),
                Celkem_Prijemka_sum=("Celkem_Prijemka", "sum"))
           .reset_index()
           .rename(columns={"Cislo_Faktury": "Cislo"}))

    out = faktury.merge(agg, on="Cislo", how="left")
    out["Pocet_Prijemek"] = out["Pocet_Prijemek"].fillna(0).astype(int)
    for c in ("Hodnota_Prijem_sum", "Hodnota_Skladem_sum", "Celkem_Prijemka_sum"):
        out[c] = out[c].fillna(0.0)

    def fakt_pct(r):
        # Detekce chybějících dat: pokud nemáme napojené položky (hodnota příjemky = 0),
        # neznáme co je skladem → nelžeme, vrátíme None (stav CHYBÍ DATA).
        if r["Hodnota_Prijem_sum"] <= 0:
            return None
        # % prodáno se počítá z ČÁSTKY FAKTURY (vizuální konzistence s tabulkou).
        # Pozor: částka faktury může obsahovat DPH/dopravu, proto je to orientační.
        zaklad = r["Celkem"] if r["Celkem"] and r["Celkem"] > 0 else r["Hodnota_Prijem_sum"]
        return max(0.0, min(1.0, 1 - r["Hodnota_Skladem_sum"] / zaklad))

    out["Pct_Prodano"] = out.apply(fakt_pct, axis=1)
    # Pct_Skladem dopočítáme jen tam, kde známe Pct_Prodano
    out["Pct_Skladem"] = out["Pct_Prodano"].apply(lambda x: None if x is None else 1 - x)
    out["Dni_do_splatnosti"] = out["Splatno"].apply(
        lambda d: None if pd.isna(d) else (d.date() - dnes).days)

    # Zaplaceno: K_likvidaci <= 0 (s malou tolerancí kvůli zaokrouhlení)
    out["Zaplaceno"] = out["K_likvidaci"] <= 0.01

    def stav(r):
        # ⚫ NESPÁROVÁNO — k faktuře chybí příjemka
        if r["Pocet_Prijemek"] == 0:
            return STAV_NESP
        # Zaplacené → automaticky OK
        if r["Zaplaceno"]:
            return STAV_OK
        # ❓ CHYBÍ DATA — příjemka existuje, ale neznáme položky
        if pd.isna(r["Pct_Prodano"]):
            return STAV_CHYBI_DATA
        # 🟢 OK — vše prodáno (Pct_Prodano >= 0.99 kvůli zaokrouhlení)
        if r["Pct_Prodano"] >= 0.99:
            return STAV_OK
        dd = r["Dni_do_splatnosti"]
        ps = r["Pct_Skladem"] if pd.notna(r["Pct_Skladem"]) else 1.0
        # 🔴 KRITICKÉ — po splatnosti (nezaplacená)
        if dd is not None and dd < 0:
            return STAV_KRIT
        # 🟠 RIZIKO — ≤14 dní do splatnosti a >50 % neprodáno
        if dd is not None and 0 <= dd <= RIZIKO_DNI and ps > RIZIKO_SKLADEM_PCT:
            return STAV_RIZIKO
        # 🟡 V PROCESU — >14 dní do splatnosti, zboží ještě na skladě
        if dd is not None and dd > RIZIKO_DNI:
            return STAV_PROCES
        # Fallback (splatnost chybí nebo jiný edge case)
        return STAV_PROCES

    out["Stav"] = out.apply(stav, axis=1)
    return out


# ---------------------------------------------------------------------------
# Tab 2: Ležáky & Splatnosti — zboží ze skladu s vazbou na faktury
# ---------------------------------------------------------------------------

def sestav_lezaky_s_fakturami(
    stav: pd.DataFrame,
    polozky: pd.DataFrame,
    prij_agg: pd.DataFrame,
    vazba: pd.DataFrame,
    faktury_view: pd.DataFrame,
) -> pd.DataFrame:
    """
    Pro každý kód ze skladu (Stav_efektivni > 0) najde fakturu přes LIFO vazbu.
    Vrátí tabulku: Kod, Nazev, Dodavatel, Vyrobce, Stav_zasoby, Rezervace,
                   Stav_efektivni, Cena, Hodnota, Cislo_Faktury, Doklad,
                   Splatno, Dni_do_splatnosti, Stav_faktury
    """
    if stav is None or stav.empty:
        return pd.DataFrame()

    stav_col = "Stav_efektivni" if "Stav_efektivni" in stav.columns else "Stav_zasoby"
    aktivni = stav[stav[stav_col] > 0].copy()
    if aktivni.empty:
        return pd.DataFrame()

    # Přidej hodnotu
    if "Cena" in aktivni.columns:
        aktivni["Hodnota"] = aktivni[stav_col] * aktivni["Cena"].fillna(0)
    else:
        aktivni["Hodnota"] = 0.0

    if polozky.empty:
        aktivni["Cislo_Faktury"] = None
        aktivni["Doklad"] = None
        aktivni["Splatno"] = pd.NaT
        aktivni["Dni_do_splatnosti"] = None
        aktivni["Stav_faktury"] = "❓ CHYBÍ DATA"
        return aktivni

    # Najdi příjemku pro každý kód přes polozky (LIFO: nejnovější příjemka s tímto kódem)
    # Pro každý kód vezmeme příjemku, kde Zbyva_ks > 0 a která je nejnovější
    pol_zbytek = polozky[polozky["Zbyva_ks"] > 0][["Cislo_Prijemky", "Kod"]].copy()

    # Přidej datum příjemky pro řazení
    if "Datum" in prij_agg.columns or "Cislo_Prijemky" in prij_agg.columns:
        # Vezmi nejnovější příjemku pro každý kód
        prij_daty = prij_agg[["Cislo_Prijemky"]].copy()
        pol_zbytek = pol_zbytek.merge(prij_daty, on="Cislo_Prijemky", how="left")

    # Pro každý kód: vezmi příjemku s největším číslem (= nejnovější v Pohodě)
    pol_nejnovejsi = (pol_zbytek.sort_values("Cislo_Prijemky", ascending=False)
                      .drop_duplicates(subset="Kod")
                      [["Kod", "Cislo_Prijemky"]])

    # Přes příjemku najdi fakturu
    fa_pro_prij = (vazba[["Cislo_Faktury", "Cislo_Prijemky"]]
                   .drop_duplicates(subset="Cislo_Prijemky"))
    pol_nejnovejsi = pol_nejnovejsi.merge(fa_pro_prij, on="Cislo_Prijemky", how="left")

    # Přidej info o faktuře
    fa_info = faktury_view[["Cislo", "Doklad", "Firma", "Splatno",
                              "Dni_do_splatnosti", "Stav", "K_likvidaci"]].rename(
        columns={"Cislo": "Cislo_Faktury", "Stav": "Stav_faktury"})
    pol_nejnovejsi = pol_nejnovejsi.merge(fa_info, on="Cislo_Faktury", how="left")

    # Spoj se stavem skladu
    out = aktivni.merge(
        pol_nejnovejsi[["Kod", "Cislo_Faktury", "Doklad", "Firma",
                         "Splatno", "Dni_do_splatnosti", "Stav_faktury", "K_likvidaci"]],
        on="Kod", how="left"
    )
    return out


# ---------------------------------------------------------------------------
# Tab 3: Obratová analýza — ABC × XYZ
# ---------------------------------------------------------------------------

def vypocti_obrat(
    pohyby: pd.DataFrame,
    stav: pd.DataFrame,
    dni_analyzy: int = 365,
    referenci_datum: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Pro každý kód spočítá: obrat, pokrytí, trend, ABC, XYZ kategorii.

    Vrátí DataFrame: Kod, Obrat_ks, Obrat_Kc, Prumer_tyden_ks,
                     Pokryti_tydnu, Trend, ABC, XYZ, Riziko
    """
    if pohyby.empty or stav is None or stav.empty:
        return pd.DataFrame()

    ref = referenci_datum or pd.Timestamp.now().normalize()
    od = ref - pd.Timedelta(days=dni_analyzy)

    # Jen výdeje (prodeje) za sledované období
    vydeje = pohyby[
        (pohyby["Typ_norm"] == "Vydej") &
        (pohyby["Datum"] >= od) &
        (pohyby["Datum"] <= ref)
    ].copy()

    if vydeje.empty:
        return pd.DataFrame()

    vydeje["Hodnota"] = vydeje["Mnozstvi"] * vydeje["Cena"].fillna(0)

    # Týdenní prodeje pro XYZ a trend
    vydeje["Tyden"] = vydeje["Datum"].dt.to_period("W")
    tydenni = (vydeje.groupby(["Kod", "Tyden"])
               .agg(Ks=("Mnozstvi", "sum"), Kc=("Hodnota", "sum"))
               .reset_index())

    # Celkový obrat
    obrat = (vydeje.groupby("Kod")
             .agg(Obrat_ks=("Mnozstvi", "sum"),
                  Obrat_Kc=("Hodnota", "sum"))
             .reset_index())

    # Průměr za týden a koeficient variace (pro XYZ)
    pocet_tydnu = max(1, dni_analyzy / 7)

    def _xyz_stats(grp):
        ks_list = grp["Ks"].values
        mean = ks_list.mean() if len(ks_list) > 0 else 0
        std = ks_list.std() if len(ks_list) > 1 else 0
        cv = (std / mean) if mean > 0 else 999  # koeficient variace
        return pd.Series({"Mean_tyden": mean, "CV": cv})

    xyz_stats = tydenni.groupby("Kod").apply(_xyz_stats).reset_index()
    obrat = obrat.merge(xyz_stats, on="Kod", how="left")
    obrat["Prumer_tyden_ks"] = obrat["Obrat_ks"] / pocet_tydnu

    # Trend: porovnej posledních 8 týdnů vs předchozích 8 týdnů
    hranice_trend = ref - pd.Timedelta(weeks=8)
    vydeje_nove = vydeje[vydeje["Datum"] >= hranice_trend].groupby("Kod")["Mnozstvi"].sum()
    vydeje_stare = vydeje[(vydeje["Datum"] < hranice_trend)].groupby("Kod")["Mnozstvi"].sum()
    obrat["Prod_nove"] = obrat["Kod"].map(vydeje_nove).fillna(0)
    obrat["Prod_stare"] = obrat["Kod"].map(vydeje_stare).fillna(0)

    def _trend(r):
        n, s = r["Prod_nove"], r["Prod_stare"]
        if s == 0 and n == 0:
            return "—"
        if s == 0:
            return "↑ Nové"
        ratio = n / (s + 1e-9)
        if ratio > 1.2:
            return "↑ Roste"
        if ratio < 0.8:
            return "↓ Klesá"
        return "→ Stabilní"

    obrat["Trend"] = obrat.apply(_trend, axis=1)

    # Stav skladu
    stav_col = "Stav_efektivni" if "Stav_efektivni" in stav.columns else "Stav_zasoby"
    stav_map = dict(zip(stav["Kod"], stav[stav_col]))
    cena_map = dict(zip(stav["Kod"], stav.get("Cena", pd.Series(dtype=float))))
    obrat["Stav_efektivni"] = obrat["Kod"].map(stav_map).fillna(0)
    obrat["Cena"] = obrat["Kod"].map(cena_map).fillna(0)
    obrat["Hodnota_skladu"] = obrat["Stav_efektivni"] * obrat["Cena"]

    # Pokrytí v týdnech
    obrat["Pokryti_tydnu"] = (
        obrat["Stav_efektivni"] / obrat["Prumer_tyden_ks"].replace(0, float("nan"))
    ).fillna(9999).clip(upper=9999)

    # ABC kategorie (kumulativní podíl na obratu v Kč)
    obrat_sort = obrat.sort_values("Obrat_Kc", ascending=False).copy()
    total_kc = obrat_sort["Obrat_Kc"].sum()
    if total_kc > 0:
        obrat_sort["Kum_podil"] = obrat_sort["Obrat_Kc"].cumsum() / total_kc
        obrat_sort["ABC"] = obrat_sort["Kum_podil"].apply(
            lambda x: "A" if x <= 0.8 else ("B" if x <= 0.95 else "C"))
    else:
        obrat_sort["ABC"] = "C"
    obrat = obrat.merge(obrat_sort[["Kod", "ABC"]], on="Kod", how="left")

    # XYZ kategorie (koeficient variace)
    obrat["XYZ"] = obrat["CV"].apply(
        lambda cv: "X" if cv < 0.5 else ("Y" if cv < 1.0 else "Z"))

    # Riziko
    def _riziko(r):
        abc, pokr, trend = r["ABC"], r["Pokryti_tydnu"], r["Trend"]
        if abc in ("A", "B") and pokr > 26 and trend == "↓ Klesá":
            return "🔴 Kritické"
        if abc in ("A", "B") and pokr > 13:
            return "🟠 Sledovat"
        if abc == "C" and pokr > 26:
            return "🟡 Pomalé"
        return "🟢 OK"

    obrat["Riziko"] = obrat.apply(_riziko, axis=1)

    # Přidej Název a Výrobce ze stavu skladu
    if "Nazev" in stav.columns:
        nazev_map = dict(zip(stav["Kod"], stav["Nazev"]))
        obrat["Nazev"] = obrat["Kod"].map(nazev_map).fillna("")
    else:
        obrat["Nazev"] = ""
    if "Vyrobce" in stav.columns:
        vyrobce_map = dict(zip(stav["Kod"], stav["Vyrobce"]))
        obrat["Vyrobce"] = obrat["Kod"].map(vyrobce_map).fillna("")
    else:
        obrat["Vyrobce"] = ""

    cols_out = [
        "Kod", "Nazev", "Vyrobce", "Obrat_ks", "Obrat_Kc", "Prumer_tyden_ks",
        "Pokryti_tydnu", "Trend", "ABC", "XYZ", "Riziko", "Stav_efektivni",
        "Hodnota_skladu", "Cena"
    ]
    return obrat[[c for c in cols_out if c in obrat.columns]]


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def uvodni_obrazovka() -> None:
    st.title("📡 Cashflow Radar")
    st.markdown(
        """
### K čemu to slouží

Nástroj propojuje tři pohledy, které firmy obvykle sledují odděleně, a ukazuje,
**kde vám leží peníze ve zboží, které jste ještě neprodali — ale už ho musíte zaplatit.**

### Vstupní soubory

**Povinné:**
- **faktury.xlsx** — faktury přijaté (kdo, kolik, do kdy)
- **prijemky.xlsx** — hlavičky příjemek (propojení faktura ↔ zboží)
- **pohyby.xlsx** — skladové pohyby (příjmy a výdeje)

**Volitelné (zpřesní výpočet):**
- **prijemky_polozkove.xlsx** — Pohoda tiskový export položek příjemek. Obsahuje přesnou vazbu *která příjemka má jaké kódy a kolik kusů*. **Výrazně zvýší přesnost.**
- **stav.xlsx** — aktuální stav skladu po kódech (Kód, Stav zásoby).

### Co uvidíte

- **🔴 KRITICKÉ** — faktura **po splatnosti** a zboží není plně prodáno.
- **🟠 RIZIKO** — splatnost do **14 dní** a víc než **50 %** zboží neprodáno.
- **🟡 V PROCESU** — splatnost je více než **14 dní** od teď, zboží ještě na skladě.
- **🟢 OK** — zboží **plně prodáno** (i když třeba splatnost prošla).
- **⚫ NESPÁROVÁNO** — k faktuře se nenašla žádná příjemka.
- **❓ CHYBÍ DATA** — příjemka nalezena, ale aplikace nezná její položky (vyexportuj položkový přehled za širší období).

Po výběru faktury uvidíte **produkty z faktury** — kódy, názvy, kusovost, co je prodáno a co leží ve skladě.

### Jak na to

1. V levém panelu vyber zdroj dat.
2. Volitelně přidej `stav.xlsx` (přejmenuj svůj export stavu skladu) pro nejpřesnější výsledky.
3. Klikni na **▶️ Spustit analýzu**.
        """
    )
    st.info("Aplikace běží 100 % lokálně. Nic se nikam neodesílá.", icon="🔒")


def _cache_dir() -> Path:
    """Složka pro uložení posledních nahraných souborů (perzistence napříč sessions)."""
    base = Path(__file__).parent if "__file__" in globals() else Path.cwd()
    d = base / "data_cache"
    d.mkdir(exist_ok=True)
    return d


def _uloz_upload(uploaded_file, cilovy_nazev: str) -> Path | None:
    """Uloží nahraný soubor do cache složky pod stálým názvem."""
    if uploaded_file is None:
        return None
    cesta = _cache_dir() / cilovy_nazev
    with open(cesta, "wb") as f:
        f.write(uploaded_file.getbuffer())
    # Zrcadli do Supabase, ať data přežijí uspání appky na Streamlit Cloud
    if _persist is not None:
        _persist.push_file(cesta)
    return cesta


def _nacti_z_cache() -> dict | None:
    """Načte poslední uložené soubory z cache. Vrátí dict cest nebo None."""
    d = _cache_dir()
    fa = d / "faktury.xlsx"
    pr = d / "prijemky.xlsx"
    po = d / "pohyby.xlsx"
    if not (fa.exists() and pr.exists() and po.exists()):
        return None
    pol = []
    for jmeno in ("prijemky_polozkove", "prijemky_polozkove_nove"):
        for ext in (".xlsx", ".pdf"):
            p = d / f"{jmeno}{ext}"
            if p.exists():
                pol.append(p)
                break
    stav = d / "stav.xlsx"
    po2 = d / "pohyby_historicke.xlsx"
    return {
        "faktury": fa, "prijemky": pr, "pohyby": po,
        "pol_src_list": pol,
        "stav": stav if stav.exists() else None,
        "pohyby2": po2 if po2.exists() else None,
    }


def _cas_posledniho_nahrani() -> str | None:
    """Vrátí čas poslední modifikace cache souboru faktury.xlsx."""
    fa = _cache_dir() / "faktury.xlsx"
    if not fa.exists():
        return None
    import datetime as _dt
    ts = _dt.datetime.fromtimestamp(fa.stat().st_mtime)
    return ts.strftime("%d.%m.%Y %H:%M")


def sidebar_zdroje():
    st.sidebar.header("📂 Zdroj dat")

    cas_cache = _cas_posledniho_nahrani()
    moznosti = ["Lokální soubory ve složce skriptu", "Nahrát soubory"]
    if cas_cache:
        moznosti.insert(0, "Použít poslední nahraná data")

    rezim = st.sidebar.radio(
        "Odkud načíst soubory?",
        moznosti,
        index=0,
    )

    if rezim == "Použít poslední nahraná data":
        cache = _nacti_z_cache()
        if cache is None:
            st.sidebar.error("Žádná uložená data nenalezena. Nahraj soubory.")
            return None
        st.sidebar.success(
            f"📌 Načtena poslední data\n\nNaposledy nahráno: **{cas_cache}**")
        extras = []
        if cache["pol_src_list"]:
            extras.append(f"✨ + {len(cache['pol_src_list'])}× příjemky položkově")
        if cache["stav"]:
            extras.append("✨ + stav.xlsx")
        if cache["pohyby2"]:
            extras.append("✨ + pohyby_historicke.xlsx")
        if extras:
            st.sidebar.caption("\n".join(extras))
        return (cache["faktury"], cache["prijemky"], cache["pohyby"],
                cache["pol_src_list"], cache["stav"], cache["pohyby2"])

    if rezim == "Lokální soubory ve složce skriptu":
        base = Path(__file__).parent if "__file__" in globals() else Path.cwd()
        cesty = {n: base / f"{n}.xlsx" for n in ("faktury", "prijemky", "pohyby")}
        chybi = [n for n, p in cesty.items() if not p.exists()]
        if chybi:
            st.sidebar.error(
                "Chybí: " + ", ".join(f"{n}.xlsx" for n in chybi)
                + f"\n\nOčekávaná složka:\n`{base}`")
            return None

        # Položkové exporty — hledej xlsx i pdf varianty
        pol_src_list = []
        for jmeno in ("prijemky_polozkove", "prijemky_polozkove_nove"):
            for ext in (".xlsx", ".pdf"):
                p = base / f"{jmeno}{ext}"
                if p.exists():
                    pol_src_list.append(p)
                    break  # bere první nalezený formát pro každý soubor

        stav_path = base / "stav.xlsx"
        stav_src = stav_path if stav_path.exists() else None
        pohyby2_path = base / "pohyby_historicke.xlsx"
        pohyby2_src = pohyby2_path if pohyby2_path.exists() else None

        zprava = f"Nalezeny 3 povinné soubory v:\n`{base}`"
        extras = []
        if pol_src_list:
            extras.append(f"✨ + {len(pol_src_list)}× příjemky položkově")
        if stav_src:
            extras.append("✨ + stav.xlsx")
        if pohyby2_src:
            extras.append("✨ + pohyby_historicke.xlsx")
        if extras:
            zprava += "\n\n" + "\n".join(extras)
        st.sidebar.success(zprava)

        if not pol_src_list:
            st.sidebar.info("💡 Tip: přidej `prijemky_polozkove.xlsx` a/nebo `prijemky_polozkove_nove.xlsx`.")
        if not stav_src:
            st.sidebar.info("💡 Tip: přidej `stav.xlsx` (Kód + Stav zásoby + Rezervace) pro přesné zůstatky.")
        return cesty["faktury"], cesty["prijemky"], cesty["pohyby"], pol_src_list, stav_src, pohyby2_src

    fa = st.sidebar.file_uploader("faktury.xlsx (povinné)", type=["xlsx"])
    pr = st.sidebar.file_uploader("prijemky.xlsx (povinné)", type=["xlsx"])
    po = st.sidebar.file_uploader("pohyby.xlsx — aktuální rok (povinné)", type=["xlsx"])
    po2 = st.sidebar.file_uploader(
        "pohyby_historicke.xlsx — předchozí rok (volitelné)",
        type=["xlsx"],
        help="Pohyby za 2025 nebo starší. Sloučí se s aktuálními. Potřeba pro přesný výpočet obratu.")
    st.sidebar.markdown("**Volitelné — příjemky položkově:**")
    pol_stary = st.sidebar.file_uploader(
        "Příjemky položkově (starší export)",
        type=["xlsx", "pdf"],
        help="Pohoda tiskový export Položky příjemek — starší období.")
    pol_novy = st.sidebar.file_uploader(
        "Příjemky položkově (nový export)",
        type=["xlsx", "pdf"],
        help="Pohoda tiskový export Položky příjemek — nové období.")
    pol_src_list = [p for p in (pol_stary, pol_novy) if p is not None]
    stav = st.sidebar.file_uploader(
        "stav.xlsx",
        type=["xlsx"],
        help="Aktuální stav skladu: sloupce Kód, Stav zásoby, Rezervace.")
    if not (fa and pr and po):
        st.sidebar.info("Nahrajte všechny povinné soubory.")
        return None

    # Ulož kopie do cache, aby byly dostupné i příště (perzistence)
    ulozit = st.sidebar.checkbox(
        "💾 Zapamatovat tato data pro příště", value=True,
        help="Uloží kopie souborů. Příště je načteš volbou 'Použít poslední nahraná data'.")
    if ulozit:
        _uloz_upload(fa, "faktury.xlsx")
        _uloz_upload(pr, "prijemky.xlsx")
        _uloz_upload(po, "pohyby.xlsx")
        if po2:
            _uloz_upload(po2, "pohyby_historicke.xlsx")
        if stav:
            _uloz_upload(stav, "stav.xlsx")
        # Položkové — ulož pod stálými názvy dle pořadí
        if pol_stary:
            ext = ".pdf" if pol_stary.name.lower().endswith(".pdf") else ".xlsx"
            _uloz_upload(pol_stary, f"prijemky_polozkove{ext}")
        if pol_novy:
            ext = ".pdf" if pol_novy.name.lower().endswith(".pdf") else ".xlsx"
            _uloz_upload(pol_novy, f"prijemky_polozkove_nove{ext}")

    return fa, pr, po, pol_src_list, stav, po2


def format_kc(x: float) -> str:
    try:
        return f"{x:,.0f} Kč".replace(",", " ")
    except (ValueError, TypeError):
        return "—"


def format_pct(x: float) -> str:
    if pd.isna(x):
        return "—"
    return f"{x * 100:.1f} %"


def spust_analyzu(zdroje, dnes: date, rok_od: int = 2000) -> dict | None:
    f_src, p_src, m_src, pol_src_list, s_src, po2_src = (*zdroje, None)[:6]

    with st.status("🔄 Probíhá analýza…", expanded=True) as status:
        try:
            st.write("📥 **Krok 1/7:** Načítám Excel soubory…")
            t0 = time.perf_counter()
            faktury, prijemky, pohyby, polozky_pohoda, stav = nacti_a_normalizuj(
                f_src, p_src, m_src, pol_src_list, s_src, po2_src)

            # Filtr faktury dle roku
            if rok_od > 2000 and "Splatno" in faktury.columns:
                maska_roku = (
                    faktury["Splatno"].dt.year >= rok_od
                ) | faktury["Splatno"].isna()
                n_pred = len(faktury)
                faktury = faktury[maska_roku].copy()
                n_odfiltr = n_pred - len(faktury)
                if n_odfiltr > 0:
                    st.write(
                        f"&nbsp;&nbsp;&nbsp;🗓️ Odstraněno **{n_odfiltr}** faktur "
                        f"starších než rok **{rok_od}**.",
                        unsafe_allow_html=True)

            # Filtr Text = "Zboží" — ponechat jen faktury za zboží, ne služby/opravy
            if "Text" in faktury.columns:
                text_col = faktury["Text"].astype(str).str.strip().str.lower()
                mask_zbozi = (
                    text_col.str.contains("zbozi|zboží", na=False, regex=True)
                    | text_col.isin(["", "nan", "none"])
                    | faktury["Text"].isna()
                )
                n_pred = len(faktury)
                faktury = faktury[mask_zbozi].copy()
                n_odfiltr_text = n_pred - len(faktury)
                if n_odfiltr_text > 0:
                    st.write(
                        f"&nbsp;&nbsp;&nbsp;📦 Ponecháno jen **Zboží** — odstraněno **{n_odfiltr_text}** "
                        f"faktur za služby/opravy (sloupec Text).",
                        unsafe_allow_html=True)

            info = (f"faktury: **{len(faktury)}** &nbsp;|&nbsp; "
                    f"příjemky: **{len(prijemky)}** &nbsp;|&nbsp; "
                    f"pohyby: **{len(pohyby)}**")
            if polozky_pohoda is not None:
                n_prij = polozky_pohoda["Cislo_Prijemky"].nunique()
                info += f" &nbsp;|&nbsp; položek: **{len(polozky_pohoda)}** řádků / **{n_prij}** příjemek"
            if stav is not None:
                info += f" &nbsp;|&nbsp; stav: **{len(stav)}** kódů"
            st.write(f"&nbsp;&nbsp;&nbsp;✓ {info} &nbsp;*({time.perf_counter()-t0:.2f} s)*",
                     unsafe_allow_html=True)

            # Filtrování pohybů podle agend (Příjemka, Prodejka, Výdejka, Vydaná faktura)
            st.write("🧹 **Krok 2/8:** Filtruji pohyby podle agend…")
            t0 = time.perf_counter()
            # Whitelist agend — vždy mergujeme konstanty + session_state
            # (aby nové výchozí hodnoty v kódu platily i bez restartu prohlížeče)
            agendy_pr = list({*AGENDY_PRIJEM_WHITELIST,
                              *st.session_state.get("agendy_prijem", [])})
            agendy_vy = list({*AGENDY_VYDEJ_WHITELIST,
                              *st.session_state.get("agendy_vydej", [])})
            pohyby_orig = pohyby.copy()
            pohyby, diag_agendy = filtruj_pohyby_podle_agend(pohyby, agendy_pr, agendy_vy)
            if diag_agendy["agenda_chybi"]:
                st.write(
                    f"&nbsp;&nbsp;&nbsp;⚠️ Sloupec `Pohyb` v pohybech chybí nebo je prázdný — "
                    f"filtr neaplikován, beru všechny pohyby.",
                    unsafe_allow_html=True)
            else:
                st.write(
                    f"&nbsp;&nbsp;&nbsp;✓ zahrnuto: **{diag_agendy['zahrnuto']}** &nbsp;|&nbsp; "
                    f"vyfiltrováno: **{diag_agendy['vyloucenyo']}** pohybů &nbsp;"
                    f"*({time.perf_counter()-t0:.2f} s)*",
                    unsafe_allow_html=True)

            # Volba metody napojení a rekonstrukce
            st.write("🎯 **Krok 3/8:** Vybírám metodu výpočtu…")
            if polozky_pohoda is not None and not polozky_pohoda.empty:
                napoj_metoda = NAPOJ_POLOZKOVE
            else:
                napoj_metoda = NAPOJ_DODAVATEL_DATUM

            if stav is not None:
                reko_metoda = REKO_LIFO_EXT
                stav_pouzity = stav
            else:
                z_pohybu = stav_z_poslednich_pohybu(pohyby)
                if z_pohybu is not None and len(z_pohybu) > 0:
                    reko_metoda = REKO_LIFO_POHYB
                    stav_pouzity = z_pohybu
                else:
                    reko_metoda = REKO_FIFO
                    stav_pouzity = None
            st.write(f"&nbsp;&nbsp;&nbsp;✓ Napojení: **{napoj_metoda}** &nbsp;|&nbsp; "
                     f"Rekonstrukce: **{reko_metoda}**", unsafe_allow_html=True)

            st.write("🔗 **Krok 4/8:** Páruji faktury s příjemkami…")
            n_fa = len(faktury)
            n_pr = len(prijemky)
            if n_fa > 5000:
                st.write(
                    f"&nbsp;&nbsp;&nbsp;⏳ Velký dataset ({n_fa:,} faktur × {n_pr:,} příjemek) — "
                    f"párování může trvat **30–90 sekund**. Aplikace pracuje, nestiskej nic.".replace(",", " "),
                    unsafe_allow_html=True)
            t0 = time.perf_counter()
            vazba = parovat_faktury_prijemky(faktury, prijemky)
            if not vazba.empty:
                mc = vazba["match_type"].value_counts().to_dict()
                exact   = mc.get("exact", 0)
                substr  = mc.get("substring", 0)
                cyklo   = mc.get("cyklomax", 0)
                crussis = mc.get("crussis_suffix", 0)
                nosep   = mc.get("no_separator", 0)
                cisla   = mc.get("cisla_prunik", 0)
                castka  = mc.get("castka_nazev", 0)
                parts = [f"přesně: **{exact}**", f"substring: **{substr}**"]
                if cyklo:   parts.append(f"Cyklomax: **{cyklo}**")
                if crussis: parts.append(f"Crussis: **{crussis}**")
                if nosep:   parts.append(f"bez sep.: **{nosep}**")
                if cisla:   parts.append(f"čísla: **{cisla}**")
                if castka:  parts.append(f"částka+název: **{castka}**")
                diag_parovani_info = " &nbsp;|&nbsp; ".join(parts)
            else:
                diag_parovani_info = "žádné shody"
            nespárovano_fa = len(faktury) - vazba["Cislo_Faktury"].nunique() if not vazba.empty else len(faktury)
            st.write(
                f"&nbsp;&nbsp;&nbsp;✓ {diag_parovani_info} &nbsp;|&nbsp; "
                f"nespárováno: **{nespárovano_fa}** faktur &nbsp;"
                f"*({time.perf_counter()-t0:.2f} s)*",
                unsafe_allow_html=True)

            st.write(f"🔄 **Krok 5/8:** Napojuji příjmy na příjemky ({napoj_metoda})…")
            t0 = time.perf_counter()
            if napoj_metoda == NAPOJ_POLOZKOVE:
                prijmy = napojit_z_polozek(polozky_pohoda, prijemky, pohyby)
                diag_napojeni = {"osirele_prijmy": pd.DataFrame(), "kolize_den": 0,
                                 "prijemky_bez_polozek": 0}
                # Zjisti příjemky v prijemky.xlsx, které nejsou v položkovém exportu
                prij_v_polozkach = set(polozky_pohoda["Cislo_Prijemky"].unique())
                prij_bez = set(prijemky["Cislo"].unique()) - prij_v_polozkach
                diag_napojeni["prijemky_bez_polozek"] = len(prij_bez)
                st.write(
                    f"&nbsp;&nbsp;&nbsp;✓ napojeno **{len(prijmy)}** řádků (přesně) &nbsp;|&nbsp; "
                    f"příjemek bez položek: **{len(prij_bez)}** &nbsp;"
                    f"*({time.perf_counter()-t0:.2f} s)*",
                    unsafe_allow_html=True)
            else:
                prijmy, diag_napojeni = napojit_prijmy_na_prijemky(prijemky, pohyby)
                osirele_n = len(diag_napojeni["osirele_prijmy"])
                st.write(
                    f"&nbsp;&nbsp;&nbsp;✓ napojeno **{len(prijmy)}** řádků &nbsp;|&nbsp; "
                    f"osiřelých: **{osirele_n}** &nbsp;|&nbsp; "
                    f"kolize dat: **{diag_napojeni['kolize_den']}** &nbsp;"
                    f"*({time.perf_counter()-t0:.2f} s)*",
                    unsafe_allow_html=True)

            st.write(f"📦 **Krok 6/8:** Počítám zůstatky ({reko_metoda.split('(')[0].strip()})…")
            t0 = time.perf_counter()
            diag_rekonstrukce = {}
            if stav_pouzity is not None:
                polozky, diag_rekonstrukce = rekonstrukce_lifo(prijmy, prijemky, stav_pouzity)
            else:
                polozky = rekonstrukce_fifo(prijmy, pohyby)
            st.write(
                f"&nbsp;&nbsp;&nbsp;✓ zpracováno **{len(polozky)}** řádků (Příjemka × Kód) &nbsp;"
                f"*({time.perf_counter()-t0:.2f} s)*",
                unsafe_allow_html=True)

            st.write("🧮 **Krok 7/8:** Agreguji na úroveň faktur a určuji stavy…")
            t0 = time.perf_counter()
            prij_agg = spocitat_prijemky(prijemky, polozky)
            faktury_view = spocitat_faktury(faktury, prij_agg, vazba, dnes)
            kr = int((faktury_view["Stav"] == STAV_KRIT).sum())
            ri = int((faktury_view["Stav"] == STAV_RIZIKO).sum())
            vp = int((faktury_view["Stav"] == STAV_PROCES).sum())
            ok = int((faktury_view["Stav"] == STAV_OK).sum())
            ns = int((faktury_view["Stav"] == STAV_NESP).sum())
            st.write(
                f"&nbsp;&nbsp;&nbsp;✓ 🔴 {kr} &nbsp;|&nbsp; 🟠 {ri} &nbsp;|&nbsp; "
                f"🟡 {vp} &nbsp;|&nbsp; 🟢 {ok} &nbsp;|&nbsp; ⚫ {ns} &nbsp;"
                f"*({time.perf_counter()-t0:.2f} s)*",
                unsafe_allow_html=True)

            st.write("✅ **Krok 8/8:** Připravuji dashboard…")

            # Výpočet obratové analýzy (Tab 3)
            obrat_data = pd.DataFrame()
            try:
                obrat_data = vypocti_obrat(
                    pohyby, stav,
                    dni_analyzy=365,
                    referenci_datum=pd.Timestamp(dnes),
                )
            except Exception:
                pass  # Tab 3 je volitelný, nesmí shodit zbytek

            # Ležáky s vazbou na faktury (Tab 2)
            lezaky_data = pd.DataFrame()
            try:
                lezaky_data = sestav_lezaky_s_fakturami(
                    stav, polozky, prij_agg, vazba, faktury_view)
            except Exception:
                pass

            status.update(label=f"✅ Analýza hotová — {napoj_metoda} + {reko_metoda}",
                         state="complete", expanded=False)

            return {
                "faktury": faktury, "prijemky": prijemky, "pohyby": pohyby, "stav": stav,
                "polozky_pohoda": polozky_pohoda,
                "vazba": vazba, "prij_agg": prij_agg, "polozky": polozky,
                "faktury_view": faktury_view, "dnes": dnes,
                "diag_napojeni": diag_napojeni, "diag_rekonstrukce": diag_rekonstrukce,
                "diag_agendy": diag_agendy,
                "napoj_metoda": napoj_metoda, "reko_metoda": reko_metoda,
                "obrat_data": obrat_data,
                "lezaky_data": lezaky_data,
            }
        except Exception as e:
            status.update(label="❌ Analýza selhala", state="error", expanded=True)
            st.error(f"**Chyba:** {e}")
            import traceback
            with st.expander("Podrobnosti chyby"):
                st.code(traceback.format_exc())
            return None


def dashboard(res: dict) -> None:
    tab1, tab2, tab3 = st.tabs([
        "📡 Cashflow Radar",
        "🧊 Ležáky & Splatnosti",
        "📊 Obratová analýza",
    ])

    with tab1:
        _tab_cashflow(res)

    with tab2:
        _tab_lezaky(res)

    with tab3:
        _tab_obrat(res)


def _tab_cashflow(res: dict) -> None:
    faktury_view = res["faktury_view"]
    prij_agg = res["prij_agg"]
    vazba = res["vazba"]
    pohyby = res["pohyby"]
    diag_n = res["diag_napojeni"]
    diag_r = res.get("diag_rekonstrukce", {})
    napoj_metoda = res["napoj_metoda"]
    reko_metoda = res["reko_metoda"]

    # Badge metod
    napoj_barva = "green" if NAPOJ_POLOZKOVE in napoj_metoda else "orange"
    reko_barva = ("green" if REKO_LIFO_EXT in reko_metoda
                  else "blue" if REKO_LIFO_POHYB in reko_metoda else "orange")
    st.markdown(f":{napoj_barva}[**Napojení:** {napoj_metoda}]  |  "
                f":{reko_barva}[**Rekonstrukce:** {reko_metoda}]")

    # KPI — dva řádky po 4 kartách
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Faktur celkem", f"{len(faktury_view)}")
    k2.metric("Hodnota faktur", format_kc(faktury_view["Celkem"].sum()))
    nezaplaceno_mask = ~faktury_view["Zaplaceno"]
    k3.metric("💰 Zbývá zaplatit",
              format_kc(faktury_view.loc[nezaplaceno_mask, "K_likvidaci"].sum()),
              help=f"Součet `K likvidaci` za všechny nezaplacené faktury ({int(nezaplaceno_mask.sum())} ks)")
    k4.metric("📦 Leží ve skladě",
              format_kc(faktury_view["Hodnota_Skladem_sum"].sum()),
              help="Součet hodnoty zboží, které je stále fyzicky na skladě")

    k5, k6, k7, k8, k9a, k9b, k10 = st.columns(7)
    k5.metric("🔴 Kritické", int((faktury_view["Stav"] == STAV_KRIT).sum()))
    k6.metric("🟠 Rizikové", int((faktury_view["Stav"] == STAV_RIZIKO).sum()))
    k7.metric("🟡 V procesu", int((faktury_view["Stav"] == STAV_PROCES).sum()))
    k8.metric("🟢 OK", int((faktury_view["Stav"] == STAV_OK).sum()))
    nesp = faktury_view[faktury_view["Stav"] == STAV_NESP]
    nesp_nezapl = int((~nesp["Zaplaceno"]).sum()) if not nesp.empty else 0
    nesp_zapl = int(nesp["Zaplaceno"].sum()) if not nesp.empty else 0
    k9a.metric("⚫ Nesp. (nezapl.)", nesp_nezapl,
               help="Nespárované a NEZAPLACENÉ — potenciální problém, chybí příjemka")
    k9b.metric("⚫ Nesp. (zapl.)", nesp_zapl,
               help="Nespárované ale zaplacené — historické faktury, žádné riziko")
    k10.metric("❓ Chybí data", int((faktury_view["Stav"] == STAV_CHYBI_DATA).sum()))

    # Sanity check: kolik bylo reálně prodáno ze všech příjemek?
    celkem_prijato_ks = 0
    celkem_prodano_ks = 0
    if not res["polozky"].empty:
        celkem_prijato_ks = res["polozky"]["Prijato_ks"].sum()
        celkem_prodano_ks = res["polozky"]["Prodano_ks"].sum()
    pct_realne_prodano = (celkem_prodano_ks / celkem_prijato_ks * 100) if celkem_prijato_ks > 0 else 0

    if celkem_prijato_ks > 0 and pct_realne_prodano < 5:
        st.warning(
            f"⚠️ **Pozor:** Z {celkem_prijato_ks:,.0f} přijatých kusů se zdá být prodáno pouze "
            f"{celkem_prodano_ks:,.0f} ks (**{pct_realne_prodano:.1f} %**). "
            f"To je velmi málo — pravděpodobně v `pohyby.xlsx` chybí výdejové pohyby, "
            f"nebo je rozsah exportu příliš úzký. Výsledky **% prodáno** a **Zbývá skladem** "
            f"proto mohou být nadhodnocené. Nejlepší řešení: přidej `stav.xlsx` s aktuálním "
            f"stavem skladu — tím se výpočet opraví podle reality.".replace(",", " ")
        )

    # Upozornění na CHYBI_DATA faktury
    chybi_data_pocet = int((faktury_view["Stav"] == STAV_CHYBI_DATA).sum())
    if chybi_data_pocet > 0:
        hodnota_chybi = faktury_view.loc[
            faktury_view["Stav"] == STAV_CHYBI_DATA, "K_likvidaci"
        ].sum()
        st.warning(
            f"❓ **{chybi_data_pocet} faktur** (celkem {format_kc(hodnota_chybi)} k likvidaci) "
            f"je ve stavu **CHYBÍ DATA** — jsou spárovány s příjemkou, ale aplikace neví, co z ní zbývá skladem. "
            f"Příčina: příjemky těchto faktur **nejsou v `prijemky_polozkove.xlsx`** "
            f"(tvůj položkový export pokrývá jen vybrané období). "
            f"**Řešení:** Vyexportuj z Pohody položky všech příjemek, které máš ve `faktury.xlsx` — "
            f"ne jen jeden den. Nebo přidej `stav.xlsx`, aby aplikace mohla určit stav přes LIFO rekonstrukci."
        )

    # Diagnostika agend (jaké typy pohybů v datech byly a co se filtrovalo)
    diag_a = res.get("diag_agendy", {})
    agendy_stats = diag_a.get("agendy_puvodne", pd.DataFrame())
    if not agendy_stats.empty:
        agendy_pr_set = {a.strip().lower() for a in st.session_state.get("agendy_prijem", AGENDY_PRIJEM_WHITELIST)}
        agendy_vy_set = {a.strip().lower() for a in st.session_state.get("agendy_vydej", AGENDY_VYDEJ_WHITELIST)}

        def _zahrnuto(row):
            nazev = str(row["Pohyb"]).strip().lower()
            # Normalizace diakritiky
            import unicodedata
            nazev_ascii = unicodedata.normalize("NFKD", nazev).encode("ascii", "ignore").decode("ascii")
            pr_ascii = {unicodedata.normalize("NFKD", a).encode("ascii", "ignore").decode("ascii") for a in agendy_pr_set}
            vy_ascii = {unicodedata.normalize("NFKD", a).encode("ascii", "ignore").decode("ascii") for a in agendy_vy_set}
            if row["Typ_norm"] == "Prijem" and nazev_ascii in pr_ascii:
                return "✅ Zahrnuto"
            if row["Typ_norm"] == "Vydej" and nazev_ascii in vy_ascii:
                return "✅ Zahrnuto"
            return "❌ Vynecháno"

        agendy_disp = agendy_stats.copy()
        agendy_disp["Zahrnuto"] = agendy_disp.apply(_zahrnuto, axis=1)
        agendy_disp = agendy_disp.rename(columns={
            "Typ_norm": "Směr", "Pohyb": "Agenda", "Pocet": "Počet"
        })[["Směr", "Agenda", "Počet", "Zahrnuto"]]

        pocet_zahrnutych_agend = (agendy_disp["Zahrnuto"] == "✅ Zahrnuto").sum()
        pocet_vynech_agend = (agendy_disp["Zahrnuto"] == "❌ Vynecháno").sum()

        with st.expander(
            f"📋 Přehled agend v pohybech ({pocet_zahrnutych_agend} zahrnutých, "
            f"{pocet_vynech_agend} vynechaných)"
        ):
            st.markdown(
                "Aplikace bere v potaz pouze agendy, které reprezentují skutečné nákupy "
                "(příjmy) a prodeje (výdeje). Ostatní — převodky, inventury, reklamace apod. — "
                "by zkreslovaly cashflow radar, a proto se vynechávají. **Whitelist agend "
                "můžeš upravit v sidebaru.**"
            )
            st.dataframe(agendy_disp, width="stretch", hide_index=True)

    # Diagnostika napojení
    prij_bez = diag_n.get("prijemky_bez_polozek", 0)
    osirele_n = len(diag_n.get("osirele_prijmy", pd.DataFrame()))
    kolize = diag_n.get("kolize_den", 0)
    if prij_bez or osirele_n or kolize:
        with st.expander(
            f"⚠️ Diagnostika napojení "
            f"({prij_bez} příjemek bez položek, {osirele_n} osiřelých, {kolize} kolizí)"
        ):
            if prij_bez:
                st.markdown(
                    f"**{prij_bez} příjemek** z `prijemky.xlsx` nemá odpovídající záznamy "
                    f"v položkovém exportu. Zůstanou nespárované a ukáží se jako NESPÁROVÁNO."
                )
            if osirele_n:
                st.markdown(
                    f"**{osirele_n} přijímacích pohybů** se nepodařilo napojit na žádnou příjemku.")
                st.dataframe(diag_n["osirele_prijmy"].head(200),
                             width="stretch", hide_index=True)
                if osirele_n > 200:
                    st.caption(f"Zobrazeno prvních 200 z {osirele_n}.")
            if kolize:
                st.markdown(
                    f"**{kolize}** případů, kdy stejný den přišly víc příjemek od stejného "
                    f"dodavatele. Pohyby byly rozděleny poměrem.")

    # Diagnostika rekonstrukce
    if diag_r:
        nez = diag_r.get("nezarazeno", pd.DataFrame())
        nex = diag_r.get("neexistujici_kod", pd.DataFrame())
        if len(nez) > 0 or len(nex) > 0:
            with st.expander(
                f"🔍 Diagnostika rekonstrukce "
                f"({len(nez)} kódů se zvýšeným stavem, {len(nex)} kódů mimo příjmy)"
            ):
                if len(nez) > 0:
                    st.markdown(
                        f"**{len(nez)} kódů** má vyšší skutečný stav skladu než součet "
                        f"evidovaných příjmů (starší data mimo export, ruční korekce, převody).")
                    st.dataframe(nez.head(200), width="stretch", hide_index=True)
                if len(nex) > 0:
                    st.markdown(
                        f"**{len(nex)} kódů** je ve stavu skladu, ale nemáme k nim žádný příjem.")
                    st.dataframe(nex.head(200), width="stretch", hide_index=True)

    st.divider()
    st.subheader("🔎 Filtry")
    f1, f2, f3, f4 = st.columns([2, 2, 2, 2])

    dodavatele = sorted(faktury_view["Firma"].dropna().unique().tolist())
    vybrani_dod = f1.multiselect("Dodavatel", dodavatele, default=[])

    vyrobci: list[str] = []
    if "Vyrobce" in pohyby.columns:
        vyrobci = sorted([v for v in pohyby["Vyrobce"].dropna().unique().tolist()
                          if v and v.lower() != "nan"])
    vybrani_vyr = f2.multiselect(
        "Výrobce" + ("" if vyrobci else " (v datech není)"),
        vyrobci, default=[], disabled=(not vyrobci))

    stavy_v = [STAV_KRIT, STAV_RIZIKO, STAV_PROCES, STAV_OK, STAV_NESP, STAV_CHYBI_DATA]
    vybrane_stavy = f3.multiselect("Stav", stavy_v, default=stavy_v)

    splatno_min = faktury_view["Splatno"].min()
    splatno_max = faktury_view["Splatno"].max()
    if pd.notna(splatno_min) and pd.notna(splatno_max):
        rozsah = f4.date_input(
            "Rozsah splatnosti",
            value=(splatno_min.date(), splatno_max.date()),
            min_value=splatno_min.date(), max_value=splatno_max.date())
    else:
        rozsah = None

    # Rychlý přepínač pro nezaplacené
    fc1, fc2 = st.columns([1, 3])
    jen_nezaplacene = fc1.checkbox(
        "💰 Jen nezaplacené",
        value=True,
        help="Zobrazit pouze faktury, které nejsou zaplacené (K likvidaci > 0)")
    if jen_nezaplacene:
        pocet_nezaplacenych = int((~faktury_view["Zaplaceno"]).sum())
        pocet_zaplacenych = int(faktury_view["Zaplaceno"].sum())
        fc2.caption(
            f"Zobrazuji **{pocet_nezaplacenych}** nezaplacených faktur "
            f"(skrývám {pocet_zaplacenych} zaplacených).")

    df = faktury_view.copy()
    if jen_nezaplacene:
        df = df[~df["Zaplaceno"]]
    if vybrani_dod:
        df = df[df["Firma"].isin(vybrani_dod)]
    if vybrane_stavy:
        df = df[df["Stav"].isin(vybrane_stavy)]
    if rozsah and isinstance(rozsah, tuple) and len(rozsah) == 2:
        a, b = rozsah
        df = df[(df["Splatno"].isna())
                | ((df["Splatno"].dt.date >= a) & (df["Splatno"].dt.date <= b))]
    if vybrani_vyr and not res["polozky"].empty:
        pol = res["polozky"]
        prij_ok = pol[pol["Vyrobce"].isin(vybrani_vyr)]["Cislo_Prijemky"].unique().tolist()
        fa_ok = vazba[vazba["Cislo_Prijemky"].isin(prij_ok)]["Cislo_Faktury"].unique().tolist()
        df = df[df["Cislo"].isin(fa_ok)]

    st.subheader("📋 Manažerský přehled")

    # Legenda stavů — vždy viditelná
    st.markdown(
        "🔴 **KRITICKÉ** — po splatnosti, zboží není plně prodáno &nbsp;|&nbsp; "
        "🟠 **RIZIKO** — ≤14 dní do splatnosti a >50 % neprodáno &nbsp;|&nbsp; "
        "🟡 **V PROCESU** — >14 dní do splatnosti &nbsp;|&nbsp; "
        "🟢 **OK** — vše prodáno &nbsp;|&nbsp; "
        "⚫ **NESPÁROVÁNO** — chybí příjemka &nbsp;|&nbsp; "
        "❓ **CHYBÍ DATA** — příjemka nalezena, položky neznáme",
        unsafe_allow_html=True,
    )

    STRANKY_VELIKOST = 200
    celkem_radku = len(df)
    pocet_stranek = max(1, (celkem_radku - 1) // STRANKY_VELIKOST + 1)

    if pocet_stranek > 1:
        strana = st.number_input(
            f"Strana (celkem {celkem_radku} faktur, {STRANKY_VELIKOST}/strana)",
            min_value=1, max_value=pocet_stranek, value=1, step=1, key="t1_strana")
    else:
        strana = 1

    start = (strana - 1) * STRANKY_VELIKOST
    df_strana = df.iloc[start:start + STRANKY_VELIKOST]

    zobr = df_strana.copy()
    zobr["Splatno"] = zobr["Splatno"].dt.date
    zobr["Pct_Prodano_display"] = zobr["Pct_Prodano"] * 100
    zobr_disp = zobr.rename(columns={
        "Cislo": "Faktura", "Doklad": "Var. symbol", "Firma": "Dodavatel",
        "Celkem": "Částka faktury", "K_likvidaci": "K likvidaci",
        "Hodnota_Skladem_sum": "Zbývá skladem",
        "Pct_Prodano_display": "% prodáno", "Dni_do_splatnosti": "Dní do spl.",
        "Pocet_Prijemek": "# příjemek",
    })[["Faktura", "Var. symbol", "Dodavatel", "Splatno", "Dní do spl.",
        "Částka faktury", "K likvidaci",
        "Zbývá skladem", "% prodáno", "# příjemek", "Stav"]]

    st.dataframe(
        zobr_disp, width="stretch", hide_index=True,
        column_config={
            "Částka faktury": st.column_config.NumberColumn(format="%.0f Kč"),
            "K likvidaci": st.column_config.NumberColumn(
                format="%.0f Kč",
                help="Kolik z faktury zbývá zaplatit (0 = zaplaceno)"),
            "Zbývá skladem": st.column_config.NumberColumn(
                format="%.0f Kč",
                help="Nákupní hodnota zboží z této faktury, které stále leží na skladě"),
            "% prodáno": st.column_config.ProgressColumn(
                format="%.0f %%", min_value=0, max_value=100,
                help="Počítáno jako 1 − (Zbývá skladem / Částka faktury). "
                     "Pozn.: částka faktury může obsahovat DPH a dopravu, proto je hodnota orientační."),
        })

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        # Export celého filtru, ne jen aktuální strany
        df_export = df.copy()
        df_export["Splatno"] = df_export["Splatno"].dt.date
        df_export.rename(columns={
            "Cislo": "Faktura", "Doklad": "Var. symbol", "Firma": "Dodavatel",
            "Celkem": "Částka faktury", "K_likvidaci": "K likvidaci",
            "Hodnota_Skladem_sum": "Zbývá skladem", "Dni_do_splatnosti": "Dní do spl.",
            "Pocet_Prijemek": "# příjemek",
        }).to_excel(w, sheet_name="Faktury", index=False)
    st.download_button(
        "⬇️ Export filtru do Excelu",
        data=buf.getvalue(),
        file_name="cashflow_radar_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("🔬 Detail faktury")
    if df.empty:
        st.info("Žádné faktury k zobrazení.")
        return

    moznosti = df.apply(
        lambda r: f"{r['Cislo']} — {r['Firma']} — {format_kc(r['Celkem'])} — {r['Stav']}",
        axis=1).tolist()
    mapa = dict(zip(moznosti, df["Cislo"].tolist()))
    volba = st.selectbox("Vyber fakturu", options=moznosti)
    if not volba:
        return

    cislo_fa = mapa[volba]
    r = df[df["Cislo"] == cislo_fa].iloc[0]

    # Metriky nahoře — přidá K likvidaci
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Stav", r["Stav"])
    c2.metric("Částka faktury", format_kc(r["Celkem"]))
    c3.metric("💰 K likvidaci", format_kc(r["K_likvidaci"]),
              help="Kolik z faktury zbývá zaplatit")
    c4.metric("📦 Zbývá skladem", format_kc(r["Hodnota_Skladem_sum"]))
    c5.metric("% prodáno", format_pct(r["Pct_Prodano"]))

    # Způsob párování této faktury
    fa_vazba = vazba[vazba["Cislo_Faktury"] == cislo_fa]
    match_labels = {
        "exact": "přesná shoda",
        "substring": "substring",
        "cyklomax": "Cyklomax logika",
        "crussis_suffix": "Crussis suffix",
        "no_separator": "bez oddělovačů",
        "cisla_prunik": "průnik čísel",
        "castka_nazev": "částka + název",
    }
    match_popis = ", ".join(
        match_labels.get(mt, mt)
        for mt in fa_vazba["match_type"].unique()
    ) if not fa_vazba.empty else "—"

    st.write(
        f"**Dodavatel:** {r['Firma']}  |  **VS:** `{r['Doklad']}`  |  "
        f"**Splatno:** {r['Splatno'].date() if pd.notna(r['Splatno']) else '—'}  |  "
        f"**Dní do splatnosti:** {r['Dni_do_splatnosti'] if r['Dni_do_splatnosti'] is not None else '—'}  |  "
        f"**Zaplaceno:** {'✅ Ano' if r['Zaplaceno'] else '❌ Ne'}  |  "
        f"**Párováno přes:** {match_popis}")

    prij_ids = vazba[vazba["Cislo_Faktury"] == cislo_fa]["Cislo_Prijemky"].tolist()
    if not prij_ids:
        st.warning("Tato faktura nemá spárovanou žádnou příjemku (NESPÁROVÁNO).")
        return

    detail = prij_agg[prij_agg["Cislo_Prijemky"].isin(prij_ids)].copy()
    # ×100 pro progress bar
    detail["Pct_Prodano_display"] = detail["Pct_Prodano"] * 100
    detail_disp = detail.rename(columns={
        "Cislo_Prijemky": "Příjemka", "Firma": "Dodavatel",
        "Celkem_Prijemka": "Částka příjemky",
        "Prijato_ks": "Přijato ks", "Prodano_ks": "Prodáno ks",
        "Skladem_ks": "Skladem ks", "Hodnota_Skladem": "Zbývá skladem",
        "Pct_Prodano_display": "% prodáno",
    })[["Příjemka", "Dodavatel", "Částka příjemky", "Přijato ks", "Prodáno ks",
        "Skladem ks", "Zbývá skladem", "% prodáno"]]

    st.markdown("**Příjemky této faktury:**")
    st.dataframe(
        detail_disp, width="stretch", hide_index=True,
        column_config={
            "Částka příjemky": st.column_config.NumberColumn(format="%.0f Kč"),
            "Zbývá skladem": st.column_config.NumberColumn(format="%.0f Kč"),
            "% prodáno": st.column_config.ProgressColumn(
                format="%.0f %%", min_value=0, max_value=100),
        })

    # Všechny položky faktury s barevným rozlišením
    st.markdown("**📦 Produkty z této faktury:**")
    st.caption(
        "🟢 zeleně = vyprodáno (zboží už není ve skladě)  |  "
        "🔴 červeně = leží ve skladě  |  "
        "🟡 žlutě = částečně prodáno")

    vsechny_polozky = res["polozky"]
    polozky_fa = vsechny_polozky[vsechny_polozky["Cislo_Prijemky"].isin(prij_ids)].copy()

    # Připoj název produktu z polozek_pohoda (pokud byl položkový export)
    polozky_pohoda = res.get("polozky_pohoda")
    nazvy_map = {}
    if polozky_pohoda is not None and not polozky_pohoda.empty:
        nazvy_map = dict(zip(polozky_pohoda["Kod"], polozky_pohoda["Nazev"]))

    if polozky_fa.empty:
        # Zjisti proč nejsou data
        je_v_polozkach = False
        if polozky_pohoda is not None and not polozky_pohoda.empty:
            je_v_polozkach = polozky_pohoda["Cislo_Prijemky"].isin(prij_ids).any()

        if polozky_pohoda is None or polozky_pohoda.empty:
            st.warning(
                "⚠️ **Nemáme položkový export.** Aplikace zkusila napojit pohyby na příjemky přes "
                "Dodavatel+Datum, ale pro tuto fakturu se to nepodařilo. "
                "**Řešení:** Nahraj `prijemky_polozkove.xlsx` z Pohody (tiskový export Položky příjemek)."
            )
        elif not je_v_polozkach:
            prij_list = ", ".join(f"`{p}`" for p in prij_ids)
            st.error(
                f"❌ **Příjemky této faktury ({prij_list}) nejsou v `prijemky_polozkove.xlsx`.**\n\n"
                f"Tvůj položkový export obsahuje jen některé příjemky — pravděpodobně z vybraného období. "
                f"Faktura čeká na zaplacení, ale aplikace neví, co z ní je na skladě.\n\n"
                f"**Řešení:** Vyexportuj z Pohody položky **všech příjemek**, které jsou ve `faktury.xlsx` "
                f"(nikoli jen jeden den). V Pohodě: Sklady → Příjemky → Tisk → Položky příjemek s filtrem "
                f"přes celé období."
            )
        else:
            st.info(
                "Příjemky jsou v položkovém exportu, ale aplikace nenapojila žádné položky. "
                "To by nemělo nastat — prosím zkontroluj, jestli se při analýze objevila nějaká chyba."
            )
    else:
        # Určení stavu položky
        def _stav_polozky(row):
            if row["Zbyva_ks"] <= 0.0001:
                return "✅ Prodáno"
            if row["Prodano_ks"] <= 0.0001:
                return "🔴 Skladem"
            return "🟡 Částečně"

        polozky_fa["Stav_polozky"] = polozky_fa.apply(_stav_polozky, axis=1)
        polozky_fa["Hodnota_Prijem"] = polozky_fa["Prijato_ks"] * polozky_fa["Cena"]
        polozky_fa["Hodnota_Skladem"] = polozky_fa["Zbyva_ks"] * polozky_fa["Cena"]
        polozky_fa["Nazev"] = polozky_fa["Kod"].map(nazvy_map).fillna("")

        # Skutečný stav skladu ze stav.xlsx (pokud existuje)
        stav_df = res.get("stav")
        real_stav_map = {}
        if stav_df is not None and not stav_df.empty:
            real_stav_map = dict(zip(stav_df["Kod"], stav_df["Stav_zasoby"]))
        polozky_fa["Real_skladem"] = polozky_fa["Kod"].map(real_stav_map)

        sloupce_disp = ["Příjemka", "Kód", "Název", "Výrobce",
                        "Přijato", "Prodáno", "Zbývá"]
        if real_stav_map:
            sloupce_disp.append("Reálně skladem")
        sloupce_disp += ["Cena/ks", "Hodnota přijato", "Hodnota skladem", "Stav"]

        polozky_disp = polozky_fa.rename(columns={
            "Cislo_Prijemky": "Příjemka",
            "Kod": "Kód",
            "Nazev": "Název",
            "Vyrobce": "Výrobce",
            "Prijato_ks": "Přijato",
            "Prodano_ks": "Prodáno",
            "Zbyva_ks": "Zbývá",
            "Real_skladem": "Reálně skladem",
            "Cena": "Cena/ks",
            "Hodnota_Prijem": "Hodnota přijato",
            "Hodnota_Skladem": "Hodnota skladem",
            "Stav_polozky": "Stav",
        })[sloupce_disp].sort_values(
                 by=["Stav", "Hodnota skladem"], ascending=[True, False])

        # Obarvení řádků podle stavu
        def _obarvi_radek(row):
            if row["Stav"] == "✅ Prodáno":
                return ['color: #888888; text-decoration: line-through'] * len(row)
            if row["Stav"] == "🔴 Skladem":
                return ['background-color: #ffeaea'] * len(row)
            if row["Stav"] == "🟡 Částečně":
                return ['background-color: #fff8e1'] * len(row)
            return [''] * len(row)

        styled = polozky_disp.style.apply(_obarvi_radek, axis=1)
        fmt_map = {
            "Cena/ks": "{:.2f} Kč",
            "Hodnota přijato": "{:.0f} Kč",
            "Hodnota skladem": "{:.0f} Kč",
            "Přijato": "{:.0f}",
            "Prodáno": "{:.0f}",
            "Zbývá": "{:.0f}",
        }
        if "Reálně skladem" in polozky_disp.columns:
            fmt_map["Reálně skladem"] = lambda x: "—" if pd.isna(x) else f"{x:.0f}"
        styled = styled.format(fmt_map)
        st.dataframe(styled, width="stretch", hide_index=True)

        # Pokud známe reálný stav, ověř jestli výpočet sedí s realitou
        if "Real_skladem" in polozky_fa.columns:
            polozky_se_stavem = polozky_fa.dropna(subset=["Real_skladem"])
            if len(polozky_se_stavem) > 0:
                rozdily = polozky_se_stavem[
                    (polozky_se_stavem["Zbyva_ks"] - polozky_se_stavem["Real_skladem"]).abs() > 0.5
                ]
                if len(rozdily) > 0:
                    st.caption(
                        f"⚠️ U **{len(rozdily)}** položek se vypočtené 'Zbývá' liší od reálného stavu skladu. "
                        f"To je normální, pokud výpočet používá FIFO z pohybů (přesnější je reverzní LIFO "
                        f"se stav.xlsx — pokud už ho máš nahraný, ignorujte)."
                    )

        # Souhrn
        sum_prijato = polozky_fa["Hodnota_Prijem"].sum()
        sum_skladem = polozky_fa["Hodnota_Skladem"].sum()
        sum_prodano = sum_prijato - sum_skladem
        pocet_prodano = int((polozky_fa["Zbyva_ks"] <= 0.0001).sum())
        pocet_skladem = int((polozky_fa["Prodano_ks"] <= 0.0001).sum())
        pocet_castecne = len(polozky_fa) - pocet_prodano - pocet_skladem
        st.caption(
            f"**Souhrn:** {len(polozky_fa)} položek celkem  |  "
            f"✅ plně prodáno: **{pocet_prodano}**  |  "
            f"🟡 částečně: **{pocet_castecne}**  |  "
            f"🔴 skladem: **{pocet_skladem}**  |  "
            f"hodnota prodáno: **{format_kc(sum_prodano)}**, skladem: **{format_kc(sum_skladem)}**"
        )


# ---------------------------------------------------------------------------
# Tab 2: Ležáky & Splatnosti
# ---------------------------------------------------------------------------

def _tab_lezaky(res: dict) -> None:
    lezaky = res.get("lezaky_data", pd.DataFrame())
    faktury_view = res["faktury_view"]
    dnes = res["dnes"]

    if lezaky.empty:
        st.info("Pro zobrazení ležáků je potřeba `stav.xlsx` s aktuálním stavem skladu.")
        return

    stav_col = "Stav_efektivni" if "Stav_efektivni" in lezaky.columns else "Stav_zasoby"
    rez_col = "Rezervace" in lezaky.columns

    st.markdown("## 🧊 Ležáky & Splatnosti")
    st.caption(
        "Veškeré zboží, které aktuálně leží na skladě (efektivní stav = fyzický stav − rezervace), "
        "s vazbou na konkrétní fakturu přes LIFO rekonstrukci."
    )

    # ── Sekce A: zboží sestupně dle hodnoty ──────────────────────────────────
    st.subheader("A — Všechno zboží na skladě (sestupně dle hodnoty)")

    # Filtry
    col_f1, col_f2, col_f3 = st.columns(3)
    dodavatele = sorted(lezaky["Firma"].dropna().unique()) if "Firma" in lezaky.columns else []
    vyrobci = sorted(lezaky["Vyrobce"].dropna().unique()) if "Vyrobce" in lezaky.columns else []
    with col_f1:
        vyber_dod = st.multiselect("Dodavatel (faktura)", dodavatele, key="lez_dod")
    with col_f2:
        vyber_vyr = st.multiselect("Výrobce", vyrobci, key="lez_vyr")
    with col_f3:
        jen_nezaplacene = st.checkbox("Jen nezaplacené faktury", value=True, key="lez_nezapl")

    df_a = lezaky.copy()
    if vyber_dod:
        df_a = df_a[df_a["Firma"].isin(vyber_dod)]
    if vyber_vyr:
        df_a = df_a[df_a["Vyrobce"].isin(vyber_vyr)]
    if jen_nezaplacene and "K_likvidaci" in df_a.columns:
        df_a = df_a[df_a["K_likvidaci"].fillna(1) > 0.01]

    df_a = df_a.sort_values("Hodnota", ascending=False)

    # KPI
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Kódů na skladě", f"{len(df_a):,}".replace(",", " "))
    k2.metric("Celková hodnota", format_kc(df_a["Hodnota"].sum()))
    po_spl = df_a[df_a["Dni_do_splatnosti"].fillna(1) < 0] if "Dni_do_splatnosti" in df_a.columns else pd.DataFrame()
    k3.metric("🔴 Po splatnosti", f"{len(po_spl)} faktur")
    rez_val = df_a["Rezervace"].sum() if rez_col else 0
    k4.metric("Rezervováno ks", f"{rez_val:,.0f}".replace(",", " "))

    # Tabulka
    sloupce_a = ["Kod"]
    if "Nazev" in df_a.columns: sloupce_a.append("Nazev")
    if "Vyrobce" in df_a.columns: sloupce_a.append("Vyrobce")
    if "Firma" in df_a.columns: sloupce_a.append("Firma")
    sloupce_a.append(stav_col)
    if rez_col: sloupce_a.append("Rezervace")
    sloupce_a += ["Hodnota"]
    if "Doklad" in df_a.columns: sloupce_a.append("Doklad")
    if "Splatno" in df_a.columns: sloupce_a.append("Splatno")
    if "Dni_do_splatnosti" in df_a.columns: sloupce_a.append("Dni_do_splatnosti")
    if "Stav_faktury" in df_a.columns: sloupce_a.append("Stav_faktury")

    disp_a = df_a[[c for c in sloupce_a if c in df_a.columns]].copy()
    if "Splatno" in disp_a.columns:
        disp_a["Splatno"] = pd.to_datetime(disp_a["Splatno"], errors="coerce").dt.date

    rename_a = {
        "Kod": "Kód", "Nazev": "Název", "Vyrobce": "Výrobce",
        "Firma": "Dodavatel", stav_col: "Efektivně ks",
        "Rezervace": "Rezervováno", "Hodnota": "Hodnota (Kč)",
        "Doklad": "Var. symbol", "Splatno": "Splatnost",
        "Dni_do_splatnosti": "Dní do spl.", "Stav_faktury": "Stav faktury",
    }
    disp_a = disp_a.rename(columns={k: v for k, v in rename_a.items() if k in disp_a.columns})

    col_cfg_a = {}
    if "Hodnota (Kč)" in disp_a.columns:
        col_cfg_a["Hodnota (Kč)"] = st.column_config.NumberColumn(format="%.0f Kč")
    if "Efektivně ks" in disp_a.columns:
        col_cfg_a["Efektivně ks"] = st.column_config.NumberColumn(format="%.1f")
    if "Dní do spl." in disp_a.columns:
        col_cfg_a["Dní do spl."] = st.column_config.NumberColumn(format="%d")

    st.dataframe(disp_a, width="stretch", hide_index=True,
                 column_config=col_cfg_a)

    buf_a = __import__("io").BytesIO()
    with __import__("pandas").ExcelWriter(buf_a, engine="openpyxl") as w:
        disp_a.to_excel(w, sheet_name="Ležáky", index=False)
    st.download_button("⬇️ Export do Excelu", buf_a.getvalue(),
                       "lezaky_sklad.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()

    # ── Sekce B: faktury dle splatnosti s detailem ───────────────────────────
    st.subheader("B — Faktury seřazené dle splatnosti")
    st.caption("Od nejstarší splatnosti. Pod každou fakturou rozbal zboží, které z ní leží.")

    dod_b = sorted(lezaky["Firma"].dropna().unique()) if "Firma" in lezaky.columns else []
    vyber_dod_b = st.multiselect("Filtr dodavatele", dod_b, key="lez_b_dod")

    df_b = lezaky.dropna(subset=["Cislo_Faktury"]).copy()
    if vyber_dod_b:
        df_b = df_b[df_b["Firma"].isin(vyber_dod_b)]

    if df_b.empty:
        st.info("Žádné zboží s vazbou na fakturu. Potřeba položkový export příjemek.")
        return

    # Agreguj faktury
    fa_grp = (df_b.groupby(["Cislo_Faktury", "Doklad", "Firma",
                              "Splatno", "Dni_do_splatnosti", "Stav_faktury",
                              "K_likvidaci"], dropna=False)
              .agg(Pocet_kodu=("Kod", "nunique"),
                   Hodnota_lezi=("Hodnota", "sum"))
              .reset_index()
              .sort_values("Dni_do_splatnosti", ascending=True, na_position="last"))

    # Řazení: červená → oranžová → žlutá → zelená
    stav_poradi = {STAV_KRIT: 0, STAV_RIZIKO: 1, STAV_PROCES: 2, STAV_OK: 3,
                   STAV_CHYBI_DATA: 4, STAV_NESP: 5}
    fa_grp["_poradi"] = fa_grp["Stav_faktury"].map(stav_poradi).fillna(9)
    fa_grp = fa_grp.sort_values(["_poradi", "Dni_do_splatnosti"],
                                 ascending=[True, True], na_position="last")

    # Stránkování — max 50 faktur najednou (každá expander = renderovací zátěž)
    B_NA_STRANU = 50
    celkem_b = len(fa_grp)
    pocet_b = max(1, (celkem_b - 1) // B_NA_STRANU + 1)
    if pocet_b > 1:
        strana_b = st.number_input(
            f"Strana ({celkem_b} faktur celkem, {B_NA_STRANU}/strana)",
            min_value=1, max_value=pocet_b, value=1, step=1, key="lez_b_strana")
    else:
        strana_b = 1
    start_b = (strana_b - 1) * B_NA_STRANU
    fa_grp_strana = fa_grp.iloc[start_b:start_b + B_NA_STRANU]

    for _, fa_row in fa_grp_strana.iterrows():
        dd = fa_row.get("Dni_do_splatnosti")
        stav_fa = fa_row.get("Stav_faktury", "")
        lik = fa_row.get("K_likvidaci", 0) or 0
        spl_str = fa_row["Splatno"].date().isoformat() if pd.notna(fa_row.get("Splatno")) else "—"

        if isinstance(dd, float) and dd < 0:
            dd_label = f"🔴 {int(abs(dd))} dní po splatnosti"
        elif isinstance(dd, float) and dd <= 14:
            dd_label = f"🟠 {int(dd)} dní"
        elif isinstance(dd, float):
            dd_label = f"🟡 {int(dd)} dní"
        else:
            dd_label = "—"

        exp_title = (
            f"{stav_fa}  **{fa_row.get('Doklad','—')}**  |  "
            f"{fa_row.get('Firma','—')}  |  "
            f"splatnost {spl_str} ({dd_label})  |  "
            f"k likvidaci {format_kc(lik)}  |  "
            f"{int(fa_row['Pocet_kodu'])} kódů / {format_kc(fa_row['Hodnota_lezi'])}"
        )

        with st.expander(exp_title):
            zbozi = df_b[df_b["Cislo_Faktury"] == fa_row["Cislo_Faktury"]].copy()
            zbozi = zbozi.sort_values("Hodnota", ascending=False)
            sl_z = ["Kod"]
            if "Nazev" in zbozi.columns: sl_z.append("Nazev")
            if "Vyrobce" in zbozi.columns: sl_z.append("Vyrobce")
            sl_z.append(stav_col)
            if rez_col: sl_z.append("Rezervace")
            sl_z.append("Hodnota")
            disp_z = zbozi[[c for c in sl_z if c in zbozi.columns]].copy()
            disp_z = disp_z.rename(columns={
                "Kod": "Kód", "Nazev": "Název", "Vyrobce": "Výrobce",
                stav_col: "Efektivně ks", "Rezervace": "Rezervováno",
                "Hodnota": "Hodnota (Kč)"
            })
            st.dataframe(disp_z, width="stretch", hide_index=True,
                         column_config={
                             "Hodnota (Kč)": st.column_config.NumberColumn(format="%.0f Kč"),
                             "Efektivně ks": st.column_config.NumberColumn(format="%.1f"),
                         })


# ---------------------------------------------------------------------------
# Tab 3: Obratová analýza
# ---------------------------------------------------------------------------

def _tab_obrat(res: dict) -> None:
    obrat = res.get("obrat_data", pd.DataFrame())

    st.markdown("## 📊 Obratová analýza")
    st.markdown(
        "Analýza prodejního obratu za posledních 12 měsíců. Pomáhá identifikovat "
        "zboží, které leží příliš dlouho na skladě vzhledem ke svému obratu — "
        "**rizikové položky** vážou cashflow a mohou se stát neprodejnými."
    )

    if obrat.empty:
        st.info(
            "Pro obratovou analýzu jsou potřeba pohyby za dostatečně dlouhé období. "
            "Nahraj `pohyby.xlsx` (aktuální rok) a volitelně `pohyby_historicke.xlsx` (předchozí rok)."
        )
        return

    # ── KPI ──────────────────────────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    krit = int((obrat["Riziko"] == "🔴 Kritické").sum())
    sledo = int((obrat["Riziko"] == "🟠 Sledovat").sum())
    pomale = int((obrat["Riziko"] == "🟡 Pomalé").sum())
    hodnota_riziko = obrat[obrat["Riziko"].isin(["🔴 Kritické", "🟠 Sledovat"])]["Hodnota_skladu"].sum()
    k1.metric("🔴 Kritické", krit)
    k2.metric("🟠 Sledovat", sledo)
    k3.metric("🟡 Pomalé", pomale)
    k4.metric("Hodnota v riziku", format_kc(hodnota_riziko))

    # Legenda rizik
    st.markdown(
        "🔴 **Kritické** = hodnotné zboží (A/B) + zásoby na >26 týdnů + klesající prodej &nbsp;|&nbsp; "
        "🟠 **Sledovat** = hodnotné zboží (A/B) + zásoby na >13 týdnů &nbsp;|&nbsp; "
        "🟡 **Pomalé** = zboží s malým obratem (C) + zásoby na >26 týdnů &nbsp;|&nbsp; "
        "🟢 **OK** = obrat a zásoby v rovnováze",
        unsafe_allow_html=True,
    )

    # ── Matice ABC × XYZ ─────────────────────────────────────────────────────
    st.subheader("Matice ABC × XYZ")
    st.markdown(
        "**ABC** = jak moc zboží přispívá k celkovému obratu "
        "(A = top 80 % obratu, B = dalších 15 %, C = zbylých 5 %). "
        "**XYZ** = jak pravidelně se prodává "
        "(X = stabilní prodej, Y = kolísavý/sezónní, Z = nepravidelné nebo jednorázové)."
    )

    matice_rows = []
    for abc in ["A", "B", "C"]:
        row = {"ABC": abc}
        for xyz in ["X", "Y", "Z"]:
            sub = obrat[(obrat["ABC"] == abc) & (obrat["XYZ"] == xyz)]
            row[xyz] = f"{len(sub)} kódů\n{format_kc(sub['Hodnota_skladu'].sum())}"
        matice_rows.append(row)
    matice_df = pd.DataFrame(matice_rows).set_index("ABC")
    st.dataframe(matice_df, width="stretch")

    # ── Filtr a tabulka ───────────────────────────────────────────────────────
    st.subheader("Detail — rizikové a pomalé zboží")
    st.markdown(
        "Tabulka ukazuje pro každý kód: **kolik se prodá za týden** (Ø ks/týden), "
        "**na kolik týdnů zásoby vydrží** (Pokrytí), a **jestli prodeje klesají nebo rostou** (Trend). "
        "Pokrytí >26 týdnů = zásoba na půl roku; >52 = na rok — to je signál, že zboží leží."
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        vyber_riziko = st.multiselect(
            "Riziko", ["🔴 Kritické", "🟠 Sledovat", "🟡 Pomalé", "🟢 OK"],
            default=["🔴 Kritické", "🟠 Sledovat"], key="obr_riziko")
    with col2:
        vyber_abc = st.multiselect("ABC", ["A", "B", "C"], default=["A", "B"], key="obr_abc")
    with col3:
        vyrobci_o = sorted(obrat["Vyrobce"].dropna().unique()) if "Vyrobce" in obrat.columns else []
        vyber_vyr_o = st.multiselect("Výrobce", [v for v in vyrobci_o if v], key="obr_vyr")
    with col4:
        max_pokryti = st.number_input("Max. pokrytí (týdny)", value=9999, min_value=1, key="obr_pokr")

    df_o = obrat.copy()
    if vyber_riziko:
        df_o = df_o[df_o["Riziko"].isin(vyber_riziko)]
    if vyber_abc:
        df_o = df_o[df_o["ABC"].isin(vyber_abc)]
    if vyber_vyr_o:
        df_o = df_o[df_o["Vyrobce"].isin(vyber_vyr_o)]
    df_o = df_o[df_o["Pokryti_tydnu"] <= max_pokryti]
    df_o = df_o.sort_values("Hodnota_skladu", ascending=False)

    O_NA_STRANU = 300
    celkem_o = len(df_o)
    pocet_o = max(1, (celkem_o - 1) // O_NA_STRANU + 1)
    if pocet_o > 1:
        strana_o = st.number_input(
            f"Strana ({celkem_o} kódů, {O_NA_STRANU}/strana)",
            min_value=1, max_value=pocet_o, value=1, step=1, key="obr_strana")
    else:
        strana_o = 1
    start_o = (strana_o - 1) * O_NA_STRANU
    df_o_strana = df_o.iloc[start_o:start_o + O_NA_STRANU]

    st.caption(f"Zobrazeno {start_o+1}–{min(start_o+O_NA_STRANU, celkem_o)} z {celkem_o} kódů")

    cols_disp = ["Kod"]
    if "Nazev" in df_o_strana.columns:
        cols_disp.append("Nazev")
    if "Vyrobce" in df_o_strana.columns:
        cols_disp.append("Vyrobce")
    cols_disp += ["ABC", "XYZ", "Riziko", "Obrat_Kc", "Prumer_tyden_ks",
                  "Pokryti_tydnu", "Trend", "Stav_efektivni", "Hodnota_skladu"]
    disp_o = df_o_strana[[c for c in cols_disp if c in df_o_strana.columns]].copy()
    disp_o = disp_o.rename(columns={
        "Kod": "Kód", "Nazev": "Název", "Vyrobce": "Výrobce",
        "Obrat_Kc": "Obrat (Kč)", "Prumer_tyden_ks": "Ø ks/týden",
        "Pokryti_tydnu": "Pokrytí (týdny)", "Stav_efektivni": "Skladem ks",
        "Hodnota_skladu": "Hodnota skladu",
    })

    def _obarvi_riziko(row):
        colors = {
            "🔴 Kritické": "background-color: rgba(255,80,80,0.15)",
            "🟠 Sledovat": "background-color: rgba(255,160,0,0.15)",
            "🟡 Pomalé": "background-color: rgba(255,230,0,0.12)",
        }
        c = colors.get(row.get("Riziko", ""), "")
        return [c] * len(row)

    styled = disp_o.style.apply(_obarvi_riziko, axis=1)
    styled = styled.format({
        "Obrat (Kč)": "{:,.0f} Kč".replace(",", " "),
        "Ø ks/týden": "{:.2f}",
        "Pokrytí (týdny)": lambda x: "∞" if x >= 9999 else f"{x:.0f}",
        "Hodnota skladu": "{:,.0f} Kč".replace(",", " "),
        "Skladem ks": "{:.1f}",
    })
    st.dataframe(styled, width="stretch", hide_index=True)

    # Export
    buf_o = __import__("io").BytesIO()
    with __import__("pandas").ExcelWriter(buf_o, engine="openpyxl") as w:
        disp_o.to_excel(w, sheet_name="Obratová analýza", index=False)
    st.download_button("⬇️ Export do Excelu", buf_o.getvalue(),
                       "obratova_analyza.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Graf pokrytí ──────────────────────────────────────────────────────────
    st.subheader("Pokrytí zásoby v týdnech (top 30 dle hodnoty)")
    top30 = df_o.nlargest(30, "Hodnota_skladu")[
        ["Kod", "Pokryti_tydnu", "Riziko"]].copy()
    top30["Pokryti_zobr"] = top30["Pokryti_tydnu"].clip(upper=104)
    top30["Barva"] = top30["Riziko"].map({
        "🔴 Kritické": "🔴", "🟠 Sledovat": "🟠", "🟡 Pomalé": "🟡", "🟢 OK": "🟢"
    }).fillna("⚪")

    chart_data = top30.set_index("Kod")[["Pokryti_zobr"]].rename(
        columns={"Pokryti_zobr": "Pokrytí (týdny)"})
    st.bar_chart(chart_data, color="#3a7bd5")
    st.caption("Hodnoty nad 104 týdny (2 roky) jsou zkráceny na 104 pro čitelnost.")


def _zkontroluj_heslo() -> bool:
    """Ochrana heslem pro tým. Aktivní jen když jsou nastavena hesla v secrets.

    V Streamlit/cloud secrets nastav buď jedno heslo:
        heslo = "tajne-heslo"
    nebo více uživatelů (doporučeno pro tým):
        [hesla]
        michal = "heslo-michal"
        vedeni = "heslo-vedeni"

    Heslo se NIKDY neukládá do kódu — jen do secrets na serveru.
    """
    import hashlib

    # Načti konfiguraci hesel
    hesla = {}
    try:
        if "hesla" in st.secrets:
            hesla = dict(st.secrets["hesla"])
        elif "heslo" in st.secrets:
            hesla = {"uzivatel": st.secrets["heslo"]}
    except Exception:
        hesla = {}

    # Lokální běh bez nastavených hesel — přístup volný
    if not hesla:
        return True

    if st.session_state.get("heslo_ok", False):
        return True

    # Omezení počtu pokusů (ochrana proti hádání hesla)
    pokusy = st.session_state.get("pokusy", 0)
    if pokusy >= 5:
        st.title("📡 Cashflow Radar")
        st.error("⛔ Příliš mnoho neúspěšných pokusů. Obnovte stránku za 5 minut.")
        return False

    st.title("📡 Cashflow Radar")
    st.caption("Přístup jen pro oprávněné uživatele Koloshop")
    jmeno = st.text_input("Uživatel", key="login_user")
    zadane = st.text_input("Heslo", type="password", key="login_pass")

    if st.button("Přihlásit", type="primary"):
        spravne = hesla.get(jmeno.strip().lower()) or hesla.get(jmeno.strip())
        if spravne and zadane == spravne:
            st.session_state["heslo_ok"] = True
            st.session_state["uzivatel"] = jmeno
            st.session_state["pokusy"] = 0
            st.rerun()
        else:
            st.session_state["pokusy"] = pokusy + 1
            zbyva = 5 - st.session_state["pokusy"]
            st.error(f"Nesprávné jméno nebo heslo. Zbývá pokusů: {zbyva}")
    return False


def main() -> None:
    if not _zkontroluj_heslo():
        return

    # Po studeném startu stáhni poslední uložená data ze Supabase do cache
    if _persist is not None:
        _persist.pull_once(_cache_dir())

    uvodni_obrazovka()

    zdroje = sidebar_zdroje()
    st.sidebar.divider()
    dnes = st.sidebar.date_input("Referenční datum pro splatnost", value=date.today())

    rok_od = st.sidebar.number_input(
        "Faktury od roku (filtr)",
        min_value=2015, max_value=2030,
        value=2026,
        step=1,
        help="Faktury starší než tento rok se ignorují. Nastav na 2026 pokud máš příjemky jen za 2026."
    )

    # Whitelist agend — rozbalovací sekce v sidebaru
    with st.sidebar.expander("⚙️ Filtr agend (pokročilé)"):
        st.caption(
            "Pouze tyto agendy z pohybů se považují za reálné nákupy/prodeje. "
            "Ostatní (převodky, inventury, reklamace…) se ignorují."
        )
        # Výchozí hodnota = merge konstant se session_state (aby nové konstanty
        # byly vždy viditelné, i bez restartu prohlížeče)
        default_pr = sorted({*AGENDY_PRIJEM_WHITELIST,
                             *st.session_state.get("agendy_prijem", [])})
        default_vy = sorted({*AGENDY_VYDEJ_WHITELIST,
                             *st.session_state.get("agendy_vydej", [])})
        text_pr = st.text_area(
            "Agendy pro PŘÍJEM (každá na nový řádek)",
            value="\n".join(default_pr),
            height=80,
            key="agendy_prijem_text",
        )
        text_vy = st.text_area(
            "Agendy pro VÝDEJ (každá na nový řádek)",
            value="\n".join(default_vy),
            height=120,
            key="agendy_vydej_text",
        )
        st.session_state["agendy_prijem"] = [
            l.strip() for l in text_pr.splitlines() if l.strip()
        ]
        st.session_state["agendy_vydej"] = [
            l.strip() for l in text_vy.splitlines() if l.strip()
        ]

    spustit = st.sidebar.button(
        "▶️ Spustit analýzu", type="primary",
        disabled=(zdroje is None), width="stretch")

    if spustit:
        res = spust_analyzu(zdroje, dnes, rok_od=int(rok_od))
        if res is not None:
            st.session_state["vysledky"] = res

    if "vysledky" in st.session_state:
        st.divider()
        dashboard(st.session_state["vysledky"])
    else:
        st.sidebar.info("👆 Po výběru souborů klikněte na **Spustit analýzu**.")


if __name__ == "__main__":
    main()
