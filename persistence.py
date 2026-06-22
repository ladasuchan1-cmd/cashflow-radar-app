"""
persistence.py — synchronizace cache souborů se Supabase Storage.

Slouží k tomu, aby poslední nahraná data přežila uspání aplikace na
Streamlit Community Cloud (kde se lokální složka data_cache/ po čase smaže).

Hlavní aplikace (cashflow_radar.py) volá:
    persistence.push_file(cesta)     — po uložení souboru do cache
    persistence.pull_once(cache_dir) — jednou po startu, stáhne data zpět

Konfigurace přes Streamlit secrets (.streamlit/secrets.toml nebo cloud Secrets):
    [supabase]
    url = "https://abcxyz.supabase.co"
    service_key = "eyJ...dlouhý service_role klíč..."
    bucket = "data"          # volitelné, výchozí 'data'
    prefix = "cashflow"      # volitelné, podsložka v bucketu

Pokud secrets chybí nebo je něco špatně, všechny funkce jsou no-op
(jen tiše nic neudělají) a aplikace běží dál jen s lokální cache.
"""

from __future__ import annotations

from pathlib import Path

try:
    import requests
except Exception:  # noqa: BLE001
    requests = None

try:
    import streamlit as st
except Exception:  # noqa: BLE001
    st = None


# Soubory, které se synchronizují (názvy v cache složce)
_SYNC_FILES = [
    "faktury.xlsx",
    "prijemky.xlsx",
    "pohyby.xlsx",
    "pohyby_historicke.xlsx",
    "stav.xlsx",
    "prijemky_polozkove.xlsx",
    "prijemky_polozkove.pdf",
    "prijemky_polozkove_nove.xlsx",
    "prijemky_polozkove_nove.pdf",
]


def _config() -> dict | None:
    """Načte konfiguraci Supabase ze streamlit secrets. None = vypnuto."""
    if st is None or requests is None:
        return None
    try:
        if "supabase" not in st.secrets:
            return None
        cfg = st.secrets["supabase"]
        url = str(cfg.get("url", "")).rstrip("/")
        key = str(cfg.get("service_key", ""))
        if not url or not key:
            return None
        return {
            "url": url,
            "key": key,
            "bucket": str(cfg.get("bucket", "data")),
            "prefix": str(cfg.get("prefix", "cashflow")).strip("/"),
        }
    except Exception:  # noqa: BLE001
        return None


def _object_path(cfg: dict, filename: str) -> str:
    """Sestaví cestu objektu v bucketu (s volitelným prefixem)."""
    if cfg["prefix"]:
        return f"{cfg['prefix']}/{filename}"
    return filename


def _headers(cfg: dict, content_type: str | None = None) -> dict:
    h = {
        "Authorization": f"Bearer {cfg['key']}",
        "apikey": cfg["key"],
    }
    if content_type:
        h["Content-Type"] = content_type
    return h


def push_file(cesta) -> bool:
    """
    Nahraje (nebo přepíše) jeden soubor z cache do Supabase Storage.
    Vrací True při úspěchu, jinak False. Při vypnuté konfiguraci tiše vrátí False.
    """
    cfg = _config()
    if cfg is None:
        return False

    cesta = Path(cesta)
    if not cesta.exists():
        return False

    filename = cesta.name
    obj = _object_path(cfg, filename)
    # upsert=true → přepíše existující objekt
    endpoint = f"{cfg['url']}/storage/v1/object/{cfg['bucket']}/{obj}"

    try:
        with open(cesta, "rb") as f:
            data = f.read()
        resp = requests.post(
            endpoint,
            headers={**_headers(cfg, "application/octet-stream"),
                     "x-upsert": "true"},
            data=data,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            return True
        # Tiché varování do UI (neshazuje appku)
        if st is not None:
            st.warning(
                f"Uložení {filename} do Supabase selhalo "
                f"({resp.status_code}).")
        return False
    except Exception as e:  # noqa: BLE001
        if st is not None:
            st.warning(f"Uložení {filename} do Supabase selhalo: {e}")
        return False


def _download_file(cfg: dict, filename: str, cilova_cesta: Path) -> bool:
    """Stáhne jeden objekt z bucketu do cílové cesty. True při úspěchu."""
    obj = _object_path(cfg, filename)
    endpoint = f"{cfg['url']}/storage/v1/object/{cfg['bucket']}/{obj}"
    try:
        resp = requests.get(endpoint, headers=_headers(cfg), timeout=30)
        if resp.status_code == 200 and resp.content:
            with open(cilova_cesta, "wb") as f:
                f.write(resp.content)
            return True
        return False  # 404 = soubor v bucketu není, to je v pořádku
    except Exception:  # noqa: BLE001
        return False


def pull_once(cache_dir) -> int:
    """
    Jednou po startu aplikace stáhne všechny synchronizované soubory
    ze Supabase do lokální cache (pokud tam ještě nejsou).

    Vrací počet stažených souborů. Při vypnuté konfiguraci vrátí 0.
    Používá streamlit session_state, aby se stahování dělalo jen jednou
    za běh session (ne při každém rerunu).
    """
    cfg = _config()
    if cfg is None:
        return 0

    # Stáhni jen jednou za session
    if st is not None:
        if st.session_state.get("_supabase_pulled", False):
            return 0
        st.session_state["_supabase_pulled"] = True

    cache_dir = Path(cache_dir)
    cache_dir.mkdir(exist_ok=True)

    stazeno = 0
    for filename in _SYNC_FILES:
        cil = cache_dir / filename
        # Pokud už lokálně existuje (čerstvě nahrané), nepřepisuj ze Supabase
        if cil.exists():
            continue
        if _download_file(cfg, filename, cil):
            stazeno += 1

    return stazeno


def status() -> str:
    """Vrátí čitelný stav konfigurace (pro diagnostiku v UI)."""
    cfg = _config()
    if cfg is None:
        if requests is None:
            return "⚠️ Chybí knihovna requests"
        return "ℹ️ Supabase není nakonfigurováno (běží jen lokální cache)"
    return f"✅ Supabase připojeno (bucket '{cfg['bucket']}', prefix '{cfg['prefix']}')"
