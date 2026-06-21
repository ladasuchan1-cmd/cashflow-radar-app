"""
persistence.py — trvalé uložení nahraných dat pro Cashflow Radar
přes Supabase Storage (zdarma, privátní úložiště).

PROČ
----
Streamlit Cloud nemá trvalý disk: po uspání / redeployi se složka `data_cache/`
smaže. Aby "poslední nahraná data zůstala až do dalšího nahrání", zrcadlíme
`data_cache/` do PRIVÁTNÍHO Supabase bucketu. Díky tomu může být GitHub repo
appky VEŘEJNÉ (jen kód) — žádné faktury v repu → žádný limit privátních appek.

Klíč k Supabase je v st.secrets (NE v repu), takže veřejné repo nic neprozradí.

POUŽITÍ (volá se z cashflow_radar.py)
-------------------------------------
  pull_once(cache_dir)     # 1× na začátku session — stáhne data z bucketu
  push_file(local_path)    # po každém uložení souboru — pošle ho do bucketu

Když není v st.secrets sekce [supabase], všechny funkce jsou no-op (appka jede
dál jen s lokální cache — typicky lokální běh na PC).
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

try:
    import requests
except Exception:  # noqa: BLE001
    requests = None


def _conf():
    """Vrátí (url, key, bucket) nebo None, když Supabase není nakonfigurovaný."""
    if requests is None:
        return None
    try:
        s = st.secrets.get("supabase", {})
    except Exception:  # noqa: BLE001 — st.secrets nemusí existovat lokálně
        return None
    url, key = s.get("url"), s.get("key")
    bucket = s.get("bucket", "radar-data")
    if not (url and key):
        return None
    return url.rstrip("/"), key, bucket


def _headers(key: str) -> dict:
    return {"apikey": key, "Authorization": f"Bearer {key}"}


def _ensure_bucket(url: str, key: str, bucket: str) -> None:
    """Vytvoří privátní bucket, pokud ještě neexistuje (ignoruje 'už existuje')."""
    try:
        requests.post(
            f"{url}/storage/v1/bucket",
            headers={**_headers(key), "Content-Type": "application/json"},
            json={"name": bucket, "id": bucket, "public": False},
            timeout=15,
        )
    except Exception:  # noqa: BLE001
        pass


def push_file(local_path: Path) -> None:
    """Nahraje jeden soubor z data_cache/ do Supabase bucketu (vytvoří/přepíše)."""
    conf = _conf()
    if conf is None:
        return
    url, key, bucket = conf
    local_path = Path(local_path)
    if not local_path.exists():
        return
    _ensure_bucket(url, key, bucket)
    data = local_path.read_bytes()
    try:
        # x-upsert: true → přepíše, pokud už soubor existuje
        r = requests.post(
            f"{url}/storage/v1/object/{bucket}/{local_path.name}",
            headers={**_headers(key),
                     "Content-Type": "application/octet-stream",
                     "x-upsert": "true"},
            data=data,
            timeout=60,
        )
        if r.status_code >= 400:
            st.sidebar.warning(
                f"Uložení {local_path.name} do Supabase selhalo ({r.status_code}).")
    except Exception as e:  # noqa: BLE001
        st.sidebar.warning(f"Supabase nedostupné: {e}")


def pull_once(cache_dir: Path) -> None:
    """1× za session stáhne soubory z bucketu do lokální cache.
    Stahuje jen soubory, které lokálně chybí (po studeném startu)."""
    if st.session_state.get("_sb_pulled"):
        return
    st.session_state["_sb_pulled"] = True

    conf = _conf()
    if conf is None:
        return
    url, key, bucket = conf
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(exist_ok=True)

    # 1) Seznam souborů v bucketu
    try:
        r = requests.post(
            f"{url}/storage/v1/object/list/{bucket}",
            headers={**_headers(key), "Content-Type": "application/json"},
            json={"prefix": "", "limit": 100,
                  "sortBy": {"column": "name", "order": "asc"}},
            timeout=30,
        )
        if r.status_code >= 400:
            return
        soubory = [o["name"] for o in r.json() if o.get("name")]
    except Exception:  # noqa: BLE001
        return

    # 2) Stáhni ty, co lokálně chybí
    for nazev in soubory:
        lokal = cache_dir / Path(nazev).name
        if lokal.exists():
            continue  # lokální (čerstvější) verzi nepřepisujeme
        try:
            rr = requests.get(
                f"{url}/storage/v1/object/{bucket}/{nazev}",
                headers=_headers(key),
                timeout=120,
            )
            if rr.status_code < 400 and rr.content:
                lokal.write_bytes(rr.content)
        except Exception:  # noqa: BLE001
            pass
