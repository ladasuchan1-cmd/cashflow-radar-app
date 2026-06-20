"""
persistence.py — trvalé uložení nahraných dat pro Cashflow Radar
na Streamlit Community Cloud (zdarma, bez vlastního serveru).

PROČ
----
Streamlit Cloud nemá trvalý disk: po uspání / redeployi se složka `data_cache/`
smaže. Aby "poslední nahraná data zůstala až do dalšího nahrání", zrcadlíme
`data_cache/` do PRIVÁTNÍHO GitHub repa (do stejné cesty `data_cache/`), který
appku pohání. Při studeném startu appka data z repa stáhne zpět.

POUŽITÍ (volá se z cashflow_radar.py)
-------------------------------------
  gh_pull_once(cache_dir)        # 1× na začátku session — stáhne data z repa
  gh_push_file(local_path)       # po každém uložení souboru — pošle ho do repa

Když není v st.secrets sekce [github], všechny funkce jsou no-op (appka jede
dál jen s lokální cache — typicky lokální běh).
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st


def _gh():
    """Vrátí (repo, branch) nebo (None, None), když není GitHub nakonfigurovaný."""
    try:
        g = st.secrets.get("github", {})
    except Exception:  # noqa: BLE001 — st.secrets nemusí existovat lokálně
        return None, None
    token, repo_name = g.get("token"), g.get("repo")
    branch = g.get("branch", "main")
    if not (token and repo_name):
        return None, None
    try:
        from github import Github
        return Github(token).get_repo(repo_name), branch
    except Exception as e:  # noqa: BLE001
        st.sidebar.warning(f"GitHub úložiště nedostupné: {e}")
        return None, None


def gh_push_file(local_path: Path) -> None:
    """Pošle jeden soubor z data_cache/ do repa (vytvoří nebo přepíše)."""
    local_path = Path(local_path)
    if not local_path.exists():
        return
    repo, branch = _gh()
    if repo is None:
        return
    cesta = f"data_cache/{local_path.name}"
    data = local_path.read_bytes()
    try:
        existujici = repo.get_contents(cesta, ref=branch)
        repo.update_file(cesta, f"data: {local_path.name}", data,
                         existujici.sha, branch=branch)
    except Exception:  # soubor v repu zatím není
        try:
            repo.create_file(cesta, f"data: {local_path.name}", data, branch=branch)
        except Exception as e:  # noqa: BLE001
            st.sidebar.warning(f"Nepodařilo se uložit {local_path.name} do repa: {e}")


def gh_pull_once(cache_dir: Path) -> None:
    """1× za session stáhne soubory z data_cache/ v repu do lokální cache.
    Stahuje jen soubory, které lokálně chybí (po studeném startu)."""
    if st.session_state.get("_gh_pulled"):
        return
    st.session_state["_gh_pulled"] = True

    repo, branch = _gh()
    if repo is None:
        return
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(exist_ok=True)
    try:
        obsah = repo.get_contents("data_cache", ref=branch)
    except Exception:
        return  # v repu zatím žádná data
    if not isinstance(obsah, list):
        obsah = [obsah]
    for f in obsah:
        if f.type != "file":
            continue
        lokal = cache_dir / Path(f.name).name
        if lokal.exists():
            continue  # lokální (čerstvější) verzi nepřepisujeme
        try:
            lokal.write_bytes(f.decoded_content)
        except Exception:  # noqa: BLE001 — velké soubory přes Git Data API
            try:
                blob = repo.get_git_blob(f.sha)
                import base64
                lokal.write_bytes(base64.b64decode(blob.content))
            except Exception:  # noqa: BLE001
                pass
