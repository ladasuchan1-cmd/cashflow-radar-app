# Cashflow Radar jako appka (stejně jako ks-porovnani)

Appka poběží na veřejné adrese `https://nazev.streamlit.app`, otevřeš ji na
mobilu i PC, **přihlásíš se heslem** a **poslední nahraná data v ní zůstanou
i po uspání** (zrcadlí se do tvého privátního GitHub repa). Vše zdarma, nic
neběží na vašem firemním serveru.

---

## Co je ve složce

| Soubor | K čemu |
|---|---|
| `cashflow_radar.py` | hlavní aplikace (už obsahuje přihlášení) |
| `persistence.py` | zrcadlí nahraná data do GitHub repa (trvalé uložení) |
| `requirements.txt` | knihovny, které Streamlit Cloud doinstaluje |
| `.streamlit/config.toml` | nastavení (barvy Koloshop, limit uploadu) |
| `.streamlit/secrets.toml.PRIKLAD` | šablona hesel a GitHub tokenu |
| `.gitignore` | aby se hesla a lokální data nikdy nenahrála na GitHub |

> Soubory v Pohodě se mohou jmenovat různě — nahráváš je v levém panelu appky
> ručně, na názvech nezáleží. Trvale se ukládají pod stálými názvy do repa.

---

## Krok 1 — Privátní repo na GitHubu

1. github.com → vpravo nahoře **+ → New repository**
2. Název: `cashflow-radar-app`
3. **Private** (důležité — jsou tam citlivé faktury)
4. Create repository
5. Nahraj sem **všechny soubory z této složky** (`secrets.toml.PRIKLAD` klidně taky;
   skutečný `secrets.toml` ani `data_cache/` se díky `.gitignore` nenahrají).

## Krok 2 — GitHub token (pro trvalé ukládání dat)

1. github.com → **Settings → Developer settings → Fine-grained tokens → Generate new token**
2. **Repository access:** jen repo `cashflow-radar-app`
3. **Permissions → Repository → Contents: Read and write**
4. Vygeneruj a **zkopíruj token** (`github_pat_…`) — uvidíš ho jen jednou.

## Krok 3 — Deploy na Streamlit Cloud

1. share.streamlit.io → **Sign in with GitHub**
2. **Create app → Deploy a public app from GitHub**
3. Repository: `cashflow-radar-app`, Branch: `main`, Main file: `cashflow_radar.py`
4. **Advanced settings → Secrets** — vlož (podle `secrets.toml.PRIKLAD`):

   ```toml
   [hesla]
   ladik = "tvoje-silne-heslo"
   kolega = "jeho-heslo"

   [github]
   token = "github_pat_…"          # z kroku 2
   repo  = "tvuj-ucet/cashflow-radar-app"
   branch = "main"
   ```
   (Jména v `[hesla]` piš malými písmeny — přihlášení je nerozlišuje.)
5. **Deploy** → počkej 2–5 min.

## Krok 4 — Použití

- Otevři adresu na mobilu i PC, přihlas se.
- Poprvé nahraj exporty z Pohody (necháš zaškrtnuto „💾 Zapamatovat tato data")
  → appka je uloží a zároveň zrcadlí do repa.
- Příště zvol v panelu **„Použít poslední nahraná data"** — předvyplní se
  poslední nahrání, dokud nenahraješ nová.
- Na telefonu: prohlížeč → Sdílet → **Přidat na plochu** = ikona jako appka.

---

## Poznámky / kompromisy (ať to víš)

- **Soukromí:** appka běží na americkém cloudu (jako ks-porovnani). Heslo +
  privátní repo jsou ochrana, ale faktury fyzicky leží mimo firmu. Pro citlivější
  provoz je bezpečnější vlastní/EU server — to už ale není „zdarma jako YoY".
- **Perzistence:** data se ukládají commitem do privátního repa (`data_cache/`).
  Historie repa tím poroste (každé nahrání = nový commit s xlsx). Pro občasné
  použití 2–3 lidí to nevadí.
- Když nevyplníš `[github]` token, appka funguje taky — ale na Streamlit Cloud
  by se data po uspání ztratila. Lokálně (na PC) `[github]` nepotřebuješ.
- Po uspání se appka při prvním otevření pár sekund „budí" — normální.
