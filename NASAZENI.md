# Cashflow Radar jako appka — Supabase verze

Appka poběží na veřejné adrese `https://nazev.streamlit.app`, otevřeš ji na
mobilu i PC, **přihlásíš se heslem** a **poslední nahraná data zůstanou i po
uspání** (ukládají se do privátního **Supabase** úložiště). Vše zdarma.

**Proč Supabase místo GitHubu:** Streamlit Cloud dává zdarma jen 1 appku
z privátního repa. Když data dáme do Supabase, může být **repo veřejné (jen kód,
žádné faktury)** → můžeš mít neomezeně appek a limit tě netrápí. Klíč k Supabase
je v Secrets (ne v repu), appku chrání heslo.

---

## Co je ve složce

| Soubor | K čemu |
|---|---|
| `cashflow_radar.py` | hlavní aplikace (obsahuje přihlášení) |
| `persistence.py` | ukládá nahraná data do Supabase Storage |
| `requirements.txt` | knihovny pro Streamlit Cloud |
| `.streamlit/config.toml` | nastavení (barvy, limit uploadu) |
| `.streamlit/secrets.toml.PRIKLAD` | šablona hesel a Supabase klíče |
| `.gitignore` | aby se hesla a data nikdy nedostala do repa |

---

## Krok 1 — Supabase úložiště

Doporučení: **nový samostatný projekt** jen pro radar (kvůli izolaci od školící
appky). Je to zdarma a 2 minuty.

1. Jdi na **supabase.com** → přihlas se → **New project**.
2. Název: `cashflow-radar`, zvol region (klidně Frankfurt/EU), heslo k DB si ulož.
3. Počkej ~1 min, než se projekt vytvoří.
4. Vlevo **Project Settings (ozubené kolo) → Data API** → zkopíruj **Project URL**
   (`https://….supabase.co`).
5. Tamtéž **API Keys** → zkopíruj **`service_role`** klíč (ten dlouhý, tajný —
   je označený jako secret). **Nepoužívej `anon` klíč.**

> Bucket nemusíš zakládat ručně — appka si privátní bucket `radar-data`
> vytvoří sama při prvním nahrání.

---

## Krok 2 — Veřejné repo na GitHubu

1. github.com → **+ → New repository**
2. Název: `cashflow-radar-app`
3. Nech **Public** (může být veřejné — žádná data tu nejsou, jen kód).
4. Create repository → **uploading an existing file**.
5. Přetáhni sem **všechny soubory z této složky** (i složku `.streamlit`).
   `.gitignore` zajistí, že se hesla ani data nikdy nenahrají.
6. **Commit changes**.

---

## Krok 3 — Deploy na Streamlit Cloud

1. share.streamlit.io → **Sign in with GitHub**
2. **Create app → Deploy a public app from GitHub**
3. Repository: `tvuj-ucet/cashflow-radar-app`, Branch: `main`,
   Main file: `cashflow_radar.py`
4. **Advanced settings → Secrets** — vlož (podle `secrets.toml.PRIKLAD`):

   ```toml
   [hesla]
   ladik = "tvoje-silne-heslo"
   kolega = "jeho-heslo"

   [supabase]
   url = "https://….supabase.co"     # Project URL z kroku 1.4
   key = "eyJ…service_role…"          # service_role klíč z kroku 1.5
   bucket = "radar-data"
   ```
   (Jména v `[hesla]` malými písmeny.)
5. **Deploy** → počkej 2–5 min. Dostaneš adresu `https://….streamlit.app`.

---

## Krok 4 — Použití

1. Otevři adresu → přihlas se.
2. Poprvé nahraj exporty z Pohody, nech zaškrtnuté **„💾 Zapamatovat tato data"**
   → uloží se do Supabase.
3. Klikni **▶️ Spustit analýzu**.
4. Příště zvol v panelu **„Použít poslední nahraná data"** — předvyplní se
   poslední nahrání, dokud nenahraješ nová.
5. Telefon: prohlížeč → Sdílet → **Přidat na plochu** = ikona jako appka.

---

## Poznámky

- **Soukromí:** appka i Supabase běží v cloudu. Faktury jsou v **privátním**
  Supabase bucketu (přístup jen přes service_role klíč v Secrets). Appku chrání
  heslo. Pro maximální soukromí by byl nejlepší vlastní/EU server — to už ale
  není „zdarma".
- **service_role klíč je mocný** (plný přístup k tomu Supabase projektu) — proto
  doporučuju samostatný projekt jen pro radar. Klíč drž jen v Secrets, nikam jinam.
- Bez sekce `[supabase]` appka taky funguje, ale na cloudu by se data po uspání
  ztratila. Lokálně na PC `[supabase]` nepotřebuješ.
