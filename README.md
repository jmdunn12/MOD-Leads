[README.md](https://github.com/user-attachments/files/30134673/README.md)
# MOD-Leads# Lead Assignment Engine — how to use it

Everything runs inside your browser. Nothing you upload leaves the machine.

## Files in this folder

- **index.html** — the tool. One file, no build step, no server needed.
- **tests.html** — the acceptance test suite. Open it after any change to `index.html`.
- **lead_engine.py** — the original Python engine, kept as the parity baseline.
- **CLAUDE_CODE_PROMPT_WEB.md** — the build spec.

---

## Try it locally first (no deploying yet)

1. Double-click **index.html**. It opens in your default browser.
2. **Setup panel** — upload once:
   - Power ranking (`Rep_Performance…Outside.xlsx`)
   - 60-day Sales Efficiency (`Sales_Efficiency….xlsx`)
   - Rep phone list (`Copy_of_Rep_Phone_List.xlsx`)
3. Chip on the top-right turns green: "reference data: loaded". Setup panel collapses.
4. Every shift after that:
   - Pick **market** — DC Metro, Hampton Roads, or Richmond
   - Pick **AM / PM / EVE**
   - Drop this shift's lead sheet (SAP Resource Planning export)
   - For PM: also drop the AM board. For EVE: drop AM then PM, most recent last.
   - Check off the **field reps** working today
   - Check off the **bath crew** (Kappa) if bath leads are on the sheet
   - Hit **Run assignment**
5. Review the board on-screen — field board on top, **bath board** below if applicable — click **Download .xlsx** for the file you forward.

Everything cached on your laptop stays on your laptop — same story on your phone.

### What the market picker does

The **Team** column on your power ranking assigns every rep to one of five teams:
`Beta` and `Delta` are DC Metro, `Gamma-VAB` is Hampton Roads, `Gamma-RVA` is Richmond, `Kappa` is the bath crew.

When you pick a market:
- Only leads in that market's Market column are considered
- The field-roster checklist shows only that market's reps
- Kappa reps never appear as field candidates in any market
- Roster presets are per-market

### How bath leads flow

Bath leads (`Product = BATHSYSTEM`) come off the field board automatically and land on a separate **Bath Board** section:

1. **Kappa first** — bath leads fill from your checked Kappa reps.
2. **Field overflow** — if there are more bath leads than Kappa can cover, the extras go to reps who ended up on the field bench, ranked by their **bath** score. Those overflow reps disappear from the field bench so nobody's double-booked.
3. Overflow rows are highlighted yellow with an "OVERFLOW" flag — clear at a glance which reps came from field.

If bath volume routinely pulls reps off the field board, the summary line calls it out.

---

## Deploy so you can bookmark a URL

Cloudflare Pages is the fewest-steps option. If you already have a GitHub account, GitHub Pages works too — instructions for both below.

### Option A — Cloudflare Pages (recommended)

1. Go to **https://pages.cloudflare.com** and click **Sign up** (free — no credit card).
2. Once you're in, click **Create a project** → **Upload assets**.
3. Give the project a name (e.g. `mod-leads`). The URL will end up as `mod-leads.pages.dev`.
4. Drag **index.html** and **tests.html** into the drop zone. Click **Deploy site**.
5. Wait ~30 seconds. Cloudflare shows a URL like `https://mod-leads.pages.dev`.
6. Open that URL on your laptop and your phone. Bookmark it on both.

**To update it later:** on the Cloudflare project page → **Create deployment** → drag the new `index.html` in → **Deploy**. The URL stays the same.

### Option B — GitHub Pages

1. Sign in to **https://github.com** (or make a free account).
2. Click **+** (top-right) → **New repository**. Name it `mod-leads`. Tick **Public**. Click **Create**.
3. On the new empty repo, click **uploading an existing file** (the link in the middle of the page). Drag `index.html` and `tests.html` in. Click **Commit changes**.
4. Click **Settings** (top of the repo) → **Pages** (left sidebar). Under **Branch**, pick `main` and `/ (root)`. Click **Save**.
5. Wait one minute, refresh. The green box at the top gives you a URL like `https://YOURNAME.github.io/mod-leads/`. Bookmark it.

**To update it later:** open the repo → click on `index.html` → the pencil icon → paste in the new contents → **Commit changes**. Live URL updates in about a minute.

### Option C — no hosting at all

If you don't want to deploy anything: put `index.html` on your desktop and double-click it. It works. Downside: the "look up city coordinates" button uses the internet, and some browsers block that from a local file — you can still type coordinates in by hand.

---

## Back up your settings

The **Back up settings** button downloads a small JSON file with every custom geocode, alias, home override, product mapping, and roster preset you've saved. Do this every couple of weeks.

To move to a new device or restore after clearing your browser: click **Restore settings** and pick that JSON file. Then re-upload the three reference spreadsheets (those aren't in the JSON — they change every week anyway).

---

## When something new shows up

- **New city** — the app catches it and shows a small form: "Broadlands — latitude, longitude". Click **Look up** (or type from Google Maps). Click **Save**. Persists forever. Coordinates are keyed by `city|state` so Mechanicsville, VA and Mechanicsville, MD stay distinct.
- **New product code** — the row still runs, scored as Windows, and gets flagged as unknown. In Setup → **Product mappings**, add e.g. `COMBO_SID_WIN_GUT` → `Siding, Windows, Gutters`. Bath is now a valid component too — combos like `COMBO_BATH_WIN_DOR` are scored on Bath + Windows + Doors.
- **New rep** — they show up automatically once they're on the power ranking AND the phone list. If they're on the ranking but not the phone list yet, use Setup → **Home overrides** — format the home as `city|state` (e.g. `crofton|md`).
- **Rep name spelled wrong on prior board** — Setup → **Name aliases**, e.g. `frank hill` → `franklin hill`.
- **New market or team** — the market picker is fixed at three (DC Metro / Hampton Roads / Richmond). If Thompson Creek adds a fourth, tell me and I'll edit `TEAMS_BY_MARKET` in the engine.
- **Drive cap or BIG floor feels wrong for a market** — Setup → **Per-market caps**. Defaults are 120 min drive cap and $5,000 BIG floor for every market; override either per market.
- **Wrong-state city warning** — if a lead says Norfolk, MD (typo), the app falls back to Norfolk, VA (the real one) and flags the row so you can verify. Silent wrong-state matches are the failure mode we're avoiding.

---

## Verify a change didn't break anything

Open `tests.html` on the deployed URL (e.g. `https://YOUR-USERNAME.github.io/mod-leads/tests.html`). All 196 tests should be green. If any go red, don't ship that version.

Also: run the same historical shift through both the app and `lead_engine.py` and diff the boards. If they differ, `lead_engine.py` wins — it's the parity baseline.
