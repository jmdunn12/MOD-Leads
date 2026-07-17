# Lead Assignment Engine — Web App Build Prompt (Claude Code)

## What I want

I have a working Python lead-assignment engine (`lead_engine.py`, attached). It runs in a chat session where I upload spreadsheets and get back an assigned board.

**Rebuild it as a single-page web app that I can save as a bookmark and open like any website.**

Constraints that matter to me:
- I'm not technical. No terminal, no Python, no installs on my end.
- It must open from a bookmark on my laptop and my phone.
- Our sales data can't be uploaded to a server. Everything must run inside my browser.
- I use it 2–3 times a day, every shift. It has to be fast to operate.

### On the "URL" part — read this first

I originally asked for "the whole tool in a URL." That specific thing doesn't work: Chrome and Firefox both block opening `data:` URLs from a bookmark, so a self-contained URL string won't load. Don't try to build that.

Instead:
1. Build the whole tool as **one self-contained `index.html` file** — all HTML, CSS, and JavaScript in that one file, no backend, no server, no build step required to run it.
2. Deploy it to a **free static host** (Cloudflare Pages, Netlify Drop, or GitHub Pages — pick whichever needs the fewest steps from me and walk me through it).
3. I bookmark that URL. Done.

Also give me the same file working as a plain `file:///` bookmark, in case I want it with no hosting at all.

---

## Architecture

- **One file**: `index.html`. Everything inline.
- **100% client-side.** No fetch to any backend. Files I pick never leave the browser.
- **Libraries via pinned CDN `<script>` tags** (fine — the page is hosted). Use **ExcelJS** for reading and writing `.xlsx` (it supports cell styling; SheetJS community edition does not). If you can bundle the libs inline without making this fragile, even better — then it works offline.
- Vanilla JS or a tiny framework, your call. Don't add a toolchain I have to maintain.
- Must work on a phone browser: responsive, tap targets, no hover-only controls.

---

## What the tool does

Each shift I:
1. Pick the slot: **AM / PM / EVE**
2. Upload that shift's **lead sheet** (SAP "Resource Planning" `.xlsx` export)
3. For PM/EVE, upload the **prior shift's board** so drives chain from where reps already are
4. Check off **which reps are available**
5. Hit Run

It scores every eligible rep against every open lead, assigns one lead per rep via constrained optimization, and gives me back a formatted `.xlsx` board plus an on-screen table.

---

## Big UX improvements over the Python version (build these in)

1. **Reference files are uploaded once, not every shift.** The power ranking, 60-day sheet, and rep phone list change every few weeks. Parse them and cache the extracted numbers in `localStorage`. Each shift I should only need to add the lead sheet.
2. **Roster is checkboxes, not typing.** Build the rep list from the cached reference data. This kills an entire class of bug — I've typo'd "titus smtih" and "stepehn bryan" and the Python version silently dropped them.
3. **Save roster presets.** Most shifts reuse a similar crew. Let me save/load named rosters.
4. **Unknown cities are handled in the UI, not in code.** Today, a new town means editing the script. Instead: detect unknown cities, show a small form ("Broadlands — enter coordinates"), save to `localStorage`, persist forever. Optionally offer a lookup against OpenStreetMap Nominatim (CORS-friendly) with manual entry as the fallback — never block the run on a network call.
5. **Settings export/import.** One button to download my saved geocodes/aliases/presets as JSON, one to load it. That's how I move it between devices and how I don't lose everything if I clear my browser.
6. **Show me the flags, not raw output.** Long drives, thin data, unfilled leads, note-pins, duplicates — surfaced plainly on screen.

---

## Data files and how to parse them

### 1. Lead sheet (uploaded each shift)
- Sheet name: `SAPUI5 Export`
- Columns: `Transaction ID`, `Customer Name`, `City`, `State`, `Market`, `Product`, `Campaign`, `Start Time`, `Notes`, `Service Provider`, `Set Type Description`, `Start Date/Time`, `Created On`, `Status`
- One row per lead.
- **`Service Provider` is often pre-filled by SAP with placeholder reps. Ignore those and assign fresh** — but read the original value first (see the cancelled-lead rule below).

### 2. Power ranking — 2-week per-product stats
- Sheet name: `2 Weeks-Overall`, **no header row**
- Col 0 = First Name, col 1 = Last Name, col 2 = Team
- Repeating per-product blocks at these starting columns:
  - Totals = 3, Windows = 19, Doors = 35, Siding = 51, Gutters = 67, Roofing = 83
- Within each block: `+0` = Leads Ran, `+12` = Projected Net % Sold (close rate), `+15` = Projected NSLI (net $/lead)
- Rep rows start at index 5
- Current file: `Rep_Performance_By_Product_By_Source_-_v3_0_9_-_Pulled_2026-07-06_-_Outside.xlsx`

### 3. 60-day Sales Efficiency
- Sheet name: `Sales Efficiency All`, header at row index 1
- Columns used: `Sales Representative Name`, `Issued`, `Close %`, `Closed $`, `Average Sale`, `Rescinded Value`, `Cancelled Value`
- `net = Closed $ − Rescinded Value − Cancelled Value`
- `NSLI_60 = net / Issued`
- `volume = net`
- Skip the row where the name is `Total`
- Current file: `Sales_Efficiency__5_.xlsx`

### 4. Rep phone / home list
- Sheet name: `Sheet1`
- Columns: `Sales Representatives DMV`, `Location` (format `"Bowie,Md"` — split on comma, take the city)
- Normalize `" bch"` → `" beach"`
- Current file: `Copy_of_Rep_Phone_List.xlsx`

### 5. Prior board (PM/EVE only)
Either a board this tool produced, or the manager's own edited version. Read `Service Provider` + `City` per row to map each rep to where they ended that slot. **Skip rows with no city** (bench rows) so those reps correctly fall back to home.

---

## Skip rules — apply before any scoring

| Rule | Signal | Action |
|---|---|---|
| Wrong market | `Market` ≠ `"DC Metro"` | Drop from board |
| Bath-only | `Product` = `"BATHSYSTEM"` | Drop — bath goes to a separate crew |
| **Cancelled** | `Service Provider` pre-filled as `"James Dunn"` | Drop — this is how a cancellation shows up |
| **Open status** | `Status` = `"Open"` | Leave unassigned — not booked yet |
| **Install Consultant** | `Campaign` = `"Install Consultant"` | Skip — handled by install consultants, not sales reps |

All case-insensitive, trimmed. **Read the original `Service Provider` before you blank the column**, or the James Dunn rule silently stops working. Report counts of each skip type after a run.

### Always-excluded reps (never assign, regardless of roster)
`raymond hieronimus`, `stephanie poteet`, `david golladay`, `gary halford`, `neil waranch`

---

## Scoring

### Constants (keep these exact)
```
K         = 12      // how hard to lean on 60-day vs 2-week product data
K2        = 8       // stabilizes a rep's 60-day NSLI toward team average
VOL_ADD   = 10000   // $ added to the top-volume rep; volume as its own factor
FLOOR     = 5000    // minimum Eff. NSLI to be eligible for a BIG lead
DRIVE_PEN = 40      // $ penalty per drive-minute (drive is the lightest factor)
CEIL      = 120     // hard max drive minutes
```

### Effective NSLI
```
GLOBAL_60      = (sum of all reps' net) / (sum of all reps' issued)
rep_prior(r)   = (issued_60 × NSLI_60 + K2 × GLOBAL_60) / (issued_60 + K2)
eff(r, prod)   = (leads_2wk × nsli_2wk + K × rep_prior) / (leads_2wk + K)
```
If a rep has no 60-day row, fall back to a team per-product prior computed as a lead-weighted average of `nsli_2wk` across available reps.

For combo products: `base_val = mean(eff(r, p) for p in components)`

### Volume (standalone additive factor)
```
vol_score = VOL_ADD × (rep_60day_net / max_60day_net_among_available)
```
This exists specifically so high-volume reps still get leads when their product-specific fit is soft. Do not turn it back into a multiplier.

### Drive
```
distance   = haversine(origin, lead_city)          // miles
minutes    = distance × 1.3 / 35 × 60              // road factor, 35 mph
penalty    = DRIVE_PEN × minutes
```
Estimated, not live routing. Say so in the UI.

### Total
```
score = base_val + vol_score − DRIVE_PEN × minutes + tiebreak
tiebreak = 1e-3 × close% + 1e-7 × nsli + 1e-10 × issued
```
Tiebreak uses 60-day data when present, otherwise the ranking's Totals block.

### Eligibility
- If drive is unknown (city not geocoded) or > `CEIL` → ineligible for that lead
- If tier is BIG and `base_val < FLOOR` → ineligible for that lead

---

## Opportunity tiers

Assignment runs **tier by tier: BIG → MEDIUM → SMALL.** Big leads claim the best reps first; used reps leave the pool.

Evaluate in this order, first match wins:

**BIG**
- Any siding product or component
- 3+ product combo
- 7+ windows
- 3+ doors
- 5+ windows AND 2+ doors

**MEDIUM**
- Roofing (without siding)
- Gutter in a combo (2+ components)
- Exactly 6 windows
- 2 doors
- 1 front-entry door
- Any combo containing windows (window-combo floor — never SMALL)

**SMALL**
- Everything else (≤5 standalone windows, single storm/sliding door, etc.)

---

## Parsing counts out of the free-text Notes field

Port these behaviors exactly — each one exists because it broke on a real lead sheet.

**Setup**
- Lowercase the notes, and **cut everything after `"pre-cust"`** (that's prior-customer history, not this job).

**Windows**
- Word numbers count: one–fifteen, twenty
- Ranges: `"6-8 windows"` → take the max (8). Also `to`, `through`, en-dash
- `"6 windows"`, `"6 vinyl windows"` (allow up to 2 filler words between number and "windows")
- Shorthand: `"6 win"`
- Bow/bay: `"bow window and 3 additional"` → 1 + 3 = 4. Without "additional/more/other", take the max rather than the sum
- Sanity clamp: 1–40
- **Fallback**: if the product code contains `WIN` but count parsed to 0, look for a bare range, then a bare number

**Doors**
- First un-glue digits from words: `"2storms"` → `"2 storms"`
- Count per type separately, then sum across distinct types:
  - front/entry: `fed`, `front entry door`, `front door`, `entry door`, and `"N ed"`
  - storm: `storm` / `storms` / `storm doors`
  - sliding: `sgd`, `sliding glass door`, `sliding patio door`, `patio door`, `sliding door`
- Also capture a generic `"N doors"` count; final count = `max(generic, front + storm + sliding)`
- **A bare mention of "door" with no number does NOT count as 1.** A *typed* mention with no number does default to 1
- Sanity clamp: 1–20
- Track which types appeared — the tier rule needs to know if a front-entry door is present

**Product tags for the Opportunity column**
- Detect siding / roof / gutter from the product code tokens **or** the notes
- **Use word boundaries** — `\broof` — otherwise "proof of ownership" registers as a roof lead. This was a real bug.

**Opportunity string format**: `"8W 1D 1 siding"` — window count, door count, then tags for siding / roof / Gutt. Em-dash if nothing parsed.

---

## Note-pin rule (also how late leads work)

If a lead's Notes says any of `send`, `assign`, `give`, `have`, `book` followed by a rep name, or `"[rep] to run"` — **pin that rep to that lead before optimizing**, and pull them out of the pool. Match on first OR last name against the available roster; only pin when the match is unambiguous.

Real examples from live sheets: `"Frank hill to run lead."` and `"Send Stephen Bryan"`.

This is also the late-lead workflow: add the row, put "Send [rep]" in notes, re-run. **Ideally the UI lets me pin a rep to a lead directly** with a dropdown, so I don't have to edit the spreadsheet at all — that's a real improvement over the Python version.

Mark pinned rows distinctly on the board.

---

## Drive origin chaining

- **AM** → every rep starts from their **home city**
- **PM** → rep starts from their **AM lead city**, falling back to home
- **EVE** → **PM lead city** → **AM lead city** → home

Implementation: accept an ordered list of prior boards, most recent last; later entries overwrite earlier ones.

**Known weakness, flag it in the UI**: a rep who ran a far-flung morning lead gets dragged from there, which sometimes produces a 100+ minute afternoon drive. Show the origin city next to the drive time so I can see *why* a number is big. (Don't try to fix this by capping — I want to see it.)

---

## Product → scoring components

```js
const PRODC = {
  'WINDOWS': ['Windows'],
  'DOORS': ['Doors'],
  'ROOFING': ['Roofing'],
  'SIDING': ['Siding'],
  'GUTTERS': ['Gutters'],
  'COMBO_WIND_DOOR': ['Windows','Doors'],
  'COMBO_WIN_DOR_ROOF': ['Windows','Doors','Roofing'],
  'COMBO_S_W_G_D': ['Siding','Windows','Gutters','Doors'],
  'COMBO_WIN_DOOR_GUT': ['Windows','Doors','Gutters'],
  'COMBO_SID_GUT_ROOF': ['Siding','Gutters','Roofing'],
  'COMBO_BATH_WIN_DOR': ['Windows','Doors'],
  'COMBO_WIN_SID_ROOF': ['Windows','Siding','Roofing'],
  'COMBO_WIN_ROOFING': ['Windows','Roofing'],
  'COMBO_SIDING_WIND': ['Windows'],
  'COMBO_BATH_W_D_G_S': ['Windows','Doors'],
  'COMBO_SID_WIN_DOOR': ['Windows','Doors'],
  'COMBO_WIND_GUTTER': ['Windows'],
  'COMBO_SIDING_GUTT': ['Siding','Gutters'],
  'COMBO_DOOR_ROOFING': ['Doors','Roofing'],
  'COMBO_S_W_G_D_R': ['Siding','Windows','Gutters','Doors','Roofing'],
  'COMBO_SID_DOOR_GUT': ['Siding','Doors','Gutters'],
  'COMBO_SIDING_DOOR': ['Siding','Doors'],
  'COMBO_GUT_ROOFING': ['Gutters','Roofing'],
  'COMBO_SID_ROOFING': ['Siding','Roofing'],
};
// Unknown product → default ['Windows'], and surface it as a flag so I can tell you the right mapping.
```

Bath components inside a combo are kept on the board but **scored on the non-bath components only** (there's no bath column in the power ranking). Flag those rows.

Combo codes use single letters: `W`=Windows, `D`=Doors, `G`=Gutters, `S`=Siding, `R`=Roofing. When a brand-new combo shows up, parse the tokens and let me confirm the mapping in the UI — then persist it.

---

## Name aliases

I type names casually. Map to canonical:

```js
const ALIASES = {
  'al kowalski':'wojciech kowalski', 'aj doroci':'ajet doroci', 'ben sands':'benjamin sands',
  'ray wander':'raymond wander', 'frank hill':'franklin hill',
  'mary beth heller':'mary-beth heller', 'marybeth heller':'mary-beth heller',
  'nick wittman':'nicholas wittman', 'jeff kaelin':'jeffrey kaelin',
  'jp feeney':'john-paul feeney', 'john paul feeney':'john-paul feeney', 'johnpaul feeney':'john-paul feeney',
  'josh brown':'joshua brown', 'josh b':'joshua brown',
  'chris mercer':'christopher mercer', 'sam ludwig':'samantha ludwig',
  'steve forss':'steven forss', 'jaems archy':'james archy',
  'zach diffenderfer':'zachary diffenderfer', 'zach diff':'zachary diffenderfer',
  'hunter w':'hunter willson', 'jacob s':'jacob szczepanik',
  'stepehn bryan':'stephen bryan', 'titus smtih':'titus smith',
};
```

Apply to: roster input, prior-board reading, home list, note-pin matching. Checkbox rosters make most of these unnecessary going forward, but prior boards and notes still carry loose spellings.

### Home overrides (reps not yet in the phone-list export)
```js
const HOME_OVERRIDES = { 'jacob szczepanik': 'crofton' };
```
Make this editable in the UI and persisted.

---

## Geocodes

Drive time needs a `(lat, lon)` for every lead city and rep origin. Port this dictionary as the seed, then let the UI grow it (persisted to `localStorage`).

```js
const GEO = {
  'washington':[38.9072,-77.0369], 'baltimore':[39.2904,-76.6122], 'silver spring':[38.9907,-77.0261],
  'rockville':[39.0840,-77.1528], 'ellicott city':[39.2673,-76.7983], 'bethesda':[38.9847,-77.0947],
  'upper marlboro':[38.8157,-76.7497], 'sykesville':[39.3737,-76.9686], 'reisterstown':[39.4690,-76.8294],
  'bowie':[38.9427,-76.7302], 'boyds':[39.2143,-77.3216], 'olney':[39.1532,-77.0669], 'odenton':[39.0840,-76.7000],
  'waldorf':[38.6246,-76.9300], 'temple hills':[38.8126,-76.9400], 'gaithersburg':[39.1434,-77.2014],
  'lanham':[38.9676,-76.8636], 'frederick':[39.4143,-77.4105], 'owings mills':[39.4196,-76.7802],
  'clinton':[38.7651,-76.8983], 'germantown':[39.1732,-77.2700], 'alexandria':[38.8048,-77.0469],
  'winchester':[39.1857,-78.1633], 'fairfax station':[38.8043,-77.3203], 'district heights':[38.8576,-76.8894],
  'fort washington':[38.7099,-77.0300], 'university park':[38.9701,-76.9447], 'manassas':[38.7509,-77.4753],
  'accokeek':[38.6671,-76.9886], 'davidsonville':[38.9290,-76.6308], 'deale':[38.7793,-76.5494],
  'westminster':[39.5754,-76.9958], 'potomac':[39.0182,-77.2086], 'derwood':[39.1287,-77.1539],
  'falls church':[38.8823,-77.1711], 'millersville':[39.0626,-76.6299], 'thurmont':[39.6237,-77.4111],
  'clifton':[38.7807,-77.3866], 'columbia':[39.2037,-76.8610], 'centreville':[38.8401,-77.4291],
  'leesburg':[39.1157,-77.5636], 'great mills':[38.2604,-76.4969], 'gwynn oak':[39.3287,-76.7305],
  'hollywood':[38.3393,-76.5566], 'upper falls':[39.4015,-76.4133], 'vienna':[38.9012,-77.2653],
  'chevy chase':[38.9682,-77.0728], 'chevrolet':[38.9682,-77.0728],
  'severna park':[39.0743,-76.5500], 'herndon':[38.9696,-77.3861], 'owings':[38.6900,-76.6100],
  'bel air':[39.5359,-76.3500],
  'annandale':[38.8304,-77.1964], 'college park':[38.9807,-76.9369], 'elkridge':[39.2126,-76.7136],
  'fulton':[39.1526,-76.9230], 'gainesville':[38.7959,-77.6147], 'king george':[38.2682,-77.1839],
  'lexington park':[38.2668,-76.4527], 'woodbridge':[38.6582,-77.2497],
  'brunswick':[39.3140,-77.6249], 'finksburg':[39.4940,-76.8800], 'great falls':[38.9979,-77.2880],
  'hagerstown':[39.6418,-77.7200], 'windsor mill':[39.3334,-76.7800],
  'brandywine':[38.6976,-76.8483], 'ijamsville':[39.3132,-77.3589], 'lake frederick':[39.0501,-78.1300],
  'mount airy':[39.3762,-77.1547], 'suitland':[38.8487,-76.9197],
  'catonsville':[39.2720,-76.7319], 'clarksville':[39.2037,-76.9469], 'keedysville':[39.4862,-77.6961],
  'locust grove':[38.3457,-77.7600], 'alexandria city':[38.8048,-77.0469], 'catlett':[38.6543,-77.6386],
  'chantilly':[38.8943,-77.4311], 'glen burnie':[39.1626,-76.6247], 'haymarket':[38.8124,-77.6361],
  'stafford':[38.4221,-77.4083],
  'halethorpe':[39.2293,-76.6705], 'lutherville timonium':[39.4243,-76.6190], 'timonium':[39.4385,-76.6094],
  'sterling':[39.0062,-77.4286],
  'ashton':[39.1532,-77.0030], 'brookeville':[39.1840,-77.0586], 'clarksburg':[39.2380,-77.2786],
  'graysonville':[38.9460,-76.2107],
  'damascus':[39.2887,-77.2036], 'mc lean':[38.9339,-77.1773], 'mclean':[38.9339,-77.1773],
  'mechanicsville':[38.4438,-76.7383],
  'montgomery village':[39.1768,-77.1953], 'rosedale':[39.3243,-76.5097], 'sparrows point':[39.2230,-76.4358],
  'cheverly':[38.9290,-76.9158], 'saint leonard':[38.4757,-76.5125], 'seat pleasant':[38.8968,-76.9047],
  'warrenton':[38.7135,-77.7950], 'crofton':[39.0007,-76.6850],
  'hanover':[39.1888,-76.7236], 'oakton':[38.8810,-77.3000],
  'parkton':[39.6204,-76.6597], 'takoma park':[38.9779,-77.0075],
  'burke':[38.7934,-77.2717], 'essex':[39.3093,-76.4750], 'gambrills':[39.0857,-76.6622], 'gambrels':[39.0857,-76.6622],
  'new windsor':[39.5417,-77.1067], 'queenstown':[38.9982,-76.1597], 'woodsboro':[39.5337,-77.3175],
  'forest hill':[39.5840,-76.3950], 'severn':[39.1372,-76.6983], 'broadlands':[39.0157,-77.5347],
  'aldie':[38.9719,-77.6403], 'havre de grace':[39.5495,-76.0902], 'lovettsville':[39.2735,-77.6386],
  'stephenson':[39.2196,-78.0000],
  'capitol heights':[38.8851,-76.9136], 'reston':[38.9586,-77.3570],
  'leonardtown':[38.2904,-76.6355], 'baldwin':[39.5093,-76.4919], 'white plains':[38.5893,-76.9789],
  'abingdon':[39.4668,-76.2986], 'brambleton':[38.9821,-77.5389], 'hampstead':[39.6046,-76.8497],
  'huntingtown':[38.6193,-76.6383], 'nottingham':[39.3743,-76.4358],
  'bealeton':[38.5738,-77.7708], 'edgewood':[39.4187,-76.2944], 'round hill':[39.1362,-77.7597],
  'boonsboro':[39.5074,-77.6522], 'bristow':[38.7573,-77.5397], 'front royal':[38.9187,-78.1944],
  'aberdeen':[39.5096,-76.1641], 'dundalk':[39.2507,-76.5205], 'linthicum heights':[39.2054,-76.6525],
  'trappe':[38.6579,-76.0577],
  'nokesville':[38.6987,-77.5786], 'belcamp':[39.4726,-76.2436], 'churchville':[39.5640,-76.2483],
  'la plata':[38.5293,-76.9753], 'manchester':[39.6612,-76.8850], 'union bridge':[39.5690,-77.1772],
  'edgewater':[38.9568,-76.5494], 'grasonville':[38.9460,-76.2107], 'kensington':[39.0259,-77.0764],
  'north potomac':[39.0954,-77.2480], 'pikesville':[39.3743,-76.7252], 'randallstown':[39.3673,-76.7950],
  'ashburn':[39.0437,-77.4875], 'cockeysville':[39.4812,-76.6427], 'hillsboro':[39.2065,-77.7252],
  'hughesville':[38.5318,-76.7886], 'jeffersonton':[38.6385,-77.8800], 'lorton':[38.7043,-77.2278],
  // rep home cities
  'pasadena':[39.1071,-76.5694], 'annapolis':[38.9784,-76.4922], 'woodbine':[39.3362,-77.0639],
  'riva':[38.9568,-76.5808], 'fairfax':[38.8462,-77.3064], 'laurel':[39.0993,-76.8483],
  'beltsville':[39.0345,-76.9075], 'middle river':[39.3343,-76.4391], 'greenbelt':[38.9959,-76.8755],
  'springfield':[38.7893,-77.1872], 'crownsville':[39.0285,-76.6041], 'stevensville':[38.9818,-76.3314],
  'hyattsville':[38.9559,-76.9456], 'chesapeake beach':[38.6857,-76.5336], 'parkville':[39.3784,-76.5394],
  'arnold':[39.0326,-76.5030], 'oxon hill':[38.8013,-76.9869], 'burtonsville':[39.1116,-76.9325],
  'arlington':[38.8799,-77.1068], 'chester':[37.3549,-77.4413],
};
```

Some cities appear twice by design (`mc lean`/`mclean`, `gambrills`/`gambrels`) because SAP spells them inconsistently. Keep both.

---

## The assignment algorithm

Python used `scipy.optimize.linear_sum_assignment`. You need the equivalent in JS.

1. Filter leads per the skip rules
2. Resolve note-pins; remove pinned reps from the pool
3. For each tier in `[BIG, MEDIUM, SMALL]`:
   - Build a matrix: rows = leads of that tier, cols = remaining reps
   - Fill with `score`; use a large negative sentinel (`-1e9`) for ineligible pairs
   - Run the Hungarian / Jonker-Volgenant algorithm to **maximize** total score
   - Treat any assignment whose cell is `≤ -1e8` as **not assigned** (leave the lead unfilled)
   - Remove assigned reps from the pool before the next tier
4. Write the board

**Implementation notes:**
- Must handle **rectangular** matrices (leads ≠ reps) — pad internally
- Use a sentinel, **not `-Infinity`** — infinities produce `NaN` and silently corrupt the solve
- Result must be deterministic; that's what the tiebreak term is for
- **Write unit tests for this.** It's the one component where a subtle bug produces plausible-looking but wrong boards, which I would not catch by eye. Test: square case, more leads than reps, more reps than leads, all-ineligible row, known-optimal fixture.

---

## Output

### On screen
A table: Customer, City, Tier, Opportunity, Rep, Origin, Drive (min), Eff. NSLI, Flags. Plus a bench list, and a summary line ("12 assigned, 2 unfilled, 1 pinned, 3 skipped: 1 cancelled / 2 open").

Make long drives, unfilled leads, and pinned rows visually obvious. Sortable columns would be nice.

### Downloaded `.xlsx`
Must match what the Python version produced — I forward this file to other people.

Take the uploaded lead sheet, drop the filtered rows, fill in `Service Provider`, and append:
- **Tier** — BIG / MEDIUM / SMALL
- **Opportunity** — e.g. `8W 1D 1 siding`
- **Rep Home** — the origin the drive was computed from
- **Est Drive (min)**
- **Eff. NSLI ($)**
- **Assignment Flag** — origin, 60-day volume, long-drive warning, thin-data warning, bath-component note, pinned status, or the reason a lead is unfilled

**Formatting:**
- Header row: navy fill `1F4E78`, white bold Arial, centered, wrapped, row height 30
- Autofilter across all columns; freeze the header row
- Thin gray `808080` borders on every used cell
- Auto-fit column widths: content + 2, capped at 34 — except Notes and Assignment Flag, capped at 55 with wrap
- Row fills:
  - Green `E2EFDA` — BIG lead
  - Yellow `FFF2CC` — drive > 75 min, or thin 2-week product data
  - Red `F8CBAD` — unfilled / uncovered
  - Light blue `D9E1F2` — note-pinned
  - Light gray `F2F2F2` — bench rows
- **BENCH section** below the last lead (after a blank row): unassigned roster reps, sorted by 60-day volume descending, showing name, home, and volume

**Assignment flag contents** (port from Python):
- `from: {origin}`
- `60-day vol $XXXK`
- `verify drive ~NNm` when drive > 75
- `thin 2-wk product volume — leaning on 60-day overall` when every scored component has < 5 leads in the 2-week data
- `limited 60-day volume (N issued) — verify for BIG lead` when tier is BIG and issued < 10
- `contains bath component — scored on windows/doors`
- `PINNED per note -> {rep}`
- Unfilled: either `UNFILLED — reachable reps all below $5,000 BIG floor`, or `UNCOVERED — only reachable: {names} (taken)`, or `UNCOVERED — no rep within 120m`

---

## Design

Read the frontend-design guidance and give this a real point of view — but **it's an operations tool used under time pressure three times a day, not a landing page.** Speed and legibility beat personality. Dense, scannable, no scroll-jacking, no decorative animation. The board table is the hero. Keyboard focus visible, works at phone width.

---

## Acceptance tests

Validate against real fixtures before you tell me it's done:

1. **Skip rules**: a sheet with a James Dunn row, an Open row, and an Install Consultant row → all three excluded, counted, and reported.
2. **Note-pin**: a lead whose notes say `"Frank hill to run lead."` → Franklin Hill pinned, everyone else optimized around him.
3. **Notes parsing**: `"6-8 windows"` → 8. `"2storms"` → 2 storm doors. `"proof of ownership"` → NOT a roof lead. `"bow window and 3 additional"` → 4.
4. **Chaining**: same lead sheet run as AM vs PM-with-prior-board → different origins, and PM origins match the prior board's cities.
5. **Hungarian**: the unit tests above.
6. **Parity check**: run a historical shift through both this app and `lead_engine.py` and diff the boards. They should match. If they don't, the Python version is the source of truth — the tool was validated against my hand-picks across ~40 shifts.

---

## How I actually work (so the UI fits it)

- I upload the sheet and tell you the slot and available reps. Sometimes I paste a screenshot of morning locations instead of a file.
- I review the board and often override a few picks based on knowing my reps. **Those overrides are shift-specific judgment — never hard-code them into the logic.** Only general repeatable rules belong in the engine (like the James Dunn / Open / Install Consultant rules).
- I want flags for anything needing a human call: long drives, thin data, unfilled leads, duplicate leads, reschedule notes, unknown cities, unknown products.
- Bath leads and virtual appointments go to different people entirely — keep them off my board but tell me they were skipped.

---

## Deliverables

1. `index.html` — the whole tool, one file, client-side only
2. Deployed to a free static host, with the URL to bookmark
3. **Numbered, non-technical steps** for: how to deploy it, how to update it later, and how to back up my settings JSON
4. A short "what to do when something new shows up" note — new city, new product code, new rep
5. The test suite from the acceptance section

Keep `lead_engine.py` in the repo as the reference implementation and the parity baseline.
