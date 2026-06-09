# KSA Ads Dashboard — Project Context

## What This Is

Two HTML dashboards for KS Agency (Kat Schmoyer's integration agency) and their ad management clients:
1. **Internal Dashboard** — for Rebecca's team to manage all clients' ad campaigns
2. **External Dashboard** — client-facing, one URL per client, shows their ad performance snapshot

Built by Rebecca Rice as part of her AI Agent Consulting Service.
**Total budget: $500 for both dashboards.**
Keep scope tight. These are data display dashboards, not full applications.

---

## Branding

**Before writing any HTML or CSS, read the KSA branding skill:**
`~/claude-skills/ksa-branding.md`

Key values: warm cream background `#FFF6EF`, pink accent `#F77F9F`, Bebas Neue for labels, MackinacPro Bold/Book for headings and body (fallback: Lora from Google Fonts if font files not in `/fonts/`).

**First task every new session:** confirm brand test file exists and was approved by Rebecca. If not, build `brand-test.html` first before touching any dashboard file.

---

## Notion Project Page

Lives under: SCC → Projects → AI Agent Consulting Service
Page ID: `36653d12-9f22-81f5-acee-c6dd992b8a5f`
Fetch this at session start for full context.

---

## Tech Stack

- **Dashboards:** Plain HTML/CSS/JS — no frameworks, no build step
- **Data source:** Google Sheets (master sheet, all clients)
- **Meta data pull:** Google Apps Script, runs daily at 3am Central Time
- **Hosting:** GitHub Pages with KSA subdomain
- **Font files:** Check `/fonts/` for MackinacPro-Bold.woff2 and MackinacPro-Book.woff2 before building. If not present, use Lora from Google Fonts as fallback.

---

## Google Sheet Structure

One master sheet. All clients. Tab naming convention: `[Client Name] - [Tab Type]`

Examples:
```
Clients                              ← master client list
Price History                        ← client | offer | product | price | effective date
CHECKBOX_STATE                       ← daily checkbox persistence, cleared at 3am
KSA - Daily Meta                     ← Apps Script writes here daily
KSA - Config                         ← all offers, products, prices, funnel order for KSA
KSA - What to Wear Purchases         ← one tab per offer per client
KSA - New Presets Purchases
HTP - Daily Meta
HTP - Config
HTP - What to Wear Tiny Offer Purchases
```

### Clients Tab Columns
Client Name | Slug | Meta Account ID(s) | Status | Notes

Seed with:
| Client Name | Slug | Meta Account IDs | Status |
|---|---|---|---|
| KS Agency | ksa | [Rebecca to provide — 2 IDs] | Active |
| Hope Taylor Photography | htp | [Rebecca to provide] | Active |
| KJP | kjp | [Rebecca to provide] | Active |
| Lola | lola | [Rebecca to provide] | Active |

### Price History Tab Columns
Client Name | Offer Name | Product Name | Price | Effective Date

Client Name is ALWAYS the first column.

Revenue formula for any purchase row (date-aware VLOOKUP — finds most recent price on or before transaction date):
```
=IFERROR(INDEX(Price_History[Price],MATCH(1,
  (Price_History[Client Name]=[@ClientName])*
  (Price_History[Offer Name]=[@OfferName])*
  (Price_History[Product Name]=[@ProductName])*
  (Price_History[Effective Date]<=[@Date]),0)),0)
```

### CHECKBOX_STATE Tab Columns
Date | Client Slug | Campaign Name | Action

Cleared nightly — Apps Script removes all rows where Date < today.

### Per-Client Config Tab Columns
Offer Name | Product Name | Position | Current Price | Active (Yes/No)

Position is an integer — determines display order in dashboard and purchase log column order.
When products are reordered, update Position. Dashboard always renders in position order.
Conversion rates follow the product by name, not by slot number.

### Per-Offer Purchases Tab Columns
Date | [Product 1 Name] | [Product 2 Name] | ... | Total Units | Total Revenue

- One row per day — ad specialist updates quantities throughout the day (not one row per transaction)
- Column headers match product names from Config tab for that offer, in Position order
- Total Revenue uses Price History VLOOKUP with row's date for accurate historical pricing
- Total Units = sum of all product columns for that row

### Daily Meta Tab Columns (per client)
Date | Campaign Name | Amount Spent | Impressions | Reach | Clicks | Leads | Purchases | Purchase Value | CPM | CPC | CTR | ROAS

Written by Apps Script every morning. No separate 7-day or 30-day tabs — all windows calculated dynamically from daily data.

---

## Apps Script (Code.gs)

### Script Properties (set once, never hardcoded in code)
- `META_TOKEN` — Meta System User access token
- `SHEET_ID` — Google Sheet ID
- `WEBAPP_SECRET` — shared secret for verifying web app requests from dashboard

### Daily trigger: `runDailyPull()`
- Time-based trigger, 3am Central (set timezone in Apps Script project settings)
- Reads all Active rows from Clients tab
- For each client, splits Meta Account IDs by comma and loops through each
- Calls Meta Marketing API per campaign:
  ```
  GET /{account_id}/insights
    ?fields=campaign_name,spend,impressions,reach,clicks,actions,action_values,cpm,cpc,ctr,purchase_roas
    &date_preset=yesterday
    &level=campaign
    &access_token={META_TOKEN}
  ```
- Appends rows to [Client Name] - Daily Meta (never overwrites)
- After all pulls: calls `clearYesterdayCheckboxes()`

### `clearYesterdayCheckboxes()`
Removes all rows from CHECKBOX_STATE where Date < today's date.

### Web App: `doPost(e)`
Verify WEBAPP_SECRET header on every request.

Actions:
- `save_config` → write offer/product changes to [Client] - Config tab
- `save_price_change` → append to Price History with effective date
- `save_checkbox_state` → add or remove row in CHECKBOX_STATE
- `setup_new_client` → add to Clients tab, create Daily Meta + Config tabs
- `setup_new_offer` → create Purchases tab, set headers, write Config rows, write Price History rows

---

## Scaling Rules

```
SCALE_THRESHOLD     = 1.4   // 7-day ROAS — scale or duplicate
ORANGE_THRESHOLD    = 1.25  // approaching scale, watch
BREAK_EVEN          = 1.0
TURN_OFF_THRESHOLD  = 0.75  // both signals must confirm before turning off
```

Status logic (per campaign, calculated from Daily Meta tab):
1. **new_watch** — < 3 days of spend data → always watch, never scale or turn off
2. **turn_off** (early) — 0 sales after Day 3 of watching
3. **scale_vertical** — 7-day ROAS >= 1.4 AND yesterday >= 1.0
4. **duplicate** — same as scale_vertical AND campaign name contains $50 or $100
5. **hold** — 7-day ROAS >= 1.4 but yesterday < 1.4 (soft yesterday)
6. **hold** — ROAS 1.25–1.39
7. **orange_flag** — ROAS 1.0–1.24 trending down from 14-day
8. **hold** — ROAS 1.0–1.24 trending up or stable
9. **orange_flag** — ROAS < 1.0 but yesterday >= 1.0 (possible recovery)
10. **turn_off** — ROAS < 0.75 AND yesterday < 1.0 (both confirm)
11. **orange_flag** — everything else below 1.0

Scale action text: `$X → $Y/day (+9%)`
If duplicate-eligible: `Scale 9% → $Y/day · Then D&D`

**No Hyros.** KSA uses Meta data + purchase log revenue only. All status decisions use Meta ROAS.

---

## Internal Dashboard

Single HTML file. Reads Google Sheet via Sheets API (published sheet + API key).

### Config at top of script
```javascript
const SHEET_ID   = 'YOUR_SHEET_ID_HERE';
const API_KEY    = 'YOUR_API_KEY_HERE';
const WEBAPP_URL = 'YOUR_APPS_SCRIPT_WEBAPP_URL_HERE';
const SCALE_THRESHOLD    = 1.4;
const ORANGE_THRESHOLD   = 1.25;
const TURN_OFF_THRESHOLD = 0.75;
```

### Layout
- Top bar (pink `#F77F9F`): KSA mark | "Internal Ads Dashboard" | active client name | sync status | refresh
- Left sidebar (white): client list, active = `#FFDCD9` bg + `#F77F9F` left border, "+ Add Client" at bottom
- Main content area: account overview by default, offer view when offer selected

### Account Overview (default)
- Date range buttons: 7-Day (default) | 30-Day | Custom
- Summary cards: Total Spend | Total Revenue | Overall ROAS | Total Profit
- Performance chart: day-by-day, toggle Spend & Revenue / ROAS, reference lines at 1.0 and 1.4
- Per-offer summary strip: one card per offer, ROAS + Revenue + Profit, click → Offer View

### Offer View
- Back button → Account Overview
- **Purchase Summary Strip at top:**
  - Shows last 7 days from that offer's Purchases tab
  - One cell per product in Position order: Product Name | Units Sold | Conversion Rate
  - Conversion rate = units / total tripwire (or total leads)
  - Reorders automatically when products are reordered in Settings
- Campaign tabs: Today's Actions | Orange Flags | New Campaigns | All Campaigns
- Campaign rows: checkbox | status badge | campaign name | ROAS | spend | revenue | profit | CPM | action text
- Expand row → day-by-day mini ROAS chart

### Checkboxes
- Saves state to CHECKBOX_STATE tab via Apps Script on check/uncheck
- Restores from CHECKBOX_STATE on page load (same day only)
- Resets automatically at 3am with daily pull

### Settings Panel
- Slide-in from right
- Client: name, slug (read-only after creation), Meta account IDs
- Offers: list with Edit / Change Price / Add Product / Add Offer
- Change Price prompts for effective date → writes to Price History
- Reorder products → updates Position → saves to Config
- All saves POST to Apps Script web app

---

## External Dashboard

Single HTML file. URL param: `?client=slug`

### On load
- Read `client` param → look up in Clients tab
- Not found → show friendly "Dashboard not found" message
- Found → load only that client's data

### Contents
- KSA branded header: "Your Ads Dashboard — powered by KS Agency"
- Client name + last updated timestamp
- Date toggle: 7-Day (default) | 30-Day
- Account summary cards: Spend | Revenue | ROAS | Profit
- Day-by-day ROAS line chart, break-even reference line at 1.0
- Per-offer cards: offer name | spend | revenue | ROAS | profit | mini ROAS sparkline
- Footer: "Questions about your ads? Reach out to your KSA team."

**No flags. No action recommendations. No internal data. No other clients.**

---

## Deployment

- GitHub Pages: automatic deploy on push (no build step needed for plain HTML)
- CNAME file in repo root with KSA subdomain
- Kat's team adds DNS CNAME record pointing subdomain → `[username].github.io`

---

## Build Order

1. **`brand-test.html`** — FIRST. Get Rebecca's approval before touching dashboards.
2. **Google Sheet setup** — create master sheet, seed with tab structure and sample data
3. **`apps-script/Code.gs`** — Meta API pull, daily trigger, web app endpoint
4. **`internal-dashboard.html`** — account overview, then offer view, then settings panel
5. **`external-dashboard.html`** — fork internal, strip flags, add URL param filter
6. **Docs** — Apps Script deployment guide, GitHub Pages + subdomain guide, Meta token guide
7. **`CNAME`** — GitHub Pages subdomain file

---

## Files to Create

```
/
├── brand-test.html
├── internal-dashboard.html
├── external-dashboard.html
├── fonts/
│   ├── MackinacPro-Bold.woff2     ← Rebecca to provide, or delete dir and use Lora
│   └── MackinacPro-Book.woff2
├── CNAME                          ← contains subdomain, e.g. ads.katschmoyer.com
├── apps-script/
│   ├── Code.gs
│   ├── deployment-guide.md        ← step-by-step Apps Script setup
│   ├── github-pages-guide.md      ← step-by-step GitHub Pages + subdomain
│   └── meta-token-guide.md        ← step-by-step Meta System User token
└── CLAUDE.md                      ← this file
```

---

## What NOT to Build

- No user authentication or login walls
- No database — sheet is the data store
- No React, Vue, or any framework — plain HTML/JS only
- No server-side code — all reads are client-side Sheets API calls
- No webhook integrations with Kajabi/Thrivecart
- No automated purchase logging from platform APIs
- No campaign progress tracker for external dashboard

---

## Known Clients

| Client Name | Slug | Notes |
|---|---|---|
| KS Agency | ksa | Two Meta ad accounts — IDs needed |
| Hope Taylor Photography | htp | One Meta account — ID needed |
| KJP | kjp | One Meta account — ID needed |
| Lola | lola | One Meta account — ID needed |

HTP offers confirmed from spreadsheet: "What to Wear Tiny Offer" (Tripwire $17, Bump $11, Upsell 1 $47, Upsell 2 $97, Upsell 3 $37, Upsell 4 $5), "New Presets"
