# KSA Ads Dashboard — Full Build Spec

## Project Summary

Build two HTML dashboards for KS Agency's ad management service:
1. **Internal Dashboard** — for Rebecca's team to manage all clients' ad campaigns
2. **External Dashboard** — client-facing snapshot, one URL per client

Data source: Google Sheets (master sheet)
Meta data: Apps Script pulls from Meta Marketing API daily at 3am Central
Hosting: GitHub Pages + KSA subdomain

**Budget constraint:** $500 total. Keep scope tight. No backend, no auth system, no database. HTML reads from a published Google Sheet via Sheets API.

---

## Step 0 — Before Anything Else: Brand Test

Build `brand-test.html` first. This is a single page showing:
- Color swatches for all brand colors
- Typography samples: page title, section heading, eyebrow label, body text, data number, button
- One sample summary card (cream background, white card, pink heading, orange eyebrow)
- One sample campaign row (with a status badge)
- One sample button

**Do not proceed to building dashboards until Rebecca approves the brand test.**

Load ksa-branding skill from `/mnt/skills/user/ksa-branding/SKILL.md` before writing any CSS.

---

## Step 1 — Google Sheet Structure

Create or configure a master Google Sheet with these tabs in this order:

### Tab: Clients
Columns: Client Name | Slug | Meta Account ID(s) | Status | Notes

Seed with known clients:
| Client Name | Slug | Meta Account IDs | Status |
|---|---|---|---|
| KS Agency | ksa | [Rebecca to provide - 2 IDs] | Active |
| Hope Taylor Photography | htp | [Rebecca to provide] | Active |
| KJP | kjp | [Rebecca to provide] | Active |
| Lola | lola | [Rebecca to provide] | Active |

### Tab: Price History
Columns: Client Name | Offer Name | Product Name | Price | Effective Date

Client Name is always first. Sorted by Client Name then Effective Date descending.
The revenue lookup formula for any purchase row is:
```
=IFERROR(
  INDEX(PriceHistory[Price],
    MATCH(1,
      (PriceHistory[Client Name]=[@ClientName]) *
      (PriceHistory[Offer Name]=[@OfferName]) *
      (PriceHistory[Product Name]=[@ProductName]) *
      (PriceHistory[Effective Date]<=[@Date]),
      0
    )
  ),
  0
)
```

### Tab: CHECKBOX_STATE
Columns: Date | Client Slug | Campaign Name | Action

Cleared nightly by Apps Script (removes rows where Date < today).

### Per-client tabs (created by Apps Script when client is set up):

**[Client Name] - Daily Meta**
Columns: Date | Campaign Name | Amount Spent | Impressions | Reach | Clicks | Leads | Purchases | Purchase Value | CPM | CPC | CTR | ROAS

**[Client Name] - Config**
Columns: Offer Name | Product Name | Position | Current Price | Active (Yes/No)

**[Client Name] - [Offer Short Name] Purchases**
Columns: Date | [Product 1 Name] | [Product 2 Name] | ... | Total Units | Total Revenue

Column headers for products come from the Config tab, in Position order.
Total Revenue uses Price History VLOOKUP with the row's date for accurate historical pricing.

---

## Step 2 — Apps Script (Code.gs)

### Script Properties (set once, never hardcoded)
- `META_TOKEN` — Meta System User access token
- `SHEET_ID` — Google Sheet ID
- `WEBAPP_SECRET` — shared secret for web app requests from dashboard

### Daily trigger function: `runDailyPull()`
- Runs at 3am Central (set via time-based trigger, Central timezone)
- Reads all rows from Clients tab where Status = Active
- For each client, splits Meta Account IDs by comma
- For each account ID, calls Meta API:
  ```
  GET /{account_id}/insights?
    fields=campaign_name,spend,impressions,reach,clicks,actions,action_values,cpm,cpc,ctr,purchase_roas
    &date_preset=yesterday
    &level=campaign
    &access_token={META_TOKEN}
  ```
- Maps response fields to Daily Meta tab columns
- Appends rows to [Client Name] - Daily Meta (does not overwrite existing data)
- After all pulls complete, calls `clearYesterdayCheckboxes()`

### `clearYesterdayCheckboxes()`
- Removes all rows from CHECKBOX_STATE where Date < today's date
- Runs at end of daily pull

### Web App: `doPost(e)`
Accepts POST requests from internal dashboard. Verify `WEBAPP_SECRET` header on every request.

Supported actions:
```javascript
// action: save_config
// Payload: { clientSlug, offerName, products: [{name, position, currentPrice, active}] }
// → Write/update rows in [Client Name] - Config tab

// action: save_price_change
// Payload: { clientSlug, offerName, productName, newPrice, effectiveDate }
// → Append row to Price History tab

// action: save_checkbox_state
// Payload: { clientSlug, campaignName, action, checked, date }
// → If checked: append row to CHECKBOX_STATE
// → If unchecked: remove matching row from CHECKBOX_STATE

// action: setup_new_client
// Payload: { clientName, slug, metaAccountIds }
// → Add row to Clients tab
// → Create [Client Name] - Daily Meta tab
// → Create [Client Name] - Config tab (empty)

// action: setup_new_offer
// Payload: { clientSlug, offerName, products: [{name, position, price}] }
// → Create [Client Name] - [Offer Short Name] Purchases tab
// → Set header row based on products in position order
// → Add rows to Config tab
// → Add rows to Price History tab with today as effective date
```

### `setupClientTabs(clientName, slug)`
Creates the two per-client tabs with correct headers. Called by setup_new_client action.

### `setupOfferPurchasesTab(clientName, offerName, products)`
Creates the purchases tab for an offer. Column headers are product names in position order, followed by Total Units and Total Revenue. Total Revenue column uses the Price History VLOOKUP formula.

---

## Step 3 — Internal Dashboard (internal-dashboard.html)

Single HTML file. Reads from Google Sheet via Sheets API (published sheet + API key).

### Config block at top of script
```javascript
const SHEET_ID = 'YOUR_SHEET_ID_HERE';
const API_KEY = 'YOUR_API_KEY_HERE';
const WEBAPP_URL = 'YOUR_APPS_SCRIPT_WEBAPP_URL_HERE';

// Scaling thresholds
const SCALE_THRESHOLD = 1.4;
const ORANGE_THRESHOLD = 1.25;
const TURN_OFF_THRESHOLD = 0.75;
```

### Layout
```
[Top Bar: KSA logo + "Internal Ads Dashboard" | client name | sync status | refresh]
[Left Sidebar: client list] | [Main Content Area]
```

**Left sidebar:**
- Lists all active clients from Clients tab
- Click to select → loads that client's data
- Active client: `#FFDCD9` background, `#F77F9F` left border
- Below client list: "+ Add Client" button (opens settings panel)

**Top bar:**
- Background: `#F77F9F`
- KSA logo or "KSA" text mark (Bebas Neue, white)
- "Internal Ads Dashboard" subtitle
- Active client name
- Sync status (last pull time)
- Refresh button

### Main Content — Account Overview (default)

When a client is selected, default view shows account overview.

**Date range buttons:** 7-Day (default) | 30-Day | Custom (date picker)

**Summary cards row:**
- Total Ad Spend
- Total Revenue (from purchase logs, not Meta)
- Overall ROAS (Revenue / Spend)
- Total Profit (Revenue - Spend)

**Performance chart:**
- Line chart, day-by-day for selected range
- Toggle: Spend & Revenue | ROAS
- Reference lines at 1.0 (break-even) and 1.4 (scale threshold) on ROAS view

**Per-offer summary strip:**
- One card per offer (from Config tab)
- Shows: Offer name, ROAS (7-day default), Revenue, Profit
- Click → enters Offer View for that offer

### Main Content — Offer View

Activated when user clicks an offer card from the account overview.

**Back button:** "← Back to Account Overview"

**Offer header:** Offer name, Meta campaign filter (shows only campaigns matching offer)

**Purchase Summary Strip (top of offer view):**
- Shows last 7 days of purchase data from that offer's Purchases tab
- One cell per product in position order: Product Name | Units Sold | Conversion Rate
- Conversion rate = units sold / total tripwire units (or total leads if no tripwire)
- Products displayed in funnel order (by Position from Config)
- If products are reordered in Settings, this strip reorders automatically

**Campaign tabs:**
- Today's Actions (scale + turn off)
- Orange Flags
- New Campaigns (< 3 days data)
- All Campaigns

**Campaign rows:**
Each row shows: checkbox | status badge | campaign name | ROAS | spend | revenue | profit | CPM | action text
- Expand row to see day-by-day mini chart
- Checkbox state saves to CHECKBOX_STATE tab via Apps Script
- Checkbox restores on refresh for same day
- Checkboxes reset at 3am

### Status Logic (applied per campaign from Daily Meta tab)

Pull all rows for this client from Daily Meta tab. Group by campaign name.

```javascript
// For each campaign, calculate windows from daily data:
const yw   = windowMetrics(rows, 1);   // yesterday
const d7w  = windowMetrics(rows, 7);   // 7-day
const d14w = windowMetrics(rows, 14);  // 14-day
const d30w = windowMetrics(rows, 30);  // 30-day

function windowMetrics(rows, days) {
  // Filter to rows within last N days
  // Aggregate: sum spend, sum purchases, weighted avg ROAS
  // Return: { spend, purchases, roas, cpm, cpc, ctr }
}

// Status decisions (same rules as SCC dashboard):
// 1. New campaign (< 3 days data) → new_watch
// 2. 0 sales after Day 3 → turn_off
// 3. 7-day ROAS >= 1.4 AND yesterday >= 1.0 → scale_vertical (or duplicate if $50/$100 in name)
// 4. 7-day ROAS >= 1.4 BUT yesterday < 1.4 → hold
// 5. ROAS 1.25–1.39 → hold
// 6. ROAS 1.0–1.24, trending down from 14-day → orange_flag
// 7. ROAS 1.0–1.24, trending up or stable → hold
// 8. ROAS < 1.0 but yesterday >= 1.0 → orange_flag
// 9. ROAS < 0.75 AND yesterday < 1.0 → turn_off
// 10. Else → orange_flag
```

### Settings Panel

Slide-in panel from right. Sections:

**Client Settings:**
- Client name, slug (read-only after creation)
- Meta account IDs (editable, comma-separated for multiple)

**Offers:**
- List of offers for this client
- Each offer shows: offer name, products list in order
- "Edit Offer" → expands to show products with drag-to-reorder (updates Position)
- "Change Price" on any product → prompts for new price and effective date → writes to Price History
- "Add Offer" → creates new offer (name + initial products + prices)
- "Add Product to Offer" → adds column to that offer's Purchases tab

All saves POST to Apps Script web app endpoint.

---

## Step 4 — External Dashboard (external-dashboard.html)

Single HTML file. Reads same Google Sheet. Filters by `?client=slug` URL parameter.

### Config block
```javascript
const SHEET_ID = 'YOUR_SHEET_ID_HERE';
const API_KEY = 'YOUR_API_KEY_HERE';
// No WEBAPP_URL needed — external dashboard is read-only
```

### On load
- Read `client` URL param
- Look up client in Clients tab
- If not found → show "Dashboard not found" message
- If found → load that client's data only

### Layout

**Header:**
- KSA branding
- "Your Ads Dashboard"
- Client name
- "Powered by KS Agency"
- Last updated timestamp

**Date toggle:** 7-Day (default) | 30-Day

**Account Summary cards:**
- Total Ad Spend
- Revenue
- ROAS
- Total Profit

**Performance Chart:**
- Day-by-day ROAS line chart for selected date range
- Reference line at 1.0 (break-even)
- Toggle: Spend & Revenue | ROAS

**Per-Offer Cards:**
- One card per active offer
- Offer name
- Spend | Revenue | ROAS | Profit
- Mini ROAS sparkline (last 7 days)

**No flags, no action recommendations, no internal data, no other clients.**

**Footer:**
"Questions about your ads? Reach out to your KSA team."

---

## Step 5 — Deployment Instructions Docs

### apps-script/README.md
Step-by-step for Apps Script setup:
1. Open the Google Sheet
2. Extensions → Apps Script
3. Paste Code.gs content
4. Set Script Properties: META_TOKEN, SHEET_ID, WEBAPP_SECRET
5. Set timezone to Central Time (File → Project Settings → Time zone)
6. Create time-based trigger: runDailyPull, day timer, between 3am-4am
7. Deploy as Web App: Execute as Me, Who has access = Anyone
8. Copy the web app URL
9. Paste web app URL into internal-dashboard.html WEBAPP_URL constant
10. Re-deploy after any script changes (new deployment URL — update dashboard)

### github-pages-setup.md
Step-by-step:
1. Create GitHub repo (e.g. `ksa-ads-dashboard`)
2. Push all files
3. Settings → Pages → Source: Deploy from branch → main → / (root)
4. GitHub Pages URL will be: `https://[username].github.io/ksa-ads-dashboard/`
5. For custom subdomain: add CNAME file with subdomain (e.g. `ads.katschmoyer.com`)
6. In Kat's DNS: add CNAME record pointing `ads` → `[username].github.io`
7. In GitHub Pages settings: enter custom domain, enable HTTPS

### meta-token-setup.md
Step-by-step for creating a Meta System User token:
1. Go to business.facebook.com → Business Settings
2. Users → System Users → Add
3. Name it "KSA Ads Dashboard" with Employee role
4. Assign ad accounts: each client's ad account with Analyst permission
5. Generate Token → select ads_read permission
6. Copy token → paste into Apps Script Script Properties as META_TOKEN
7. Note: System User tokens don't expire — you only do this once

---

## Files to Create

```
/
├── brand-test.html              ← BUILD THIS FIRST
├── internal-dashboard.html
├── external-dashboard.html
├── fonts/
│   ├── MackinacPro-Bold.woff2   ← Rebecca to provide, or remove dir and use Lora
│   └── MackinacPro-Book.woff2
├── CNAME                        ← subdomain for GitHub Pages
├── apps-script/
│   ├── Code.gs
│   ├── README.md                ← deployment instructions
│   ├── github-pages-setup.md
│   └── meta-token-setup.md
└── CLAUDE.md                    ← project context (already provided)
```

---

## What NOT to Build

- No user authentication or login
- No database — sheet is the data store
- No React, Vue, or any framework — plain HTML/JS only
- No server-side code — all reads are client-side Sheets API calls
- No automated offer creation from the sheet side — settings panel in dashboard handles this
- No campaign progress tracker for external dashboard (cut from scope)
- No webhook integrations with Kajabi/Thrivecart (cut from scope)
- No auto-calculation of purchases from platform data (cut from scope)

---

## Known Client Details

**KSA (KS Agency):**
- Two Meta ad accounts — Rebecca to provide both IDs
- Offers: Rebecca to provide list

**HTP (Hope Taylor Photography):**
- Offers seen in spreadsheet: "What to Wear Tiny Offer", "New Presets"
- Products for What to Wear: Tripwire ($17), Bump ($11), Up-Sell 1 ($47), Upsell 2 ($97), Upsell 3 ($37), Upsell 4 ($5)

All other client details: Rebecca to provide Meta account IDs and confirm offer lists before first data pull.
