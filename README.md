# Claude Team Cost Report

Per-user cost reporting for a Claude Team (claude.ai) organization.

The Team plan bills two ways at once: **flat seat fees** (which never
appear in any analytics export) and **usage-credit spend** (which does).
This tool joins the two so you can answer *"how much are we actually
paying for each person?"* — something the admin dashboard doesn't show
directly, and something the Analytics API can't answer either, since
that API is Enterprise-only.

It takes the CSV exports you can download from the claude.ai analytics
pages, adds your seat prices, and produces:

- a console summary (totals, per-user table, top spend lines)
- an `.xlsx` workbook with live formulas, so seat prices can be changed
  later inside the spreadsheet and everything recalculates

```
Claude Team cost report — 2026-08-01 to 2026-08-31  [format: spend, roster joined]
Members: 14  Seats: {'Standard': 13, 'Unassigned': 1}
Seat fees: $390.00   Usage spend: $82.62   TOTAL: $472.62
```

---

## 1. Which reports to grab

Both exports come from the claude.ai **Analytics** area. You need to be
an **Owner or Primary Owner** of the Team organization to see it
(Enterprise Admins can see everything except Spend).

To get there: log in to claude.ai → click your **initials in the lower
left corner** → select **Analytics** (or go straight to
`https://claude.ai/analytics/overview`).

### Report A — Spend report (primary input, recommended)

Per-user, **per-product, per-model** token usage and spend. Supports
custom date ranges.

1. Open **Analytics → Overview** (`claude.ai/analytics/overview`).
2. Scroll down past "Who's using Claude?", "How are they using Claude?"
   and "What are the results?" to the section titled
   **"How much is Claude costing?"**.
   > The `30d` dropdowns near the top of the page only change the
   > on-screen charts — the export dialog has its own date picker.
3. Click **"Export spend report"**.
4. Pick a time period: **MTD**, **Last Month**, **Last 90 Days**, or
   **Custom**. With Custom you choose exact start/end dates — up to
   **90 days back**, and the most recent data available is
   **yesterday** (spend data refreshes daily with a one-day delay).
5. Click **Download**. You'll get a file named like:

   ```
   spend-report-<org-uuid>-2026-08-01-to-2026-08-31.csv
   ```

   Keep that filename — the script parses the date range from it.

Columns: `user_email`, `account_uuid`, `product` (Chat, Claude Code,
Cowork, Office Agents, Claude Design, ...), `model`, `total_requests`,
`total_prompt_tokens`, `total_completion_tokens`,
`total_net_spend_usd` (after discounts/credits — what you actually
spent), `total_gross_spend_usd` (before discounts), plus cache-token
detail.

**Important quirk:** the spend report only contains members who had
usage in the selected window. Someone occupying a paid seat who never
opened Claude won't appear — which is exactly why Report B exists.

### Report B — Members analytics export (the roster)

One row per member with role, seat tier, activity counts, and a
rolled-up "Estimated Spend (USD)".

1. Open **Analytics → Overview**.
2. In the **"Who's using Claude?"** section, find the **Members**
   panel and click **"See all"** (or use the members view under the
   Chat analytics pages).
3. Export the members list. You'll get a file named like:

   ```
   members-analytics-<org-uuid>-2026-08-15-to-2026-09-13.csv
   ```

Columns include: `Name`, `Email`, `Role`, `Seat Tier`, `Last Active`,
`Days Active`, `Chats`, `Messages`, per-product activity counts, and
`Estimated Spend (USD)`.

### Which do I feed the script?

| You have | Command | You get |
|---|---|---|
| Spend report **+** members export | `report.py spend.csv --roster members.csv` | **Best:** full roster incl. zero-usage seats, real seat tiers, per-product/model detail |
| Spend report only | `report.py spend.csv` | Per-user + detail, but only users with usage; all seats assumed Standard |
| Members export only | `report.py members.csv` | Full roster + rolled-up spend, no product/model detail |

The format is auto-detected from the columns — no flag needed.

---

## 2. Installation

Python 3.9+.

```bash
git clone <this-repo>
cd claude-team-cost-report
python3 -m venv .venv && source .venv/bin/activate   # optional
pip install -r requirements.txt                      # pandas, openpyxl
```

---

## 3. Running it

Recommended monthly run (spend report + roster):

```bash
python3 claude_team_cost_report.py \
    spend-report-<org>-2026-08-01-to-2026-08-31.csv \
    --roster members-analytics-<org>-2026-08-15-to-2026-09-13.csv
```

Members export only:

```bash
python3 claude_team_cost_report.py members-analytics-<org>-....csv
```

### Options

| Flag | Default | Meaning |
|---|---|---|
| `--roster FILE` | — | Members-analytics CSV used to supply the full member list and real seat tiers when the main input is a spend report |
| `--seat-price N` | `30.00` | Price per **Standard** seat for the period. Team Standard is $30/member/month billed monthly, $25/member/month billed annually — pass `--seat-price 25` if you're on annual billing |
| `--premium-price N` | `150.00` | Price per **Premium** seat, if any appear in your roster |
| `--out FILE` | derived from input filename | Output `.xlsx` path |
| `--no-xlsx` | off | Console summary only, skip the workbook |
| `--console-usage` | off | Also fetch `platform.claude.com` (Console/API) cost data via the Admin API and add it as its own sheet — see [§6](#6-consoleapi-usage-optional---console-usage) |
| `--console-start DATE` | same period as the input filename | Start date (`YYYY-MM-DD`) for the Console/API pull |
| `--console-end DATE` | same period as the input filename | End date (`YYYY-MM-DD`) for the Console/API pull |

Seat pricing is an input **you** provide — no export contains it, and
Anthropic's prices can change. Verify current pricing at
<https://support.claude.com> before trusting the defaults. `Unassigned`
seat tiers are costed at $0.

### Examples

```bash
# Annual billing prices
python3 claude_team_cost_report.py spend.csv --roster members.csv --seat-price 25

# Just the numbers, no spreadsheet
python3 claude_team_cost_report.py spend.csv --roster members.csv --no-xlsx

# Custom output location
python3 claude_team_cost_report.py spend.csv --roster members.csv \
    --out reports/2026-08.xlsx
```

---

## 4. What you get

### Console

- Period (parsed from the filename), detected format, member count,
  seat-tier breakdown
- Totals: seat fees, usage spend, grand total
- Per-user table sorted by total cost
- Top 5 spend lines by user/product/model (spend-report input only)
- A reminder if you ran a spend report without `--roster`

### Workbook (`claude-team-per-user-cost_<period>.xlsx`)

| Sheet | Contents |
|---|---|
| **Assumptions** | Seat prices in editable yellow input cells. Change `B5` (Standard) or `B6` (Premium) and every cost formula recalculates — no need to rerun the script for a price change |
| **Per-User Cost** | One row per member: seat tier, requests, usage spend, seat fee (formula driven by Assumptions), total cost, % of total, with a totals row |
| **By Product & Model** | (spend-report input only) The full breakdown: requests, prompt/completion tokens, net and gross spend per user/product/model |
| **Console API Usage** | (only with `--console-usage`) Workspace / model / cost-type breakdown of `platform.claude.com` API spend, with a totals row. Kept separate from Per-User Cost — see [§6](#6-consoleapi-usage-optional---console-usage) |

Seat fees, totals, and percentages are live formulas, not hardcoded
values. Open the file once in Excel or LibreOffice so the formulas
calculate before reading values programmatically.

---

## 5. How costs are computed

```
per-user total = seat fee (by seat tier, from your inputs)
              + usage spend (net spend from the export, after discounts/credits)
```

- Usage spend uses `total_net_spend_usd` (spend report) or
  `Estimated Spend (USD)` (members export). Gross spend is carried
  into the detail sheet for reference.
- With `--roster`, users are joined on lowercased email
  (outer join): roster members with no usage get $0 usage spend;
  spend-report users missing from the roster are kept and assumed
  Standard.

---

## 6. Console/API usage (optional, `--console-usage`)

Team seat fees and Cowork/Chat usage spend are one billing relationship.
If your org also has developers calling the API directly through
**Claude Console** (`platform.claude.com`) — pay-as-you-go, billed by
API key/workspace — that's a **separate** billing relationship, not
covered by the Team analytics exports at all. `--console-usage` fetches
that data too, via the Console [Admin
API](https://platform.claude.com/docs/en/manage-claude/usage-cost-api),
so one workbook can show the complete picture.

### Setup

1. In [Claude Console](https://platform.claude.com) → **Settings →
   Admin API keys**, create an Admin API key (starts with
   `sk-ant-admin01-...`). You need to be an org admin.
2. Export it as an environment variable — never pass it on the command
   line (it'll end up in your shell history):

   ```bash
   export ANTHROPIC_ADMIN_KEY=sk-ant-admin01-...
   ```

### Running it

```bash
python3 claude_team_cost_report.py \
    spend-report-<org>-2026-08-01-to-2026-08-31.csv \
    --roster members-analytics-<org>-2026-08-15-to-2026-09-13.csv \
    --console-usage
```

By default it pulls the same date range as the Team CSV's filename.
Pass `--console-start`/`--console-end` to use a different window.

### What you get

A **Console API Usage** sheet: spend broken out by workspace, model,
cost type (tokens / web search / code execution), and service tier,
with a totals row — plus the same breakdown's top lines and grand
total printed to the console.

**This total is *not* added into Per-User Cost or its TOTAL row.**
Console/API usage is billed by API key and workspace, not by named
Team member, so there's no reliable way to attribute it to a person —
attempting to fold it into the per-user total would silently misstate
individual costs. The two totals are shown side by side (console
output and the Assumptions sheet notes) so you can see the org's full
Claude spend, without pretending they're the same kind of number.

---

## 7. Caveats & gotchas

- **90-day lookback ceiling.** The spend export can't reach further
  back than 90 days, so run this monthly and archive the CSVs if you
  want history.
- **One-day data delay.** Don't export for a period ending today;
  yesterday is the freshest complete day.
- **Estimates, not invoices.** Export spend should closely match your
  invoice, but the billing dashboard is authoritative.
- **Date ranges don't need to match.** The roster is only used for
  the member list and seat tiers, so a roster from a slightly
  different window is fine — just make sure it's recent enough to
  reflect current seats.
- **No API on Team.** The Analytics API requires the Enterprise plan;
  on Team, the CSV export is the only way to get this data out, so
  the download step stays manual. (`--console-usage` is a *different*
  API for a *different* product — see [§6](#6-consoleapi-usage-optional---console-usage) — this doesn't contradict the point above.)
- **`--console-usage` needs network access and an Admin API key.**
  Everything else in this tool is local file I/O; this one flag calls
  `api.anthropic.com`. If `ANTHROPIC_ADMIN_KEY` isn't set, or the key
  isn't an Admin key, it fails with a clear error before writing
  anything.
- **Privacy.** The exports contain employee emails and usage data.
  The included `.gitignore` excludes `*.csv` and `*.xlsx` so you don't
  accidentally commit them.

---

## 8. Repo layout

```
claude_team_cost_report.py   the tool (stdlib + pandas + openpyxl)
requirements.txt             pandas, openpyxl
README.md                    this file
.gitignore                   keeps exports and workbooks out of git
```
