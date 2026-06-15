# Church Lookup Tool — Design Spec

**Date:** 2026-06-14  
**Project:** Global Conference Church Research  
**Author:** PM (leakin185)

---

## Context

The church research team is preparing for an upcoming global conference. They have ~200 churches to research across a shared Google Sheet. Researchers need to fill in website, Instagram, Facebook, YouTube, and a short description for each church — currently done manually, which is slow and creates a 25% completion bottleneck.

This tool automates the lookup so researchers can trigger it themselves with one click, without any technical setup.

---

## Target Sheet

- **Sheet URL:** https://docs.google.com/spreadsheets/d/1iy8pGKyZ23G01TsfZgB7IkWhVOuHbhcmvP4OnbQB4MQ
- **Tab name:** `Churches`
- **Scale:** ~200 churches, 50–500 range

### Column Map

| Col | Field | Role |
|-----|-------|------|
| A | Assigned To | Read-only (do not touch) |
| B | New/Returning | Read-only |
| C | Attended GIC | Read-only |
| D | Country | **Input** — used in search query |
| E | Church | **Input** — primary search term |
| F | Links* | Read-only (manual notes) |
| G | Website | **Output** |
| H | Instagram | **Output** |
| I | Facebook | **Output** |
| J | YouTube | **Output** |
| K | Info | **Output** — short description from meta tag |
| L | Team No. | Read-only |
| M | No. of Delegates | Read-only |
| N | Freshsales ID | Read-only |
| O | Status | Read-only (script checks this; never writes it) |

**Rule:** The script only writes to cols G–K. It never overwrites a cell that already has a value, so manual edits are safe.

---

## Architecture

A single Google Apps Script file bound to the shared sheet. No external server, no API keys, no cost.

### Menu (added via `onOpen`)

```
Church Lookup
├── Lookup This Row
└── Lookup All Incomplete (batch of 20)
```

Appears automatically for every user who opens the sheet.

### Functions

| Function | Trigger | Behaviour |
|----------|---------|-----------|
| `onOpen()` | Sheet open | Adds the custom menu |
| `lookupSelectedRow()` | Menu click | Looks up the church on the active row |
| `lookupAllIncomplete()` | Menu click | Finds up to 20 rows where Website is blank and Status ≠ "Complete", processes them, alerts "Done — processed N churches. Run again for more." |
| `processRow(sheet, rowNum)` | Internal | Reads Church + Country, calls `lookupChurch()`, writes results to cols G–K |
| `lookupChurch(name, country)` | Internal | Runs web searches, returns `{website, instagram, facebook, youtube, info}` |
| `duckDuckGoSearch(query)` | Internal | Fetches DuckDuckGo HTML search results page (no API key) |
| `extractUrlsFromHtml(html)` | Internal | Parses `<a>` tags from search result HTML, returns array of URLs |
| `enrichFromWebsite(result)` | Internal | Visits the church website, scrapes `<meta>` description and social links from footer/nav |

### Why DuckDuckGo / Yahoo Instead of Google

| Option | Why not |
|--------|---------|
| **Google Custom Search API** | Free tier limited to 100 queries/day, then $5 per 1,000. Requires a Google Cloud project, API key, and Custom Search Engine ID — too much setup for a tool meant to be pasted into a sheet by a non-technical PM. |
| **Google HTML scraping** | Google aggressively blocks automated requests with CAPTCHAs and IP bans. Even `UrlFetchApp` requests from Google's own Apps Script infrastructure get blocked after a few queries. |
| **DuckDuckGo HTML scraping** | Provides a dedicated lightweight endpoint (`html.duckduckgo.com/html/`) with minimal bot detection. Free, no API key, works reliably with `UrlFetchApp`. |
| **Yahoo HTML scraping** | Also lenient on bot detection. Used as a fallback when DuckDuckGo is temporarily unavailable. |

None of these are official APIs — the tool scrapes server-rendered HTML search pages directly. This means results could break if the search engine changes its HTML structure, but for a one-time ~200-church research task the tradeoff is acceptable.

### Search Strategy (free, no API)

1. **Query:** `"Church Name" Country church` → fetch DuckDuckGo HTML results
2. **Categorise URLs** from results:
   - `instagram.com/*` → Instagram
   - `facebook.com/*` → Facebook
   - `youtube.com/*` → YouTube
   - Everything else (not a known social/aggregator domain) → Website candidate
3. **Visit the website** (if found) → scrape `<meta name="description">` for Info, and scan page HTML for any social links not yet found
4. **Fallback:** If nothing is found, write `"Not found"` to col K (Info) so researchers know the row was attempted

### Batch & Timeout Handling

- Apps Script hard limit: 6 minutes per execution
- Batch size: 20 rows max per run (~10–15 seconds per church = ~3–5 min total, safely under limit)
- Researchers re-trigger for the next batch until all done

### Skip Logic

Skip a row if any of these are true:
- Col E (Church) is empty
- Col G (Website) already has a value
- Col O (Status) = "Complete"

---

## Limitations

- **Social links:** ~80% accuracy — social media pages sometimes don't rank in search results
- **Instagram/Facebook content:** Meta blocks page scraping; links come from search results only, not live profile data
- **Follower counts:** Not attempted — blocked without authenticated API access (which costs money)
- **JS-heavy sites:** Apps Script cannot execute JavaScript, so single-page apps may return empty descriptions

---

## One-Time Setup (PM does this once)

1. Open the Google Sheet
2. **Extensions → Apps Script**
3. Paste the script, save
4. Click **Run → `onOpen`** once to grant authorization
5. Reload the sheet — the "Church Lookup" menu appears for everyone

---

## Verification

1. Open the sheet, confirm the "Church Lookup" menu appears
2. Click a row with a blank Website → **Lookup This Row** → verify cols G–K fill within 10 seconds
3. Check a known church (e.g. "Alfa Omega Church", Indonesia) to confirm the website URL matches `alfaomegachurch.com`
4. Run **Lookup All Incomplete** → confirm alert shows count, re-run confirms it skips already-filled rows
5. Manually edit one cell → re-run → confirm the script does not overwrite it
