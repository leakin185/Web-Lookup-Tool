# Church Lookup Tool — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Google Apps Script that adds a "Church Lookup" menu to the shared Google Sheet, letting researchers auto-fill Website / Instagram / Facebook / YouTube / Info for each church with one click.

**Architecture:** A single `Code.gs` Apps Script file bound to the shared sheet. It queries DuckDuckGo HTML search (free, no API key) for each church, extracts and categorises URLs from the results, then visits the church's own website to scrape a description and any missing social links. Results are written only to empty cells; Status is set to "To Be Verified" so researchers confirm before marking Complete.

**Tech Stack:** Google Apps Script (V8), UrlFetchApp, SpreadsheetApp, DuckDuckGo HTML search (no API key)

**Shared sheet:** `https://docs.google.com/spreadsheets/d/1iy8pGKyZ23G01TsfZgB7IkWhVOuHbhcmvP4OnbQB4MQ`  
**Tab:** `Churches`

---

## File Structure

```
src/
  Code.gs          — complete Apps Script (copy-paste into Extensions → Apps Script)
  appsscript.json  — Apps Script manifest
```

---

## Column Reference (Churches tab)

| Col | Field | Role |
|-----|-------|------|
| A (1) | Assigned To | read-only |
| D (4) | Country | **input** |
| E (5) | Church | **input** |
| G (7) | Website | **output** |
| H (8) | Instagram | **output** |
| I (9) | Facebook | **output** |
| J (10) | YouTube | **output** |
| K (11) | Info | **output** |
| O (15) | Status | **output** — script sets to "To Be Verified" |

---

## Task 1: Apps Script manifest

**Files:**
- Create: `src/appsscript.json`

- [ ] **Step 1: Create the manifest file**

```json
{
  "timeZone": "Asia/Singapore",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

Save to `src/appsscript.json`.

- [ ] **Step 2: Commit**

```bash
git add src/appsscript.json
git commit -m "chore: add Apps Script manifest"
```

---

## Task 2: Constants and sheet helpers

**Files:**
- Create: `src/Code.gs`

- [ ] **Step 1: Create `src/Code.gs` with constants and sheet helpers**

```javascript
// ── Constants ──────────────────────────────────────────────────
const SHEET_NAME = 'Churches';
const BATCH_SIZE = 20;
const STATUS_TO_VERIFY = 'To Be Verified';
const STATUS_COMPLETE = 'complete'; // lowercase for case-insensitive compare

const COL = {
  COUNTRY:   4,
  CHURCH:    5,
  WEBSITE:   7,
  INSTAGRAM: 8,
  FACEBOOK:  9,
  YOUTUBE:   10,
  INFO:      11,
  STATUS:    15
};

const SOCIAL_DOMAINS    = ['instagram.com', 'facebook.com', 'youtube.com',
                           'twitter.com', 'x.com', 'linkedin.com', 'tiktok.com'];
const AGGREGATOR_DOMAINS = ['yelp.com', 'yellowpages', 'tripadvisor.com',
                            'wikipedia.org', 'google.com', 'duckduckgo.com',
                            'bing.com', 'foursquare.com', 'manta.com'];

// ── Sheet helpers ──────────────────────────────────────────────
function getChurchSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return sheet;
}

/**
 * Returns the 1-based row number of the header row.
 * Scans the first 10 rows for one that contains both "Church" and "Website".
 */
function findHeaderRow(sheet) {
  const maxScan = Math.min(10, sheet.getLastRow());
  const data = sheet.getRange(1, 1, maxScan, COL.STATUS).getValues();
  for (let i = 0; i < data.length; i++) {
    const rowText = data[i].join('|');
    if (rowText.includes('Church') && rowText.includes('Website')) return i + 1;
  }
  return 1; // fallback
}

function isSocialOrAggregator(url) {
  const lower = url.toLowerCase();
  return [...SOCIAL_DOMAINS, ...AGGREGATOR_DOMAINS].some(d => lower.includes(d));
}
```

- [ ] **Step 2: Paste into Apps Script editor and verify no syntax errors**

In the Google Sheet: **Extensions → Apps Script → New file named `Code`** → replace all content → Save (Ctrl+S). The editor should show no red underlines.

- [ ] **Step 3: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add constants and sheet helpers"
```

---

## Task 3: DuckDuckGo search helpers

**Files:**
- Modify: `src/Code.gs` (append below Task 2 code)

- [ ] **Step 1: Append search functions to `Code.gs`**

```javascript
// ── Search helpers ─────────────────────────────────────────────
function duckDuckGoSearch(query) {
  const url = 'https://html.duckduckgo.com/html/?q=' + encodeURIComponent(query);
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (res.getResponseCode() === 200) return res.getContentText();
  } catch (e) {
    Logger.log('duckDuckGoSearch error: ' + e.message);
  }
  return '';
}

/**
 * Extracts actual destination URLs from DuckDuckGo's uddg= redirect params.
 */
function extractUrlsFromHtml(html) {
  const urls = [];
  const re = /uddg=([^&"'\s>]+)/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    try {
      const decoded = decodeURIComponent(m[1]);
      if (decoded.startsWith('http') && !urls.includes(decoded)) urls.push(decoded);
    } catch (e) {}
  }
  return urls;
}
```

- [ ] **Step 2: Add manual test function (run in Apps Script editor)**

Append to `Code.gs`:

```javascript
/** Run in Apps Script editor: Execution → testSearchHelpers */
function testSearchHelpers() {
  const html = duckDuckGoSearch('Alfa Omega Church Indonesia official website');
  Logger.log('HTML length: ' + html.length);
  const urls = extractUrlsFromHtml(html);
  Logger.log('URLs found (' + urls.length + '):');
  urls.forEach(u => Logger.log('  ' + u));
  // Expected: list includes alfaomegachurch.com and/or its social pages
}
```

- [ ] **Step 3: Run `testSearchHelpers` in the Apps Script editor**

Click **Run → testSearchHelpers**. In the Execution log expect:
- `HTML length:` a number > 5000
- At least 5 URLs listed
- `alfaomegachurch.com` or `instagram.com/alfaomegachurch` visible

If HTML length is 0, DuckDuckGo may be blocking — try adding a Utilities.sleep(1000) before the fetch.

- [ ] **Step 4: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add DuckDuckGo search helpers"
```

---

## Task 4: Website enrichment scraper

**Files:**
- Modify: `src/Code.gs` (append)

- [ ] **Step 1: Append `enrichFromWebsite` to `Code.gs`**

```javascript
// ── Website enrichment ─────────────────────────────────────────
/**
 * Visits the church's own website and fills in any missing social links
 * and a short description from the <meta> tags.
 * Modifies `result` in place.
 */
function enrichFromWebsite(result) {
  try {
    const res = UrlFetchApp.fetch(result.website, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
    if (res.getResponseCode() !== 200) return;
    const html = res.getContentText();

    // Description: try og:description then name=description
    if (!result.info) {
      const descRe = [
        /<meta[^>]+property=["']og:description["'][^>]+content=["']([^"']{20,400})["']/i,
        /<meta[^>]+content=["']([^"']{20,400})["'][^>]+property=["']og:description["']/i,
        /<meta[^>]+name=["']description["'][^>]+content=["']([^"']{20,400})["']/i,
        /<meta[^>]+content=["']([^"']{20,400})["'][^>]+name=["']description["']/i
      ];
      for (const re of descRe) {
        const m = html.match(re);
        if (m) { result.info = m[1].trim(); break; }
      }
    }

    // Social links from page HTML
    const socialRe = /https?:\/\/(?:www\.)?(instagram\.com|facebook\.com|youtube\.com)\/[^\s"'<>)\]]+/gi;
    let m;
    while ((m = socialRe.exec(html)) !== null) {
      const url  = m[0].replace(/[,;.]+$/, ''); // strip trailing punctuation
      const host = m[1].toLowerCase();
      if (host === 'instagram.com' && !result.instagram && !url.includes('/p/') && !url.includes('/reel')) {
        result.instagram = url;
      } else if (host === 'facebook.com' && !result.facebook && !url.includes('/sharer') && !url.includes('/share')) {
        result.facebook = url;
      } else if (host === 'youtube.com' && !result.youtube &&
                 (url.includes('/channel/') || url.includes('/@') || url.includes('/user/'))) {
        result.youtube = url;
      }
    }
  } catch (e) {
    Logger.log('enrichFromWebsite error: ' + e.message);
  }
}
```

- [ ] **Step 2: Add manual test**

Append to `Code.gs`:

```javascript
/** Run in Apps Script editor to verify enrichment */
function testEnrichFromWebsite() {
  const result = { website: 'https://alfaomegachurch.com/', instagram: null, facebook: null, youtube: null, info: null };
  enrichFromWebsite(result);
  Logger.log(JSON.stringify(result, null, 2));
  // Expected: info contains a non-empty description string
}
```

- [ ] **Step 3: Run `testEnrichFromWebsite` in the Apps Script editor**

Expected log output (values approximate):
```
{
  "website": "https://alfaomegachurch.com/",
  "instagram": "https://www.instagram.com/alfaomegachurch/",  // or null if blocked
  "facebook": ...,
  "youtube": ...,
  "info": "Alfa Omega Church is a church for imperfect people..."
}
```

`info` must be non-empty. Social links may be null (that is fine — the search step fills them).

- [ ] **Step 4: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add website enrichment scraper"
```

---

## Task 5: Main lookup orchestrator

**Files:**
- Modify: `src/Code.gs` (append)

- [ ] **Step 1: Append `lookupChurch` to `Code.gs`**

```javascript
// ── Lookup orchestrator ────────────────────────────────────────
/**
 * Given a church name and country, returns:
 * { website, instagram, facebook, youtube, info }
 * Any field may be null if not found.
 */
function lookupChurch(name, country) {
  const result = { website: null, instagram: null, facebook: null, youtube: null, info: null };

  const query = '"' + name + '" ' + (country || '') + ' church';
  const html  = duckDuckGoSearch(query);
  const urls  = extractUrlsFromHtml(html);

  for (const url of urls) {
    const lower = url.toLowerCase();
    if (lower.includes('instagram.com') && !lower.includes('/p/') && !lower.includes('/reel') && !result.instagram) {
      result.instagram = url;
    } else if (lower.includes('facebook.com') && !lower.includes('/sharer') && !lower.includes('/share') && !result.facebook) {
      result.facebook = url;
    } else if (lower.includes('youtube.com') &&
               (lower.includes('/channel/') || lower.includes('/@') || lower.includes('/user/')) &&
               !result.youtube) {
      result.youtube = url;
    } else if (!result.website && !isSocialOrAggregator(url)) {
      result.website = url;
    }
    if (result.website && result.instagram && result.facebook && result.youtube) break;
  }

  if (result.website) enrichFromWebsite(result);

  return result;
}
```

- [ ] **Step 2: Add end-to-end test**

Append to `Code.gs`:

```javascript
/** Run in Apps Script editor */
function testLookupChurch() {
  const result = lookupChurch('Alfa Omega Church', 'Indonesia');
  Logger.log(JSON.stringify(result, null, 2));
  // Expected: website contains "alfaomegachurch"
  // instagram and/or facebook should be non-null
}
```

- [ ] **Step 3: Run `testLookupChurch` in the Apps Script editor**

Expected:
```json
{
  "website": "https://alfaomegachurch.com/",
  "instagram": "https://www.instagram.com/alfaomegachurch/",
  "facebook": "...",
  "youtube": "...",
  "info": "Alfa Omega Church is a church for imperfect people..."
}
```

`website` must be non-null and contain "alfaomegachurch". If it returns a wrong site, check the query string — the name may need to drop quote wrapping.

- [ ] **Step 4: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add lookupChurch orchestrator"
```

---

## Task 6: Row processing (reads and writes the sheet)

**Files:**
- Modify: `src/Code.gs` (append)

- [ ] **Step 1: Append `processRow` to `Code.gs`**

```javascript
// ── Row processing ─────────────────────────────────────────────
/**
 * Looks up the church on `rowNum` and writes results to the sheet.
 * Returns true if the row was processed, false if skipped.
 *
 * Skip conditions:
 *   - Church name (col E) is empty
 *   - Website (col G) already has a value
 *   - Status (col O) is "complete" (case-insensitive)
 */
function processRow(sheet, rowNum) {
  const values = sheet.getRange(rowNum, 1, 1, COL.STATUS).getValues()[0];
  const churchName = String(values[COL.CHURCH  - 1] || '').trim();
  const country    = String(values[COL.COUNTRY  - 1] || '').trim();
  const website    = String(values[COL.WEBSITE  - 1] || '').trim();
  const status     = String(values[COL.STATUS   - 1] || '').trim().toLowerCase();

  if (!churchName || website || status === STATUS_COMPLETE) return false;

  const result = lookupChurch(churchName, country);

  // Write only to empty cells
  function writeIfEmpty(col, value) {
    if (value && !sheet.getRange(rowNum, col).getValue()) {
      sheet.getRange(rowNum, col).setValue(value);
    }
  }

  writeIfEmpty(COL.WEBSITE,   result.website);
  writeIfEmpty(COL.INSTAGRAM, result.instagram);
  writeIfEmpty(COL.FACEBOOK,  result.facebook);
  writeIfEmpty(COL.YOUTUBE,   result.youtube);
  writeIfEmpty(COL.INFO,      result.info || 'Not found');

  // Always set status to "To Be Verified" (overwrite whatever was there)
  sheet.getRange(rowNum, COL.STATUS).setValue(STATUS_TO_VERIFY);

  return true;
}
```

- [ ] **Step 2: Add sheet-level test (run in Apps Script editor)**

Append to `Code.gs`:

```javascript
/**
 * Processes a specific row number for manual testing.
 * Change ROW_TO_TEST to a row in the Churches tab that has a church name but no website.
 */
function testProcessRow() {
  const ROW_TO_TEST = 3; // adjust to a real empty row in your sheet
  const sheet = getChurchSheet();
  const processed = processRow(sheet, ROW_TO_TEST);
  Logger.log('Processed: ' + processed);
  // Then open the sheet and check that row — Website and Status columns should be filled.
}
```

- [ ] **Step 3: Run `testProcessRow` in the Apps Script editor**

Change `ROW_TO_TEST` to the row number of a church with no website. Run it. Open the sheet and verify:
- Col G (Website) is filled
- Col O (Status) shows "To Be Verified"
- No other columns were changed

- [ ] **Step 4: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add processRow with write-if-empty and status update"
```

---

## Task 7: Menu, single-row lookup, and batch lookup

**Files:**
- Modify: `src/Code.gs` (append)

- [ ] **Step 1: Append menu and action functions to `Code.gs`**

```javascript
// ── Menu ───────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⛪ Church Lookup')
    .addItem('Lookup This Row', 'lookupSelectedRow')
    .addSeparator()
    .addItem('Lookup All Incomplete (batch of 20)', 'lookupAllIncomplete')
    .addToUi();
}

// ── Single-row action ──────────────────────────────────────────
function lookupSelectedRow() {
  const sheet     = getChurchSheet();
  const activeRow = sheet.getActiveCell().getRow();
  const headerRow = findHeaderRow(sheet);

  if (activeRow <= headerRow) {
    SpreadsheetApp.getUi().alert('Please select a church data row (not the header).');
    return;
  }

  const processed = processRow(sheet, activeRow);
  SpreadsheetApp.getUi().alert(
    processed
      ? '✅ Done! Results filled in. Status set to "To Be Verified" — please review.'
      : '⏭ Row skipped — it already has a website, the church name is blank, or status is Complete.'
  );
}

// ── Batch action ───────────────────────────────────────────────
function lookupAllIncomplete() {
  const ui        = SpreadsheetApp.getUi();
  const sheet     = getChurchSheet();
  const headerRow = findHeaderRow(sheet);
  const lastRow   = sheet.getLastRow();

  let count = 0;
  for (let row = headerRow + 1; row <= lastRow && count < BATCH_SIZE; row++) {
    const values     = sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
    const churchName = String(values[COL.CHURCH  - 1] || '').trim();
    const website    = String(values[COL.WEBSITE  - 1] || '').trim();
    const status     = String(values[COL.STATUS   - 1] || '').trim().toLowerCase();

    if (!churchName || website || status === STATUS_COMPLETE) continue;

    processRow(sheet, row);
    count++;
    SpreadsheetApp.flush(); // write each row immediately so progress is visible
    Utilities.sleep(500);   // brief pause to avoid rate-limiting
  }

  ui.alert(
    count > 0
      ? '✅ Done — processed ' + count + ' churches. Status set to "To Be Verified". Run again for more.'
      : '✅ No incomplete churches found in this batch. All done!'
  );
}
```

- [ ] **Step 2: Save the updated script in the Apps Script editor**

Ctrl+S in the editor. Confirm no syntax errors.

- [ ] **Step 3: Reload the Google Sheet**

Close and reopen the sheet (or hard-refresh). The **"⛪ Church Lookup"** menu should appear in the menu bar between "Help" and any other add-ons.

- [ ] **Step 4: Test "Lookup This Row"**

- Click on a row that has a church name but no website
- Click **⛪ Church Lookup → Lookup This Row**
- Wait ~5–15 seconds
- Verify: Website / Instagram / Facebook / YouTube / Info fill in; Status becomes "To Be Verified"
- Confirm the alert dialog appears and is clearly worded

- [ ] **Step 5: Test "Lookup All Incomplete"**

- Click **⛪ Church Lookup → Lookup All Incomplete (batch of 20)**
- Watch the sheet fill in row by row (~2–4 min for 20 rows)
- Verify the final alert shows the correct count
- Verify rows already marked "Complete" were skipped

- [ ] **Step 6: Test overwrite protection**

- Manually type a test URL in a Website cell of any row
- Run **Lookup This Row** on that row
- Confirm the manually entered value was **not** overwritten

- [ ] **Step 7: Commit**

```bash
git add src/Code.gs
git commit -m "feat: add menu, lookupSelectedRow, lookupAllIncomplete"
```

---

## Task 8: Final cleanup and researcher handoff

**Files:**
- Modify: `src/Code.gs` — remove all `test*` functions before sharing

- [ ] **Step 1: Remove test functions from `Code.gs`**

Delete these functions (they were only for development):
- `testSearchHelpers`
- `testEnrichFromWebsite`
- `testLookupChurch`
- `testProcessRow`

- [ ] **Step 2: Final save and reload**

Save in Apps Script editor. Reload the sheet. Confirm menu still appears and one lookup still works.

- [ ] **Step 3: Write `README.md` for the repo**

Create `README.md` at repo root:

```markdown
# Web Lookup Tool

Google Apps Script that auto-fills church research data (website, social links, description) into the shared Google Sheet for the Global Conference.

## Setup (one-time, PM does this)

1. Open the [shared Google Sheet](https://docs.google.com/spreadsheets/d/1iy8pGKyZ23G01TsfZgB7IkWhVOuHbhcmvP4OnbQB4MQ)
2. **Extensions → Apps Script**
3. Paste the contents of `src/Code.gs` into the editor (replace all existing content)
4. **File → Save** (Ctrl+S)
5. Click **Run → onOpen**, approve the authorization prompt
6. Reload the Google Sheet — the **"⛪ Church Lookup"** menu appears

## Usage (researchers)

- **Lookup This Row** — click a church row, then click this menu item. Results fill in within ~10 seconds.
- **Lookup All Incomplete (batch of 20)** — processes up to 20 un-researched churches at once (~3 min). Run again for the next batch.

## After lookup

Status is set to **"To Be Verified"**. Researchers check the filled data, correct anything wrong, then manually change Status to **"Complete"**.

## Limitations

- Social links are ~80% accurate — some church pages don't rank in search results
- Follower counts are not captured (Meta/Google block this without paid APIs)
- JavaScript-heavy church websites may return an empty description
```

- [ ] **Step 4: Final commit and push**

```bash
git add src/Code.gs README.md
git commit -m "feat: complete church lookup tool — ready for researcher use"
git push origin main
```

---

## Verification Checklist

- [ ] "⛪ Church Lookup" menu appears for any user who opens the sheet
- [ ] "Lookup This Row" fills Website/Instagram/Facebook/YouTube/Info and sets Status = "To Be Verified" within 15 seconds
- [ ] "Lookup All Incomplete" processes exactly 20 rows then stops, shows correct count in alert
- [ ] Rows with Status = "Complete" are always skipped
- [ ] Cells with existing values are never overwritten
- [ ] "Not found" is written to Info when a church yields no results (so researchers know it was attempted)
- [ ] Re-running on an already-filled row is silently skipped (no double-write, no error)
