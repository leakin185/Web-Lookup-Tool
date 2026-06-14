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
                            'bing.com', 'foursquare.com', 'manta.com',
                            'yahoo.com', 'search.yahoo.com'];

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

// ── Search helpers ─────────────────────────────────────────────
function duckDuckGoSearch(query) {
  const url = 'https://html.duckduckgo.com/html/?q=' + encodeURIComponent(query);
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (res.getResponseCode() === 200) return res.getContentText();
  } catch (e) {
    Logger.log('duckDuckGoSearch error: ' + e.message);
  }
  return '';
}

// Extracts destination URLs from DuckDuckGo's uddg= redirect params.
function extractUrlsFromDDG(html) {
  const urls = [];
  const re = /uddg=([^&"'>\s]+)/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    try {
      const decoded = decodeURIComponent(m[1]);
      if (decoded.startsWith('http') && !urls.includes(decoded)) urls.push(decoded);
    } catch (e) {}
  }
  return urls;
}

function yahooSearch(query) {
  const url = 'https://search.yahoo.com/search?p=' + encodeURIComponent(query);
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (res.getResponseCode() === 200) return res.getContentText();
  } catch (e) {
    Logger.log('yahooSearch error: ' + e.message);
  }
  return '';
}

// Extracts destination URLs from Yahoo's /RU=<encoded-url>/ redirect params.
function extractUrlsFromYahoo(html) {
  const urls = [];
  const re = /\/RU=(https?[^/\s"']+)\//g;
  let m;
  while ((m = re.exec(html)) !== null) {
    try {
      const decoded = decodeURIComponent(m[1]);
      if (decoded.startsWith('http') && !urls.includes(decoded)) urls.push(decoded);
    } catch (e) {}
  }
  return urls;
}

// Searches DDG first; falls back to Yahoo if DDG is blocked (returns no URLs).
function searchUrls(query) {
  const ddgUrls = extractUrlsFromDDG(duckDuckGoSearch(query));
  if (ddgUrls.length > 0) return ddgUrls;
  Logger.log('DDG returned 0 URLs — falling back to Yahoo');
  return extractUrlsFromYahoo(yahooSearch(query));
}

// ── Website enrichment ─────────────────────────────────────────────
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

// ── Lookup orchestrator ────────────────────────────────────────
/**
 * Given a church name and country, returns:
 * { website, instagram, facebook, youtube, info }
 * Any field may be null if not found.
 */
function lookupChurch(name, country) {
  const result = { website: null, instagram: null, facebook: null, youtube: null, info: null };

  // Try exact-name search first; fall back to unquoted if no URLs come back.
  // searchUrls() tries DDG first, then Yahoo as backup.
  const countryStr = country ? ' ' + country : '';
  let urls = searchUrls('"' + name + '"' + countryStr + ' church');
  if (urls.length === 0) {
    urls = searchUrls(name + countryStr + ' church');
  }

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
    if (result.website && (result.instagram || result.facebook || result.youtube)) break;
  }

  if (result.website) enrichFromWebsite(result);

  return result;
}

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

  // Write only to empty cells (uses initial batch read to avoid extra API calls)
  let wroteAnything = false;

  function writeIfEmpty(col, value) {
    if (value && !String(values[col - 1] || '').trim()) {
      sheet.getRange(rowNum, col).setValue(value);
      wroteAnything = true;
    }
  }

  writeIfEmpty(COL.WEBSITE,   result.website);
  writeIfEmpty(COL.INSTAGRAM, result.instagram);
  writeIfEmpty(COL.FACEBOOK,  result.facebook);
  writeIfEmpty(COL.YOUTUBE,   result.youtube);
  writeIfEmpty(COL.INFO,      result.info);

  if (wroteAnything) {
    sheet.getRange(rowNum, COL.STATUS).setValue(STATUS_TO_VERIFY);
  } else {
    // Nothing found — write "Not found" to Info only if it's currently empty
    if (!String(values[COL.INFO - 1] || '').trim()) {
      sheet.getRange(rowNum, COL.INFO).setValue('Not found');
    }
  }

  return true;
}

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

  const startTime = Date.now();
  let count = 0;
  for (let row = headerRow + 1; row <= lastRow && count < BATCH_SIZE; row++) {
    const values     = sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
    const churchName = String(values[COL.CHURCH  - 1] || '').trim();
    const website    = String(values[COL.WEBSITE  - 1] || '').trim();
    const status     = String(values[COL.STATUS   - 1] || '').trim().toLowerCase();

    if (!churchName || website || status === STATUS_COMPLETE) continue;

    if (Date.now() - startTime > 270000) { // 4m30s safety margin
      ui.alert('⏱ Approaching time limit — processed ' + count + ' churches so far. Run again to continue.');
      return;
    }

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
