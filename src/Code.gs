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
