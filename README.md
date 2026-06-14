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
