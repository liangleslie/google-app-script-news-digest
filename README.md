# Google Sheets Cloud News Scraper

A serverless, container-bound Google Apps Script news pipeline running on Google Sheets. It scrapes news articles across modular site scrapers, summarises and clusters cross-agency stories using the Google AI Studio Gemini API, scores topic relevance against user-defined criteria, and emails a curated HTML news digest on weekdays.

## Architecture & Pipeline Flow

`runNewsPipeline()` (in [`main_runner.js`](./main_runner.js)) executes the following stages:

1. **Scrape (`Step 1a`)**: Runs modular site scrapers ([`scraper-channelnewsasia.js`](./scraper-channelnewsasia.js), [`scraper-straitstimes.js`](./scraper-straitstimes.js)). Each scraper pulls evenly across its configured RSS feeds using round-robin selection, filtering for articles published strictly after `LAST_RUN_AT`, and fetches full article HTML concurrently via `UrlFetchApp.fetchAll`.
2. **Store (`Step 1b`)**: Appends newly discovered articles (deduplicated by normalised URL against `Articles` and recent `Archived Articles`) to the `Articles` sheet and updates `LAST_RUN_AT` in the `Config` tab.
3. **Batch Summarise (`Step 1c`)**: Summarises unsummarised articles (~80–100 words each) in token-budgeted batches (`SUMMARY_MAX_ARTICLES_PER_BATCH`, default `20`) using a dynamic ID-keyed `required` JSON schema ([`buildSummaryResponseSchema`](./gemini.js)) so no article in a batch can be omitted.
4. **Holistic Cross-Agency Clustering & Relevance Scoring (`Step 1d`)**:
   - **Cluster & Synthesise**: Groups related articles covering the same underlying event or discussion across news agencies under a shared `Discussion Tag` and generates a ~100–150 word synthesised `Discussion Summary`.
   - **Score**: Evaluates all topic clusters in a single holistic call (`1–10`) against active topics in the `Topics` tab, storing the score, matched topics, and a 2–3 sentence (`~40–60 word`) justification.
5. **Weekday HTML Digest & Archiving (`Step 1e` & `Step 1f`)**:
   - **Weekdays (Mon–Fri, `Asia/Singapore`)**: Shortlists pending topics with `Relevance Score > SHORTLIST_ABOVE` (default `> 7`, i.e. `8–10`), renders [`email_template.html`](./email_template.html) (one card box per topic with synthesised summary and links to each constituent article), sends a BCC email to active recipients, and moves all processed rows from `Articles` to `Archived Articles`.
   - **Weekends (Sat–Sun)**: Skips emailing and archiving so weekend articles accumulate in `Articles` and merge into a single combined Monday digest.
   - **7-Day Archive Retention**: Automatically prunes rows in `Archived Articles` older than `ARCHIVE_RETENTION_DAYS` (default `7` days).

---

## Project Structure

| File | Purpose |
| :--- | :--- |
| [`main_runner.js`](./main_runner.js) | Main pipeline entry point (`runNewsPipeline`), custom GSheet menu (`onOpen`), sheet I/O, archiving, and shared RSS/HTML utilities |
| [`constants.js`](./constants.js) | Sheet schemas, RSS feed URLs, default configuration values, and externalised Gemini system prompts (`DEFAULT_PROMPTS`) |
| [`setup.js`](./setup.js) | `setupSheets()` and `setupDailyTrigger()` helpers to initialize the 5 spreadsheet tabs and daily time-driven trigger |
| [`scraper-channelnewsasia.js`](./scraper-channelnewsasia.js) | Channel NewsAsia scraper module (Latest, Singapore, Asia, Business feeds) |
| [`scraper-straitstimes.js`](./scraper-straitstimes.js) | The Straits Times scraper module (Singapore, Asia, Business, World feeds) |
| [`gemini.js`](./gemini.js) | Gemini REST API (`v1beta`) client, `models.get` token-limit discovery, dynamic schema builders, and batch/bisection runner |
| [`mailer.js`](./mailer.js) | Topic-box grouping and `MailApp` BCC digest sender |
| [`email_template.html`](./email_template.html) | Responsive inline-CSS HTML email template |
| [`tests/selfcheck.js`](./tests/selfcheck.js) | Local Node.js unit and integration self-check suite (excluded from Apps Script via [`.claspignore`](./.claspignore)) |

---

## Google Sheet Structure

Running **News Scraper → Initialize Sheets** (or `setupSheets()` in the Apps Script editor) automatically creates and formats these 5 tabs:

1. **`Articles`**: Active working queue (`Scraped At`, `Source`, `Published At`, `Title`, `URL`, `Content`, `Summary`, `Relevance Score`, `Matched Topics`, `Score Reason`, `Discussion Tag`, `Discussion Summary`, `Emailed`). Emptied into `Archived Articles` after each weekday digest.
2. **`Archived Articles`**: Historical & debugging archive with identical columns to `Articles` (retains raw `Content`, summaries, clusters, and scores).
3. **`Topics`**: User-defined topics of interest (`Topic Name`, `Description / Criteria`, `Active` checkbox).
4. **`Recipients`**: Digest mailing list (`Email`, `Name`, `Active` checkbox).
5. **`Config`**: Runtime settings and optional prompt overrides:
   - `LAST_RUN_AT`: ISO timestamp of the last completed scrape (auto-updated).
   - `GEMINI_MODEL`: Model ID (e.g. `gemini-3.8-flash`, `gemini-2.5-flash`).
   - `THINKING_LEVEL`: `low` (or blank to omit for non-thinking models).
   - `SUMMARY_MAX_ARTICLES_PER_BATCH`: Max full-text articles per Step 1c call (default `20`); articles are divided evenly across the minimum required runs (e.g. 50 articles at cap 40 → 2 runs of 25).
   - `MAX_INPUT_TOKENS_PER_CALL` / `MAX_OUTPUT_TOKENS_PER_CALL`: `AUTO` (uses dynamic `models.get` limits) or integer cap.
   - `SHORTLIST_ABOVE`: Minimum relevance score threshold (`7` = shortlists `8–10`).
   - `MAX_ARTICLES_PER_SITE`: Max new articles per scraper per run (default `40`, round-robin across feeds).
   - `MAX_ARTICLE_CHARS`: Max extracted characters stored per article (default `12000`).
   - `ARCHIVE_RETENTION_DAYS`: Auto-prune rows in `Archived Articles` older than this many days (default `7`).
   - `PROMPT_SUMMARIZE` / `PROMPT_SCORE` / `PROMPT_CLUSTER`: Leave blank to use `DEFAULT_PROMPTS` in [`constants.js`](./constants.js), or enter text to override in the sheet.

---

## Setup & Deployment

### 1. Local Development with `clasp`
```bash
# Log in to clasp
npx @google/clasp login

# Bind to your container-bound Apps Script project
npx @google/clasp clone <YOUR_SCRIPT_ID>

# Run local syntax & unit self-checks
node tests/selfcheck.js

# Push to Google Apps Script
npx @google/clasp push
```

### 2. Configure API Key & Triggers
1. Open the Apps Script editor (**Extensions → Apps Script** in your Google Sheet).
2. Go to **Project Settings → Script Properties** and add:
   - Property: `GEMINI_API`
   - Value: Your Google AI Studio API key
3. Refresh the Google Sheet and click **News Scraper → Initialize Sheets** to generate the 5 tabs.
4. Populate your **Topics** and **Recipients** tabs.
5. Click **News Scraper → Setup Daily Trigger** (or point a daily time-driven trigger at `runNewsPipeline`).

---

## Adding a New News Site Scraper

1. Create a new file named `scraper-<sitename>.js` exposing `function scrape<SiteName>(existingUrls, maxItems, maxChars, lastRunMs)` that returns an array of `{ source, publishedAt, title, url, content }` objects (using `fetchAllSafe`, `selectRssItemsRoundRobin`, and `extractArticleBody`).
2. Register the scraper inside `runNewsPipeline()` in [`main_runner.js`](./main_runner.js).

