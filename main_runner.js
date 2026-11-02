/**
 * Main orchestrator, sheet I/O helpers, archiving, and shared RSS/HTML scraping utilities.
 * Constants and prompts live in `constants.js`.
 */

/**
 * Adds custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('News Scraper')
    .addItem('Run Full Pipeline', 'runNewsPipeline')
    .addItem('Initialize Sheets', 'setupSheets')
    .addItem('Setup Daily Trigger', 'setupDailyTrigger')
    .addSeparator()
    .addItem('Summarise & Score Only', 'runSummarizeAndScoreOnly')
    .addItem('Send Pending Digest & Archive', 'runSendDigestOnly')
    .addItem('Archive Processed Articles Now', 'runArchiveOnly')
    .addItem('Prune Archived Articles (>7d)', 'runPruneArchivedOnly')
    .addToUi();
}

/**
 * Main trigger entry point:
 * 1a) Runs modular site scrapers for news published after `LAST_RUN_AT`
 * 1b) Stores new articles in `Articles` and updates `LAST_RUN_AT` in `Config`
 * 1c) Summarises unsummarised articles in token-budgeted batches via Gemini
 * 1d) Scores summaries against active Topics AND tags/synthesises related articles across news agencies
 * 1e) On weekdays (Mon-Fri): emails HTML digest of shortlisted topics and moves `Articles` -> `Archived Articles`.
 *     On weekends (Sat-Sun): skips emailing & archiving so `Articles` accumulates for Monday's combined digest.
 */
function runNewsPipeline(nowOpt) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    console.warn('Another runNewsPipeline execution is already active. Exiting.');
    return;
  }

  const runDate = nowOpt instanceof Date ? nowOpt : new Date();
  const startMs = Date.now();
  try {
    clearContinuationTrigger_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets(ss);

    const ctx = buildPipelineContext_(ss, startMs + DEADLINE_MS);

    // Step 1a & 1b: Scrape each site module for items newer than LAST_RUN_AT and append immediately
    const scrapers = [
      { name: 'CNA', fn: scrapeChannelNewsAsia },
      { name: 'Straits Times', fn: scrapeStraitsTimes },
    ];

    let allScrapersCompleted = true;
    for (const scraper of scrapers) {
      if (Date.now() >= ctx.deadlineAt) {
        ctx.deadlineTripped = true;
        allScrapersCompleted = false;
        console.warn('Deadline reached before running scraper: ' + scraper.name);
        break;
      }
      try {
        const items = scraper.fn(
          ctx.existingUrls,
          ctx.config.MAX_ARTICLES_PER_SITE,
          ctx.config.MAX_ARTICLE_CHARS,
          ctx.config.LAST_RUN_MS
        );
        if (items && items.length > 0) {
          appendArticles(ctx.articlesSheet, ctx.colMap, items);
          for (const it of items) {
            ctx.existingUrls.add(it.url);
          }
          console.log(scraper.name + ': appended ' + items.length + ' new articles.');
        } else {
          console.log(scraper.name + ': 0 new articles since LAST_RUN_AT.');
        }
      } catch (err) {
        allScrapersCompleted = false;
        console.error('Scraper failed [' + scraper.name + ']: ' + (err && err.stack ? err.stack : err));
      }
    }

    if (allScrapersCompleted) {
      const isoRunTime = formatIsoSgt(runDate);
      writeConfigValue(ctx.configSheet, 'LAST_RUN_AT', isoRunTime);
      ctx.config.LAST_RUN_AT = isoRunTime;
      ctx.config.LAST_RUN_MS = runDate.getTime();
    }

    // Step 1c: Batch summarise unsummarised articles
    if (Date.now() < ctx.deadlineAt) {
      summarizePending(ctx);
    } else {
      ctx.deadlineTripped = true;
    }

    // Step 1d: Batch score unscored summaries against active Topics + cluster & synthesise discussions
    if (Date.now() < ctx.deadlineAt) {
      scorePending(ctx);
    } else {
      ctx.deadlineTripped = true;
    }

    // Step 1e: Email digest & archive only on weekdays (Mon-Fri in Asia/Singapore)
    if (isWeekday(runDate)) {
      sendDigest(ctx);
      if (!ctx.deadlineTripped && canArchiveArticles(ctx.articlesSheet, ctx.colMap)) {
        const archivedCount = archiveArticles(ctx.articlesSheet, ctx.archivedSheet, ctx.colMap);
        if (archivedCount > 0) {
          console.log('Archived ' + archivedCount + ' article(s) to "' + SHEET_NAMES.ARCHIVED_ARTICLES + '".');
        }
      }
    } else {
      console.log('Weekend run: skipping email digest and keeping articles in "' + SHEET_NAMES.ARTICLES + '" for Monday digest.');
    }

    // Step 1f: Automatically prune Archived Articles older than ARCHIVE_RETENTION_DAYS (default 7 days)
    const prunedCount = pruneArchivedArticles(
      ctx.archivedSheet,
      ctx.colMap,
      ctx.config.ARCHIVE_RETENTION_DAYS,
      runDate
    );
    if (prunedCount > 0) {
      console.log(
        'Pruned ' +
          prunedCount +
          ' expired article(s) older than ' +
          ctx.config.ARCHIVE_RETENTION_DAYS +
          ' days from "' +
          SHEET_NAMES.ARCHIVED_ARTICLES +
          '".'
      );
    }

    if (ctx.deadlineTripped) {
      scheduleContinuationTrigger_();
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menu helper: runs only Gemini summarise + score/cluster steps on existing sheet rows.
 */
function runSummarizeAndScoreOnly() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets(ss);
    const ctx = buildPipelineContext_(ss, Date.now() + DEADLINE_MS);
    summarizePending(ctx);
    scorePending(ctx);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menu helper: sends pending shortlisted digest and archives processed rows immediately.
 */
function runSendDigestOnly() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets(ss);
    const ctx = buildPipelineContext_(ss, Date.now() + DEADLINE_MS, true);
    sendDigest(ctx);
    if (canArchiveArticles(ctx.articlesSheet, ctx.colMap)) {
      archiveArticles(ctx.articlesSheet, ctx.archivedSheet, ctx.colMap);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menu helper: archives all fully processed rows from `Articles` to `Archived Articles`.
 */
function runArchiveOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets(ss);
  const ctx = buildPipelineContext_(ss, Date.now() + DEADLINE_MS, true);
  if (canArchiveArticles(ctx.articlesSheet, ctx.colMap)) {
    return archiveArticles(ctx.articlesSheet, ctx.archivedSheet, ctx.colMap);
  }
  console.warn('Cannot archive yet: some rows in Articles are still unsummarised, unscored, or PENDING email.');
  return 0;
}

/**
 * Menu helper: prunes rows in `Archived Articles` older than `ARCHIVE_RETENTION_DAYS` (default 7 days).
 */
function runPruneArchivedOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets(ss);
  const ctx = buildPipelineContext_(ss, Date.now() + DEADLINE_MS, true);
  return pruneArchivedArticles(
    ctx.archivedSheet,
    ctx.colMap,
    ctx.config.ARCHIVE_RETENTION_DAYS,
    new Date()
  );
}

/**
 * Builds execution context from spreadsheet tabs and verifies Gemini API + model limits.
 */
function buildPipelineContext_(ss, deadlineAt, skipGeminiInit) {
  const articlesSheet = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  const archivedSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVED_ARTICLES);
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  const headers = articlesSheet
    .getRange(1, 1, 1, articlesSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());
  const colMap = buildColMap(headers);

  const config = readConfig(configSheet);
  const topics = readActiveTopics(ss.getSheetByName(SHEET_NAMES.TOPICS));
  const recipients = readActiveRecipients(ss.getSheetByName(SHEET_NAMES.RECIPIENTS));
  const existingUrls = readExistingUrls(articlesSheet, archivedSheet, colMap);

  let apiKey = '';
  let modelLimits = { inputTokenLimit: 1048576, outputTokenLimit: 65536 };
  if (!skipGeminiInit) {
    apiKey = getApiKey();
    modelLimits = getModelLimits(config.GEMINI_MODEL, apiKey);
  }

  return {
    ss,
    articlesSheet,
    archivedSheet,
    configSheet,
    headers,
    colMap,
    config,
    topics,
    recipients,
    existingUrls,
    apiKey,
    modelLimits,
    deadlineAt,
    deadlineTripped: false,
  };
}

/**
 * Maps header names to 0-based column indices. Throws if a required header is missing.
 */
function buildColMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i;
  });
  for (const req of ARTICLES_HEADERS) {
    if (typeof map[req] !== 'number') {
      throw new Error('Articles sheet is missing required column header: "' + req + '"');
    }
  }
  return map;
}

/**
 * Parses a token limit setting from Config: returns null for "AUTO" or blank (meaning use models.get dynamic limit),
 * or a positive integer when manually capped.
 */
function parseTokenLimitSetting_(val, minFloor) {
  const s = String(val ?? '').trim().toUpperCase();
  if (!s || s === 'AUTO') return null;
  const n = parseInt(s, 10);
  return Number.isFinite(n) && n > 0 ? Math.max(minFloor, n) : null;
}

/**
 * Parses LAST_RUN_AT (Date object or ISO/date string) into epoch milliseconds, or 0 if unset.
 */
function parseLastRunMs(val) {
  if (!val) return 0;
  if (val instanceof Date) {
    const ms = val.getTime();
    return Number.isFinite(ms) ? ms : 0;
  }
  const s = String(val).trim();
  if (!s) return 0;
  const ms = Date.parse(s);
  return Number.isFinite(ms) ? ms : 0;
}

/**
 * Reads Config tab key-value pairs (including LAST_RUN_AT and optional prompt overrides) with fallbacks to `constants.js`.
 */
function readConfig(configSheet) {
  const raw = {};
  let rawLastRunVal = '';
  if (configSheet && configSheet.getLastRow() >= 2) {
    const rows = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
    for (const [k, v] of rows) {
      const key = String(k || '').trim();
      if (!key) continue;
      if (key === 'LAST_RUN_AT') rawLastRunVal = v;
      raw[key] = String(v ?? '').trim();
    }
  }
  return {
    LAST_RUN_AT: raw.LAST_RUN_AT || '',
    LAST_RUN_MS: parseLastRunMs(rawLastRunVal),
    GEMINI_MODEL: raw.GEMINI_MODEL || 'gemini-3.8-flash',
    THINKING_LEVEL: raw.THINKING_LEVEL !== undefined ? raw.THINKING_LEVEL : 'low',
    SUMMARY_MAX_ARTICLES_PER_BATCH: parseTokenLimitSetting_(
      raw.SUMMARY_MAX_ARTICLES_PER_BATCH !== undefined ? raw.SUMMARY_MAX_ARTICLES_PER_BATCH : '20',
      1
    ),
    MAX_INPUT_TOKENS_PER_CALL: parseTokenLimitSetting_(raw.MAX_INPUT_TOKENS_PER_CALL, 1000),
    MAX_OUTPUT_TOKENS_PER_CALL: parseTokenLimitSetting_(raw.MAX_OUTPUT_TOKENS_PER_CALL, 512),
    SHORTLIST_ABOVE: parseInt(raw.SHORTLIST_ABOVE || '7', 10) || 7,
    MAX_ARTICLES_PER_SITE: Math.max(1, parseInt(raw.MAX_ARTICLES_PER_SITE || '40', 10) || 40),
    MAX_ARTICLE_CHARS: Math.max(500, parseInt(raw.MAX_ARTICLE_CHARS || '12000', 10) || 12000),
    ARCHIVE_RETENTION_DAYS: Math.max(
      1,
      parseInt(raw.ARCHIVE_RETENTION_DAYS || String(DEFAULT_ARCHIVE_RETENTION_DAYS), 10) ||
        DEFAULT_ARCHIVE_RETENTION_DAYS
    ),
    PROMPT_SUMMARIZE: raw.PROMPT_SUMMARIZE || DEFAULT_PROMPTS.PROMPT_SUMMARIZE,
    PROMPT_SCORE: raw.PROMPT_SCORE || DEFAULT_PROMPTS.PROMPT_SCORE,
    PROMPT_CLUSTER: raw.PROMPT_CLUSTER || DEFAULT_PROMPTS.PROMPT_CLUSTER,
  };
}

/**
 * Writes or updates a single key's Value (Col B) in the Config sheet.
 */
function writeConfigValue(configSheet, key, value) {
  if (!configSheet) return;
  const lastRow = configSheet.getLastRow();
  if (lastRow >= 2) {
    const keys = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === key) {
        configSheet.getRange(2 + i, 2).setValue(value);
        return;
      }
    }
  }
  configSheet.getRange(Math.max(lastRow + 1, 2), 1, 1, 3).setValues([[key, value, '']]);
}

/**
 * Formats a Date as an ISO 8601 timestamp in Asia/Singapore (+08:00).
 */
function formatIsoSgt(date) {
  const d = date instanceof Date ? date : new Date();
  try {
    if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
      return Utilities.formatDate(d, 'Asia/Singapore', "yyyy-MM-dd'T'HH:mm:ss'+08:00'");
    }
  } catch (_) {}
  const sgtMs = d.getTime() + 8 * 3600 * 1000;
  return new Date(sgtMs).toISOString().replace(/\.\d{3}Z$/, '+08:00');
}

/**
 * Returns true if `date` falls on Monday-Friday in Asia/Singapore timezone, false on Saturday/Sunday.
 */
function isWeekday(date) {
  const d = date instanceof Date ? date : new Date();
  try {
    if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
      const isoDay = parseInt(Utilities.formatDate(d, 'Asia/Singapore', 'u'), 10);
      return isoDay >= 1 && isoDay <= 5;
    }
  } catch (_) {}
  // Fallback for Node test environment: shift UTC by +8h (Asia/Singapore) and check getUTCDay() (1=Mon..5=Fri)
  const sgtDay = new Date(d.getTime() + 8 * 3600 * 1000).getUTCDay();
  return sgtDay >= 1 && sgtDay <= 5;
}

/**
 * Reads active topics from Topics tab.
 */
function readActiveTopics(topicsSheet) {
  if (!topicsSheet || topicsSheet.getLastRow() < 2) return [];
  const rows = topicsSheet.getRange(2, 1, topicsSheet.getLastRow() - 1, 3).getValues();
  const topics = [];
  for (const [name, criteria, active] of rows) {
    const cleanName = String(name || '').trim();
    const cleanCriteria = String(criteria || '').trim();
    if (cleanName && isTruthyFlag_(active)) {
      topics.push({ name: cleanName, criteria: cleanCriteria });
    }
  }
  return topics;
}

/**
 * Reads active recipient email addresses from Recipients tab.
 */
function readActiveRecipients(recipientsSheet) {
  if (!recipientsSheet || recipientsSheet.getLastRow() < 2) return [];
  const rows = recipientsSheet.getRange(2, 1, recipientsSheet.getLastRow() - 1, 3).getValues();
  const emails = [];
  for (const [email, name, active] of rows) {
    const cleanEmail = String(email || '').trim();
    if (cleanEmail && cleanEmail.includes('@') && !cleanEmail.endsWith('@example.com') && isTruthyFlag_(active)) {
      emails.push({ email: cleanEmail, name: String(name || '').trim() });
    }
  }
  return emails;
}

function isTruthyFlag_(val) {
  if (val === true) return true;
  const s = String(val || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}

/**
 * Reads URL column from `Articles` plus recent tail of `Archived Articles` for URL deduplication.
 */
function readExistingUrls(articlesSheet, archivedSheet, colMap) {
  const urls = new Set();
  const urlCol = colMap['URL'] + 1;

  if (articlesSheet && articlesSheet.getLastRow() >= 2) {
    const values = articlesSheet.getRange(2, urlCol, articlesSheet.getLastRow() - 1, 1).getValues();
    for (const [u] of values) {
      const norm = normalizeUrl(u);
      if (norm) urls.add(norm);
    }
  }

  if (archivedSheet && archivedSheet.getLastRow() >= 2) {
    const archLast = archivedSheet.getLastRow();
    const count = Math.min(archLast - 1, ARCHIVE_DEDUP_LOOKBACK_ROWS);
    const startRow = archLast - count + 1;
    const values = archivedSheet.getRange(startRow, urlCol, count, 1).getValues();
    for (const [u] of values) {
      const norm = normalizeUrl(u);
      if (norm) urls.add(norm);
    }
  }

  return urls;
}

/**
 * Appends newly scraped articles to the bottom of the Articles sheet in one batch write.
 */
function appendArticles(articlesSheet, colMap, items) {
  if (!items || items.length === 0) return;
  const numCols = Object.keys(colMap).length;
  const now = new Date();
  const rows = items.map(item => {
    const row = new Array(numCols).fill('');
    row[colMap['Scraped At']] = now;
    row[colMap['Source']] = item.source || '';
    row[colMap['Published At']] = item.publishedAt || '';
    row[colMap['Title']] = item.title || '';
    row[colMap['URL']] = item.url || '';
    row[colMap['Content']] = item.content || '';
    return row;
  });
  const startRow = articlesSheet.getLastRow() + 1;
  articlesSheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
}

/**
 * Returns true only if every row in `Articles` has completed summarisation, topic clustering,
 * scoring, and emailing without any `[ERROR]` markers or `PENDING` email status.
 * Blocks archiving if any row failed so errors remain visible in `Articles`.
 */
function canArchiveArticles(articlesSheet, colMap) {
  if (!articlesSheet || articlesSheet.getLastRow() < 2) return false;
  const numRows = articlesSheet.getLastRow() - 1;
  const summaryVals = articlesSheet.getRange(2, colMap['Summary'] + 1, numRows, 1).getValues();
  const scoreVals = articlesSheet.getRange(2, colMap['Relevance Score'] + 1, numRows, 1).getValues();
  const reasonVals = articlesSheet.getRange(2, colMap['Score Reason'] + 1, numRows, 1).getValues();
  const tagVals = articlesSheet.getRange(2, colMap['Discussion Tag'] + 1, numRows, 1).getValues();
  const tagSumVals = articlesSheet.getRange(2, colMap['Discussion Summary'] + 1, numRows, 1).getValues();
  const emailedVals = articlesSheet.getRange(2, colMap['Emailed'] + 1, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const sum = String(summaryVals[i][0] || '').trim();
    const scoreRaw = String(scoreVals[i][0] ?? '').trim();
    const scoreNum = Number(scoreRaw);
    const reason = String(reasonVals[i][0] || '').trim();
    const tag = String(tagVals[i][0] || '').trim();
    const tagSum = String(tagSumVals[i][0] || '').trim();
    const emailed = String(emailedVals[i][0] || '').trim().toUpperCase();

    if (!sum || sum.startsWith('[ERROR]')) {
      console.warn('Blocking archive: row ' + (2 + i) + ' has missing or [ERROR] Summary.');
      return false;
    }
    if (scoreRaw === '' || !Number.isInteger(scoreNum) || scoreNum < 1 || reason.startsWith('[ERROR]')) {
      console.warn('Blocking archive: row ' + (2 + i) + ' has missing or [ERROR] Relevance Score.');
      return false;
    }
    if (!tag || !tagSum) {
      console.warn('Blocking archive: row ' + (2 + i) + ' is missing Discussion Tag / Discussion Summary.');
      return false;
    }
    if (emailed === 'PENDING' || !emailed) {
      console.warn('Blocking archive: row ' + (2 + i) + ' still has Emailed=' + (emailed || 'EMPTY') + '.');
      return false;
    }
  }
  return true;
}

/**
 * Moves all data rows (including raw `Content` for debugging) from `Articles` to `Archived Articles`,
 * then clears `Articles` data rows so `Articles` is empty for the next run.
 */
function archiveArticles(articlesSheet, archivedSheet, colMap) {
  if (!articlesSheet || !archivedSheet) return 0;
  const lastRow = articlesSheet.getLastRow();
  if (lastRow < 2) return 0;

  const numRows = lastRow - 1;
  const numCols = Object.keys(colMap).length;
  const dataRange = articlesSheet.getRange(2, 1, numRows, numCols);
  const rows = dataRange.getValues();

  const destStartRow = archivedSheet.getLastRow() + 1;
  archivedSheet.getRange(destStartRow, 1, rows.length, numCols).setValues(rows);

  dataRange.clearContent();
  return rows.length;
}

/**
 * Automatically prunes rows from `Archived Articles` whose `Scraped At` timestamp is older
 * than `retentionDays` (default 7 days). Uses a single batch clear + rewrite of retained rows.
 */
function pruneArchivedArticles(archivedSheet, colMap, retentionDaysOpt, nowOpt) {
  if (!archivedSheet) return 0;
  const lastRow = archivedSheet.getLastRow();
  if (lastRow < 2) return 0;

  const days = Math.max(1, Number(retentionDaysOpt) || DEFAULT_ARCHIVE_RETENTION_DAYS);
  const nowMs = nowOpt instanceof Date ? nowOpt.getTime() : Date.now();
  const cutoffMs = nowMs - days * 24 * 60 * 60 * 1000;

  const numRows = lastRow - 1;
  const numCols = Object.keys(colMap).length;
  const dataRange = archivedSheet.getRange(2, 1, numRows, numCols);
  const allRows = dataRange.getValues();
  const scrapedIdx = colMap['Scraped At'];

  const keptRows = [];
  for (let i = 0; i < allRows.length; i++) {
    const rawDate = allRows[i][scrapedIdx];
    const ts = rawDate instanceof Date ? rawDate.getTime() : Date.parse(String(rawDate || ''));
    if (isNaN(ts) || ts > cutoffMs) {
      keptRows.push(allRows[i]);
    }
  }

  const prunedCount = allRows.length - keptRows.length;
  if (prunedCount > 0) {
    dataRange.clearContent();
    if (keptRows.length > 0) {
      archivedSheet.getRange(2, 1, keptRows.length, numCols).setValues(keptRows);
    }
  }
  return prunedCount;
}

/**
 * Schedules a one-off continuation trigger 60s later when a run hits the 4.5-min soft deadline.
 */
function scheduleContinuationTrigger_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(CONTINUATION_PROP_KEY)) return;
    const trigger = ScriptApp.newTrigger('runNewsPipeline')
      .timeBased()
      .after(60 * 1000)
      .create();
    props.setProperty(CONTINUATION_PROP_KEY, trigger.getUniqueId());
    console.log('Scheduled 1-min continuation trigger: ' + trigger.getUniqueId());
  } catch (err) {
    console.error('Could not schedule continuation trigger: ' + err);
  }
}

/**
 * Deletes the one-off continuation trigger if one was registered by the previous run.
 */
function clearContinuationTrigger_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const triggerId = props.getProperty(CONTINUATION_PROP_KEY);
    if (!triggerId) return;
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(t);
      }
    }
    props.deleteProperty(CONTINUATION_PROP_KEY);
  } catch (err) {
    console.warn('Could not clear continuation trigger: ' + err);
  }
}

// ============================================================================
// Shared RSS & HTML Scraping Utilities
// ============================================================================

/**
 * Normalises article URLs by stripping query parameters, hashes, and trailing slashes.
 */
function normalizeUrl(rawUrl) {
  const s = String(rawUrl || '').trim();
  if (!s) return '';
  return s
    .replace(/#.*$/, '')
    .replace(/\?.*$/, '')
    .replace(/\/+$/, '');
}

/**
 * Parses an RSS 2.0 XML string via XmlService and returns deduplicated new items
 * whose `<pubDate>` is strictly newer than `lastRunMs` (when `lastRunMs > 0`).
 */
function parseRssItems(xmlText, sourceName, existingUrls, seenInRun, maxItems, lastRunMs) {
  const items = [];
  if (!xmlText) return items;

  const doc = XmlService.parse(xmlText);
  const root = doc.getRootElement();
  const channel = root.getChild('channel');
  if (!channel) return items;

  const cutoffMs = Number(lastRunMs) || 0;
  const itemEls = channel.getChildren('item');

  for (const el of itemEls) {
    if (items.length >= maxItems) break;
    const rawLink = el.getChildText('link') || '';
    const url = normalizeUrl(rawLink);
    if (!url || !/^https?:\/\//i.test(url)) continue;
    if (existingUrls.has(url) || seenInRun.has(url)) continue;

    const pubDate = (el.getChildText('pubDate') || '').trim();
    if (cutoffMs > 0 && pubDate) {
      const pubMs = Date.parse(pubDate);
      if (Number.isFinite(pubMs) && pubMs <= cutoffMs) {
        continue;
      }
    }

    seenInRun.add(url);
    const title = decodeEntities(stripHtml(el.getChildText('title') || '')).trim();
    const description = decodeEntities(stripHtml(el.getChildText('description') || '')).trim();

    if (title) {
      items.push({
        source: sourceName,
        publishedAt: pubDate,
        title,
        url,
        rssDescription: description,
      });
    }
  }
  return items;
}

/**
 * Parses multiple RSS feed responses for a site and selects up to `maxItems` articles
 * using round-robin across feeds so Feed #1 never starves later section feeds (e.g. Asia, Business, World).
 */
function selectRssItemsRoundRobin(feedResponses, feedUrls, sourceName, existingUrls, maxItems, lastRunMs) {
  const limit = maxItems || 40;
  const seenInRun = new Set();
  const perFeedQueues = [];

  for (let i = 0; i < feedResponses.length; i++) {
    const res = feedResponses[i];
    if (!res || res.getResponseCode() !== 200) continue;
    try {
      const feedItems = parseRssItems(
        res.getContentText(),
        sourceName,
        existingUrls,
        seenInRun,
        limit,
        lastRunMs
      );
      if (feedItems.length > 0) {
        perFeedQueues.push(feedItems);
      }
    } catch (err) {
      const feedLabel = (feedUrls && feedUrls[i]) || 'feed#' + i;
      console.warn(sourceName + ' RSS parse error for ' + feedLabel + ': ' + err);
    }
  }

  const selected = [];
  let addedInPass = true;
  while (selected.length < limit && addedInPass) {
    addedInPass = false;
    for (const queue of perFeedQueues) {
      if (selected.length >= limit) break;
      if (queue.length > 0) {
        selected.push(queue.shift());
        addedInPass = true;
      }
    }
  }
  return selected;
}

/**
 * Fetches multiple HTTP requests concurrently via UrlFetchApp.fetchAll,
 * falling back to individual fetches if a batch-level DNS/timeout exception occurs.
 */
function fetchAllSafe(requests) {
  if (!requests || requests.length === 0) return [];
  const formatted = requests.map(req =>
    typeof req === 'string'
      ? { url: req, method: 'get', muteHttpExceptions: true, followRedirects: true }
      : Object.assign({ method: 'get', muteHttpExceptions: true, followRedirects: true }, req)
  );

  try {
    return UrlFetchApp.fetchAll(formatted);
  } catch (batchErr) {
    console.warn('UrlFetchApp.fetchAll threw (' + batchErr + '); falling back to per-URL fetch.');
    return formatted.map(req => {
      try {
        return UrlFetchApp.fetch(req.url, req);
      } catch (singleErr) {
        console.warn('Single fetch failed for ' + req.url + ': ' + singleErr);
        return null;
      }
    });
  }
}

/**
 * Extracts readable article body text from HTML:
 * 1) Checks <script type="application/ld+json"> for articleBody (including @graph arrays)
 * 2) Falls back to <article> or main content <p> paragraphs
 * 3) Truncates cleanly to maxChars
 */
function extractArticleBody(html, maxChars) {
  const cap = maxChars || 12000;
  if (!html) return '';

  const ldRegex = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = ldRegex.exec(html)) !== null) {
    try {
      const parsed = JSON.parse(match[1].trim());
      const body = findJsonLdArticleBody_(parsed);
      if (body && body.length >= 120) {
        return cleanAndCapText_(body, cap);
      }
    } catch (_) {}
  }

  const articleMatch = /<article\b[^>]*>([\s\S]*?)<\/article>/i.exec(html);
  const scopeHtml = (articleMatch ? articleMatch[1] : html)
    .replace(/<(script|style|noscript|svg|footer|nav|aside|header|form)\b[^>]*>[\s\S]*?<\/\1>/gi, ' ');

  const pRegex = /<p\b[^>]*>([\s\S]*?)<\/p>/gi;
  const paragraphs = [];
  let pMatch;
  while ((pMatch = pRegex.exec(scopeHtml)) !== null) {
    const text = decodeEntities(stripHtml(pMatch[1])).replace(/\s+/g, ' ').trim();
    if (text.length >= 40 && !isBoilerplateParagraph_(text)) {
      paragraphs.push(text);
    }
  }

  const joined = paragraphs.join('\n\n');
  return cleanAndCapText_(joined, cap);
}

function findJsonLdArticleBody_(node) {
  if (!node || typeof node !== 'object') return '';
  if (typeof node.articleBody === 'string' && node.articleBody.trim()) {
    return node.articleBody.trim();
  }
  if (Array.isArray(node)) {
    for (const item of node) {
      const found = findJsonLdArticleBody_(item);
      if (found) return found;
    }
  }
  if (Array.isArray(node['@graph'])) {
    return findJsonLdArticleBody_(node['@graph']);
  }
  return '';
}

function isBoilerplateParagraph_(text) {
  return /^(sign up for|subscribe to|follow us on|download the .* app|read more:|also read:|copyright ©)/i.test(text);
}

function cleanAndCapText_(raw, cap) {
  const cleaned = decodeEntities(stripHtml(raw))
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return cleaned.length > cap ? cleaned.slice(0, cap) : cleaned;
}

function stripHtml(str) {
  return String(str || '').replace(/<[^>]+>/g, ' ');
}

function decodeEntities(str) {
  return String(str || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#0*39;|&apos;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&rsquo;|&lsquo;/gi, "'")
    .replace(/&ldquo;|&rdquo;/gi, '"')
    .replace(/&ndash;|&mdash;/gi, '-')
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(parseInt(dec, 10)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}
