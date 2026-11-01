/**
 * Spreadsheet setup & trigger provisioning based on "Google Sheet Structure Guide".
 * Uses schema and default constants defined in `constants.js`.
 */

/**
 * Creates and formats all required tabs (Articles, Archived Articles, Topics, Recipients, Config)
 * according to the Google Sheet Structure Guide without overwriting existing user data.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ssOpt]
 * @return {Object} Summary of created/updated tabs
 */
function setupSheets(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet found. Bind this script to a Google Sheet or pass a Spreadsheet instance.');
  }

  const summary = { created: [], verified: [] };
  const articleColWidths = [145, 110, 145, 260, 200, 240, 300, 115, 180, 240, 220, 320, 95];

  // 1. Articles tab (active queue — emptied into Archived Articles after weekday digest)
  const articlesSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.ARTICLES, ARTICLES_HEADERS, summary);
  setColumnWidths_(articlesSheet, articleColWidths);

  // 2. Archived Articles tab (debug & history store)
  const archivedSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.ARCHIVED_ARTICLES, ARTICLES_HEADERS, summary);
  setColumnWidths_(archivedSheet, articleColWidths);

  // 3. Topics tab
  const topicsSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.TOPICS, TOPICS_HEADERS, summary);
  setColumnWidths_(topicsSheet, [220, 480, 80]);
  if (topicsSheet.getLastRow() < 2) {
    topicsSheet.getRange(2, 1, DEFAULT_TOPICS_ROWS.length, TOPICS_HEADERS.length).setValues(DEFAULT_TOPICS_ROWS);
  }
  applyCheckboxValidation_(topicsSheet, 2, 3, 50);

  // 4. Recipients tab
  const recipientsSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.RECIPIENTS, RECIPIENTS_HEADERS, summary);
  setColumnWidths_(recipientsSheet, [260, 160, 80]);
  if (recipientsSheet.getLastRow() < 2) {
    let defaultEmail = 'you@example.com';
    try {
      const activeEmail = Session.getActiveUser().getEmail();
      if (activeEmail) defaultEmail = activeEmail;
    } catch (_) {}
    recipientsSheet.getRange(2, 1, 1, RECIPIENTS_HEADERS.length).setValues([[defaultEmail, 'Primary Recipient', true]]);
  }
  applyCheckboxValidation_(recipientsSheet, 2, 3, 50);

  // 5. Config tab (merge missing keys without clobbering existing user values)
  const configSheet = ensureSheetWithHeaders_(ss, SHEET_NAMES.CONFIG, CONFIG_HEADERS, summary);
  setColumnWidths_(configSheet, [230, 360, 440]);
  ensureConfigDefaults_(configSheet);

  console.log('setupSheets complete:', JSON.stringify(summary));
  return summary;
}

/**
 * Ensures a sheet exists and has the expected header row in row 1.
 */
function ensureSheetWithHeaders_(ss, name, headers, summary) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    summary.created.push(name);
  } else {
    summary.verified.push(name);
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existingHeaders = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim())
    : [];

  const needsHeaderWrite = headers.some((h, i) => existingHeaders[i] !== h);
  if (needsHeaderWrite) {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Ensures all default Config keys (including LAST_RUN_AT, SUMMARY_MAX_ARTICLES_PER_BATCH, and prompt keys)
 * exist in Config tab, and clears legacy auto-seeded prompt strings so `constants.js` updates take effect.
 */
function ensureConfigDefaults_(configSheet) {
  const lastRow = configSheet.getLastRow();
  const existingKeys = new Set();
  if (lastRow >= 2) {
    const rows = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      const key = String(rows[i][0] || '').trim();
      const val = String(rows[i][1] || '').trim();
      if (!key) continue;
      existingKeys.add(key);

      // Clear legacy auto-seeded short default prompts so updated DEFAULT_PROMPTS in constants.js apply
      if (
        (key === 'PROMPT_SUMMARIZE' && val.startsWith('You are an objective news analyst. Summarise each provided news article in 2 to 3 concise')) ||
        (key === 'PROMPT_SCORE' && val.includes('1-sentence `reason`')) ||
        (key === 'PROMPT_CLUSTER' && val.includes('2 to 4 sentence synthesis'))
      ) {
        configSheet.getRange(2 + i, 2).setValue('');
      }
    }
  }
  const missingRows = DEFAULT_CONFIG_ROWS.filter(([key]) => !existingKeys.has(key));
  if (missingRows.length > 0) {
    const startRow = Math.max(lastRow + 1, 2);
    configSheet.getRange(startRow, 1, missingRows.length, CONFIG_HEADERS.length).setValues(missingRows);
  }
}

/**
 * Applies boolean checkbox data validation to a column range.
 */
function applyCheckboxValidation_(sheet, startRow, col, numRows) {
  try {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
  } catch (_) {}
}

/**
 * Sets column widths safely.
 */
function setColumnWidths_(sheet, widths) {
  try {
    widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));
  } catch (_) {}
}

/**
 * Installs a daily time-driven trigger for runNewsPipeline (at 7am-8am SGT) if not already present.
 */
function setupDailyTrigger() {
  const targetFn = 'runNewsPipeline';
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === targetFn && t.getEventType() === ScriptApp.EventType.CLOCK) {
      console.log('Daily trigger for runNewsPipeline already exists (id=' + t.getUniqueId() + ').');
      return { created: false, triggerId: t.getUniqueId() };
    }
  }
  const trigger = ScriptApp.newTrigger(targetFn)
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  console.log('Created daily trigger for runNewsPipeline (id=' + trigger.getUniqueId() + ').');
  return { created: true, triggerId: trigger.getUniqueId() };
}
