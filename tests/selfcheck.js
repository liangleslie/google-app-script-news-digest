const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        const val = (this.sheet.grid[this.row - 1 + r] || [])[this.col - 1 + c];
        rowArr.push(val !== undefined ? val : '');
      }
      out.push(rowArr);
    }
    return out;
  }
  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      const targetR = this.row - 1 + r;
      if (!this.sheet.grid[targetR]) this.sheet.grid[targetR] = [];
      for (let c = 0; c < values[r].length; c++) {
        this.sheet.grid[targetR][this.col - 1 + c] = values[r][c];
      }
    }
    this.sheet.trimTrailingEmptyRows_();
  }
  setValue(val) {
    this.setValues([[val]]);
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      const targetR = this.row - 1 + r;
      if (!this.sheet.grid[targetR]) continue;
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.grid[targetR][this.col - 1 + c] = '';
      }
    }
    this.sheet.trimTrailingEmptyRows_();
  }
  setFontWeight() {}
  setBackground() {}
  setDataValidation() {}
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.grid = [];
    this.frozenRows = 0;
  }
  trimTrailingEmptyRows_() {
    while (this.grid.length > 0) {
      const last = this.grid[this.grid.length - 1] || [];
      if (last.every(cell => cell === '' || cell === undefined)) {
        this.grid.pop();
      } else {
        break;
      }
    }
  }
  getLastRow() {
    return this.grid.length;
  }
  getLastColumn() {
    return this.grid.reduce((max, r) => Math.max(max, r ? r.length : 0), 0);
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange(this, row, col, numRows, numCols);
  }
  setFrozenRows(n) {
    this.frozenRows = n;
  }
  setColumnWidth() {}
}

class MockSpreadsheet {
  constructor() {
    this.sheets = new Map();
  }
  getSheetByName(name) {
    return this.sheets.get(name) || null;
  }
  insertSheet(name) {
    const s = new MockSheet(name);
    this.sheets.set(name, s);
    return s;
  }
}

// Lightweight XmlService mock for testing parseRssItems with LAST_RUN_AT cutoff
const MockXmlService = {
  parse(xmlText) {
    const itemRegex = /<item>([\s\S]*?)<\/item>/gi;
    const items = [];
    let m;
    while ((m = itemRegex.exec(xmlText)) !== null) {
      const block = m[1];
      const getTag = tag => {
        const tm = new RegExp('<' + tag + '>([\\s\\S]*?)<\\/' + tag + '>', 'i').exec(block);
        return tm ? tm[1] : '';
      };
      items.push({
        getChildText: tag => getTag(tag),
      });
    }
    return {
      getRootElement: () => ({
        getChild: name =>
          name === 'channel'
            ? {
                getChildren: childName => (childName === 'item' ? items : []),
              }
            : null,
      }),
    };
  },
};

const sandbox = {
  console,
  Date,
  Math,
  JSON,
  Set,
  Map,
  Array,
  Object,
  String,
  Number,
  RegExp,
  Error,
  parseInt,
  parseFloat,
  isNaN,
  encodeURIComponent,
  XmlService: MockXmlService,
  SpreadsheetApp: {
    newDataValidation: () => ({
      requireCheckbox: () => ({ build: () => ({}) }),
    }),
  },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'leslie@test.local' }),
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: k => (k === 'GEMINI_API' ? 'test-api-key' : null),
    }),
  },
};

vm.createContext(sandbox);

const files = [
  'constants.js',
  'setup.js',
  'main_runner.js',
  'scraper-channelnewsasia.js',
  'scraper-straitstimes.js',
  'gemini.js',
  'mailer.js',
];

for (const file of files) {
  const code = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
  vm.runInContext(code, sandbox, { filename: file });
}

// 1. Verify setupSheets creates all 5 tabs (Articles, Archived Articles, Topics, Recipients, Config)
const mockSs = new MockSpreadsheet();
const setupRes = sandbox.setupSheets(mockSs);
assert.deepStrictEqual(Array.from(setupRes.created), [
  'Articles',
  'Archived Articles',
  'Topics',
  'Recipients',
  'Config',
]);

const configSheet = mockSs.getSheetByName('Config');
let cfg = sandbox.readConfig(configSheet);
assert.strictEqual(cfg.LAST_RUN_AT, '');
assert.strictEqual(cfg.LAST_RUN_MS, 0);
assert.strictEqual(cfg.GEMINI_MODEL, 'gemini-3.8-flash');
assert.strictEqual(cfg.SUMMARY_MAX_ARTICLES_PER_BATCH, 20);
assert.strictEqual(cfg.MAX_INPUT_TOKENS_PER_CALL, null);
assert.strictEqual(cfg.MAX_OUTPUT_TOKENS_PER_CALL, null);
assert.ok(cfg.PROMPT_SUMMARIZE.includes('80 to 100 words'));
assert.ok(cfg.PROMPT_CLUSTER.includes('100 to 150 words'));
assert.ok(cfg.PROMPT_SCORE.includes('40 to 60 words'));

// 2. Verify writeConfigValue updates LAST_RUN_AT and readConfig parses LAST_RUN_MS
const testLastRunIso = '2026-11-01T08:00:00+08:00';
sandbox.writeConfigValue(configSheet, 'LAST_RUN_AT', testLastRunIso);
cfg = sandbox.readConfig(configSheet);
assert.strictEqual(cfg.LAST_RUN_AT, testLastRunIso);
assert.strictEqual(cfg.LAST_RUN_MS, Date.parse(testLastRunIso));

// 3. Verify parseRssItems skips articles published at or before LAST_RUN_AT and keeps newer articles
const sampleRss = `
<rss><channel>
  <item>
    <title>New Article After Last Run</title>
    <link>https://www.channelnewsasia.com/singapore/new-article?utm=1</link>
    <pubDate>Sun, 01 Nov 2026 09:30:00 +0800</pubDate>
    <description>New article description</description>
  </item>
  <item>
    <title>Exact Same Time As Last Run (Should Skip)</title>
    <link>https://www.channelnewsasia.com/singapore/exact-time</link>
    <pubDate>Sun, 01 Nov 2026 08:00:00 +0800</pubDate>
    <description>Exact time description</description>
  </item>
  <item>
    <title>Stale Article Before Last Run (Should Skip)</title>
    <link>https://www.channelnewsasia.com/singapore/old-article</link>
    <pubDate>Sat, 31 Oct 2026 22:15:00 +0800</pubDate>
    <description>Old article description</description>
  </item>
</channel></rss>`;

const rssItems = sandbox.parseRssItems(
  sampleRss,
  'CNA',
  new Set(),
  new Set(),
  20,
  cfg.LAST_RUN_MS
);
assert.strictEqual(rssItems.length, 1);
assert.strictEqual(rssItems[0].title, 'New Article After Last Run');
assert.strictEqual(rssItems[0].url, 'https://www.channelnewsasia.com/singapore/new-article');

// 4. Verify isWeekday (Asia/Singapore timezone)
// Sunday 2026-11-01 10:00 SGT -> false
assert.strictEqual(sandbox.isWeekday(new Date('2026-11-01T10:00:00+08:00')), false);
// Saturday 2026-10-31 23:00 SGT -> false
assert.strictEqual(sandbox.isWeekday(new Date('2026-10-31T23:00:00+08:00')), false);
// Monday 2026-11-02 07:00 SGT -> true
assert.strictEqual(sandbox.isWeekday(new Date('2026-11-02T07:00:00+08:00')), true);
// Friday 2026-11-06 18:00 SGT -> true
assert.strictEqual(sandbox.isWeekday(new Date('2026-11-06T18:00:00+08:00')), true);

// 5. Verify canArchiveArticles & archiveArticles (Articles -> Archived Articles)
const artSheet = mockSs.getSheetByName('Articles');
const archSheet = mockSs.getSheetByName('Archived Articles');
const colMap = sandbox.buildColMap(artSheet.getRange(1, 1, 1, 13).getValues()[0]);

// Simulate 2 weekend articles sitting in Articles (both grouped/tagged; one PENDING email, one SKIPPED)
artSheet.getRange(2, 1, 2, 13).setValues([
  [new Date('2026-11-01T08:00:00+08:00'), 'CNA', 'Sun, 01 Nov 2026', 'Weekend Story 1', 'https://cna.sg/1', 'Full raw body 1', 'Summary 1', 9, 'AI', 'Match', 'AI Tag', 'Synthesised AI Summary', 'PENDING'],
  [new Date('2026-11-01T09:00:00+08:00'), 'Straits Times', 'Sun, 01 Nov 2026', 'Weekend Story 2', 'https://st.sg/2', 'Full raw body 2', 'Summary 2', 4, '', 'Low score', 'Other Tag', 'Other Topic Summary', 'SKIPPED'],
]);

// While row 2 is still PENDING email, canArchiveArticles must be false
assert.strictEqual(sandbox.canArchiveArticles(artSheet, colMap), false);

// If a row has [ERROR] in Summary or Score Reason, canArchiveArticles must also block archiving
artSheet.getRange(2, colMap['Emailed'] + 1).setValue('YES');
artSheet.getRange(3, colMap['Summary'] + 1).setValue('[ERROR] Simulated failure');
assert.strictEqual(sandbox.canArchiveArticles(artSheet, colMap), false);

// Restore valid Summary on row 3 -> now all rows are valid, clustered, scored, and non-PENDING
artSheet.getRange(3, colMap['Summary'] + 1).setValue('Summary 2');
assert.strictEqual(sandbox.canArchiveArticles(artSheet, colMap), true);

// Archive moves all rows (with raw Content intact) to Archived Articles and empties Articles
const archivedCount = sandbox.archiveArticles(artSheet, archSheet, colMap);
assert.strictEqual(archivedCount, 2);
assert.strictEqual(artSheet.getLastRow(), 1); // Only header row remains in Articles!
assert.strictEqual(archSheet.getLastRow(), 3); // Header + 2 archived rows in Archived Articles
const archivedRows = archSheet.getRange(2, 1, 2, 13).getValues();
assert.strictEqual(archivedRows[0][colMap['Content']], 'Full raw body 1');
assert.strictEqual(archivedRows[1][colMap['Content']], 'Full raw body 2');

// 6. Verify selectRssItemsRoundRobin pulls evenly across all feeds instead of starving Feed #2/#3/#4
const makeFeedRes = urls => ({
  getResponseCode: () => 200,
  getContentText: () =>
    '<rss><channel>' +
    urls
      .map(
        (u, idx) =>
          `<item><title>Title ${u}</title><link>${u}</link><pubDate>Mon, 02 Nov 2026 06:00:00 +0800</pubDate><description>Desc ${idx}</description></item>`
      )
      .join('') +
    '</channel></rss>',
});
const feedResList = [
  makeFeedRes(['https://cna.sg/f1-a', 'https://cna.sg/f1-b', 'https://cna.sg/f1-c']),
  makeFeedRes(['https://cna.sg/f2-a', 'https://cna.sg/f2-b', 'https://cna.sg/f2-c']),
  makeFeedRes(['https://cna.sg/f3-a', 'https://cna.sg/f3-b', 'https://cna.sg/f3-c']),
  makeFeedRes(['https://cna.sg/f4-a', 'https://cna.sg/f4-b', 'https://cna.sg/f4-c']),
];
const roundRobinSelected = sandbox.selectRssItemsRoundRobin(
  feedResList,
  ['f1', 'f2', 'f3', 'f4'],
  'CNA',
  new Set(),
  4,
  0
);
assert.deepStrictEqual(
  Array.from(roundRobinSelected.map(it => it.url)),
  ['https://cna.sg/f1-a', 'https://cna.sg/f2-a', 'https://cna.sg/f3-a', 'https://cna.sg/f4-a']
);

// 7. Verify dynamic ID-keyed schemas & validators (`buildSummaryResponseSchema`, `buildScoreResponseSchema`)
const sumSchema = sandbox.buildSummaryResponseSchema(['a0', 'a1', 'a2']);
assert.deepStrictEqual(Array.from(sumSchema.required), ['a0', 'a1', 'a2']);
assert.deepStrictEqual(Array.from(sumSchema.propertyOrdering), ['a0', 'a1', 'a2']);

const keyedIdMap = new Map([
  ['a0', { rowNumber: 2 }],
  ['a1', { rowNumber: 3 }],
]);
const keyedSumValid = sandbox.validateSummaryResults(
  { a0: 'Summary zero', a1: 'Summary one' },
  keyedIdMap
);
assert.strictEqual(keyedSumValid.length, 2);
assert.strictEqual(keyedSumValid[0].summary, 'Summary zero');
assert.strictEqual(keyedSumValid[1].summary, 'Summary one');

const scoreSchema = sandbox.buildScoreResponseSchema(['Singapore AI & Tech Policy']);
assert.deepStrictEqual(Array.from(scoreSchema.required), ['results']);
const keyedScoreMap = new Map([
  ['s0', { rowNumbers: [2, 3] }],
  ['s1', { rowNumbers: [4] }],
]);
const keyedScoreValid = sandbox.validateScoreResults(
  {
    results: [
      { id: 's0', score: 9, matchedTopics: ['Singapore AI & Tech Policy'], reason: 'Strong match' },
      { id: 's1', score: 3, matchedTopics: [], reason: 'Unrelated' },
    ],
  },
  keyedScoreMap,
  new Set(['Singapore AI & Tech Policy'])
);
assert.strictEqual(keyedScoreValid.length, 2);
assert.strictEqual(keyedScoreValid[0].score, 9);
assert.strictEqual(keyedScoreValid[1].score, 3);

// 8. Verify partial-completion retry in runBatchWithBisect (when Flash-Lite returns only first 2 of 4 items)
const partialCompletedRows = [];
const fourItems = [{ rowNumber: 2 }, { rowNumber: 3 }, { rowNumber: 4 }, { rowNumber: 5 }];
sandbox.runBatchWithBisect(
  fourItems,
  sub => {
    // Simulate model lazily returning only the first 2 items of whatever batch it receives
    return sub.slice(0, 2).map(item => ({ item, summary: 'Sum ' + item.rowNumber }));
  },
  validRecs => partialCompletedRows.push(...validRecs.map(r => r.item.rowNumber)),
  () => {},
  { deadlineAt: Date.now() + 60000 }
);
assert.deepStrictEqual(partialCompletedRows, [2, 3, 4, 5]);

// 8b. Verify balanced batching with SUMMARY_MAX_ARTICLES_PER_BATCH (50 items @ max 40 -> [25, 25]; 90 items @ max 40 -> [30, 30, 30])
const fiftyItems = Array.from({ length: 50 }, (_, i) => ({ id: i }));
const balanced50 = sandbox.packBatches(fiftyItems, {
  basePromptTokens: 100,
  inBudget: 800000,
  outBudget: 48000,
  inCost: () => 500,
  outPerItem: 220,
  maxItems: 40,
});
assert.deepStrictEqual(
  Array.from(balanced50.map(b => b.length)),
  [25, 25]
);

const ninetyItems = Array.from({ length: 90 }, (_, i) => ({ id: i }));
const balanced90 = sandbox.packBatches(ninetyItems, {
  basePromptTokens: 100,
  inBudget: 800000,
  outBudget: 48000,
  inCost: () => 500,
  outPerItem: 220,
  maxItems: 40,
});
assert.deepStrictEqual(
  Array.from(balanced90.map(b => b.length)),
  [30, 30, 30]
);

// Also verify readExistingUrls picks up recent URLs from Archived Articles even when Articles is empty
const existingUrlsAfterArchive = sandbox.readExistingUrls(artSheet, archSheet, colMap);
assert.ok(existingUrlsAfterArchive.has('https://cna.sg/1'));
assert.ok(existingUrlsAfterArchive.has('https://st.sg/2'));

// 9. Verify pruneArchivedArticles deletes rows older than 7 days and keeps rows <= 7 days old
const nowRef = new Date('2026-11-10T08:00:00+08:00');
const nineDaysAgo = new Date('2026-11-01T07:00:00+08:00'); // > 7 days before Nov 10 -> prune
const threeDaysAgo = new Date('2026-11-07T09:00:00+08:00'); // < 7 days before Nov 10 -> keep
archSheet.getRange(2, 1, 2, 13).setValues([
  [nineDaysAgo, 'CNA', 'Sun, 01 Nov 2026', 'Expired 9d Old', 'https://cna.sg/old', 'Body 1', 'Sum 1', 9, 'AI', 'Match', 'AI Tag', 'Tag Sum', 'YES'],
  [threeDaysAgo, 'Straits Times', 'Sat, 07 Nov 2026', 'Fresh 3d Old', 'https://st.sg/fresh', 'Body 2', 'Sum 2', 8, 'Economy', 'Match', 'Eco Tag', 'Eco Sum', 'YES'],
]);
const prunedArchCount = sandbox.pruneArchivedArticles(archSheet, colMap, 7, nowRef);
assert.strictEqual(prunedArchCount, 1);
assert.strictEqual(archSheet.getLastRow(), 2); // Header + 1 kept row
assert.strictEqual(archSheet.getRange(2, colMap['Title'] + 1).getValues()[0][0], 'Fresh 3d Old');

console.log('ALL SELF-CHECKS PASSED');
