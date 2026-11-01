/**
 * Google AI Studio (Gemini REST API v1beta) batch summariser, cross-agency topic
 * cluster/synthesiser, and relevance scorer.
 * Uses dynamic ID-keyed `required` object schemas for summarisation and scoring so
 * Gemini's constrained decoder cannot stop early or omit items in a batch.
 */

const CLUSTER_RESPONSE_SCHEMA = {
  type: 'OBJECT',
  required: ['clusters'],
  properties: {
    clusters: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        required: ['topicTag', 'topicSummary', 'articleIds'],
        propertyOrdering: ['topicTag', 'topicSummary', 'articleIds'],
        properties: {
          topicTag: { type: 'STRING' },
          topicSummary: { type: 'STRING' },
          articleIds: {
            type: 'ARRAY',
            items: { type: 'STRING' },
          },
        },
      },
    },
  },
};

/**
 * Dynamically builds an OBJECT schema where every article ID (`a0`, `a1`, ..., `aN`)
 * in the batch is an explicit `required` property in `propertyOrdering`.
 * Forces Gemini's constrained decoder to generate a summary for every single ID before closing `}`.
 */
function buildSummaryResponseSchema(ids) {
  const properties = {};
  for (const id of ids) {
    properties[id] = { type: 'STRING' };
  }
  return {
    type: 'OBJECT',
    required: ids,
    propertyOrdering: ids,
    properties,
  };
}

/**
 * Builds the strict JSON response schema for relevance scoring.
 * Uses a single fixed-depth `results` array schema (with `enum` on `matchedTopics`) so
 * all 50-200+ topic clusters can be scored together in ONE single holistic API call
 * without hitting Gemini's ~64-property object schema limit.
 */
function buildScoreResponseSchema(activeTopicNames) {
  const topicItemsSchema =
    activeTopicNames && activeTopicNames.length > 0
      ? { type: 'STRING', enum: activeTopicNames }
      : { type: 'STRING' };

  return {
    type: 'OBJECT',
    required: ['results'],
    properties: {
      results: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          required: ['id', 'score', 'matchedTopics', 'reason'],
          propertyOrdering: ['id', 'score', 'matchedTopics', 'reason'],
          properties: {
            id: { type: 'STRING' },
            score: { type: 'INTEGER' },
            matchedTopics: {
              type: 'ARRAY',
              items: topicItemsSchema,
            },
            reason: { type: 'STRING' },
          },
        },
      },
    },
  };
}

/**
 * Retrieves the Gemini API key strictly from Script Property `GEMINI_API`.
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API');
  if (!key || !String(key).trim()) {
    throw new Error(
      'Missing Script Property "GEMINI_API". Set it in Apps Script -> Project Settings -> Script Properties.'
    );
  }
  return String(key).trim();
}

/**
 * Fetches model inputTokenLimit and outputTokenLimit from `GET v1beta/models/{model}`,
 * caching the result in CacheService for 6 hours.
 */
function getModelLimits(model, apiKey) {
  const cleanModel = String(model || 'gemini-3.8-flash').replace(/^models\//, '').trim();
  const cacheKey = 'gemini_limits_' + cleanModel;
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (_) {}

  const url = GEMINI_API_BASE + encodeURIComponent(cleanModel);
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'x-goog-api-key': apiKey },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const bodyText = res.getContentText();
  if (code !== 200) {
    throw new Error(
      'Gemini models.get failed for model "' +
        cleanModel +
        '" (HTTP ' +
        code +
        '): ' +
        bodyText.slice(0, 300)
    );
  }

  const info = JSON.parse(bodyText);
  const limits = {
    inputTokenLimit: Number(info.inputTokenLimit) || 1048576,
    outputTokenLimit: Number(info.outputTokenLimit) || 65536,
  };

  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(limits), 21600);
  } catch (_) {}

  console.log(
    'Model "' +
      cleanModel +
      '" limits: input=' +
      limits.inputTokenLimit +
      ', output=' +
      limits.outputTokenLimit
  );
  return limits;
}

/**
 * Conservative token estimator (~3 characters per token, ~33% safety margin over English's ~4 chars/token).
 * ponytail: chars/3 heuristic avoids extra network round-trips. Ceiling: under-estimates dense CJK text.
 * Upgrade path: call POST v1beta/models/{m}:countTokens when exact counts are required.
 */
function estimateTokens(text) {
  return Math.ceil(String(text || '').length / 3);
}

/**
 * Greedy token-budget packer: adds items to the current batch while input token budget,
 * output token budget, and optional `maxItems` cap all hold. Always allows >= 1 item per batch.
 */
function packBatches(items, opts) {
  const { basePromptTokens, inBudget, outBudget, inCost, outPerItem, maxItems } = opts;
  const batches = [];
  let cur = [];
  let inSum = basePromptTokens;
  let outSum = 0;

  for (const it of items) {
    const c = inCost(it);
    if (
      cur.length > 0 &&
      (inSum + c > inBudget ||
        outSum + outPerItem > outBudget ||
        (maxItems && cur.length >= maxItems))
    ) {
      batches.push(cur);
      cur = [];
      inSum = basePromptTokens;
      outSum = 0;
    }
    cur.push(it);
    inSum += c;
    outSum += outPerItem;
  }
  if (cur.length > 0) {
    batches.push(cur);
  }
  return batches;
}

/**
 * Computes effective input, output, and maxOutputTokens budgets from model limits (`models.get`),
 * applying optional manual caps only when `Config` specifies an integer instead of `AUTO`.
 */
function computeTokenBudgets(config, modelLimits) {
  const dynamicInLimit = Math.floor((modelLimits.inputTokenLimit || 1048576) * 0.8);
  const dynamicOutLimit = modelLimits.outputTokenLimit || 65536;

  const inBudget = config.MAX_INPUT_TOKENS_PER_CALL
    ? Math.min(config.MAX_INPUT_TOKENS_PER_CALL, dynamicInLimit)
    : dynamicInLimit;

  const maxOut = config.MAX_OUTPUT_TOKENS_PER_CALL
    ? Math.min(config.MAX_OUTPUT_TOKENS_PER_CALL, dynamicOutLimit)
    : dynamicOutLimit;

  // Reserve 25% of maxOutputTokens for internal thinking tokens so JSON is not truncated
  const outBudget = Math.floor(maxOut * 0.75);
  return { inBudget, outBudget, maxOut };
}

/**
 * Interpolates `{{KEY}}` placeholders in a prompt template from `constants.js` / `Config` tab.
 */
function renderPromptTemplate(template, vars) {
  return String(template || '').replace(/\{\{(\w+)\}\}/g, (_, k) =>
    vars && vars[k] !== undefined ? String(vars[k]) : ''
  );
}

/**
 * Step 1c: Summarises unsummarised rows in the Articles sheet in token-budgeted batches,
 * using a dynamic `required: ['a0', ..., 'aN']` object schema per batch so every article is guaranteed a summary.
 */
function summarizePending(ctx) {
  const { articlesSheet, colMap, config, apiKey, modelLimits, deadlineAt } = ctx;
  const lastRow = articlesSheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const titleCol = colMap['Title'] + 1;
  const sourceCol = colMap['Source'] + 1;
  const contentCol = colMap['Content'] + 1;
  const summaryCol = colMap['Summary'] + 1;

  const summaries = articlesSheet.getRange(2, summaryCol, numRows, 1).getValues();
  const titles = articlesSheet.getRange(2, titleCol, numRows, 1).getValues();
  const sources = articlesSheet.getRange(2, sourceCol, numRows, 1).getValues();

  const pendingIndices = [];
  for (let i = 0; i < numRows; i++) {
    const existingSummary = String(summaries[i][0] || '').trim();
    if (!existingSummary) {
      pendingIndices.push(i);
    }
  }
  if (pendingIndices.length === 0) {
    console.log('Step 1c (Summarise): 0 pending articles.');
    return;
  }

  const firstIdx = pendingIndices[0];
  const lastIdx = pendingIndices[pendingIndices.length - 1];
  const spanCount = lastIdx - firstIdx + 1;
  const contentSlice = articlesSheet
    .getRange(2 + firstIdx, contentCol, spanCount, 1)
    .getValues();

  const pendingItems = [];
  for (const idx of pendingIndices) {
    const rawContent = String(contentSlice[idx - firstIdx][0] || '').trim();
    const title = String(titles[idx][0] || '').trim();
    const source = String(sources[idx][0] || '').trim();
    if (!rawContent && !title) continue;
    pendingItems.push({
      rowNumber: 2 + idx,
      title,
      source,
      content: (rawContent || title).slice(0, config.MAX_ARTICLE_CHARS),
    });
  }

  if (pendingItems.length === 0) return;

  const systemInstruction = config.PROMPT_SUMMARIZE || DEFAULT_PROMPTS.PROMPT_SUMMARIZE;
  const budgets = computeTokenBudgets(config, modelLimits);
  const basePromptTokens = estimateTokens(systemInstruction) + 150;

  const batches = packBatches(pendingItems, {
    basePromptTokens,
    inBudget: budgets.inBudget,
    outBudget: budgets.outBudget,
    inCost: it => estimateTokens(it.title) + estimateTokens(it.content) + 30,
    outPerItem: EST_SUMMARY_OUT_TOKENS,
    maxItems: config.SUMMARY_MAX_ARTICLES_PER_BATCH,
  });

  console.log(
    'Step 1c (Summarise): ' +
      pendingItems.length +
      ' pending articles packed into ' +
      batches.length +
      ' token-budgeted batch(es).'
  );

  for (const batch of batches) {
    if (Date.now() >= deadlineAt) {
      ctx.deadlineTripped = true;
      console.warn('Deadline reached during Step 1c (Summarise); remaining batches deferred.');
      break;
    }

    runBatchWithBisect(
      batch,
      subBatch => {
        const idToItem = new Map();
        const ids = [];
        const payload = subBatch.map((it, i) => {
          const id = 'a' + i;
          ids.push(id);
          idToItem.set(id, it);
          return { id, source: it.source, title: it.title, content: it.content };
        });

        const parsed = callGeminiJson({
          model: config.GEMINI_MODEL,
          apiKey,
          systemInstruction,
          userText: JSON.stringify({ articles: payload }),
          responseSchema: buildSummaryResponseSchema(ids),
          maxOutputTokens: budgets.maxOut,
          thinkingLevel: config.THINKING_LEVEL,
          deadlineAt,
        });

        return validateSummaryResults(parsed, idToItem);
      },
      validRecords => {
        for (const rec of validRecords) {
          articlesSheet.getRange(rec.item.rowNumber, summaryCol).setValue(rec.summary);
        }
      },
      (poisonItem, errMsg) => {
        const cleanReason = String(errMsg || 'Summarisation failed').slice(0, 200);
        articlesSheet
          .getRange(poisonItem.rowNumber, summaryCol)
          .setValue('[ERROR] ' + cleanReason);
      },
      ctx
    );
  }
}

/**
 * Step 1d:
 * 1) Groups similar articles on the same topic/discussion across news agencies into
 *    `Discussion Tag` (Col K) and generates a synthesised `Discussion Summary` (Col L)
 *    for ALL summarised articles (`clusterPendingArticles`).
 * 2) Scores each grouped discussion/article against active `Topics` in token-budgeted batches
 *    using a dynamic ID-keyed `required` schema (`buildScoreResponseSchema`),
 *    propagating the topic score across all member articles in the same discussion group.
 */
function scorePending(ctx) {
  const { articlesSheet, colMap, config, topics, apiKey, modelLimits, deadlineAt } = ctx;
  if (!topics || topics.length === 0) {
    console.warn('Step 1d (Score): No active topics configured in Topics tab. Skipping scoring.');
    return;
  }

  const lastRow = articlesSheet.getLastRow();
  if (lastRow < 2) return;

  // Sub-step 1d-i: Group & synthesise all summarised articles by topic/discussion FIRST
  clusterPendingArticles(ctx);

  if (Date.now() >= deadlineAt) {
    ctx.deadlineTripped = true;
    return;
  }

  // Sub-step 1d-ii: Score each topic cluster against active Topics
  const numRows = lastRow - 1;
  const titleCol = colMap['Title'] + 1;
  const sourceCol = colMap['Source'] + 1;
  const summaryCol = colMap['Summary'] + 1;
  const scoreCol = colMap['Relevance Score'] + 1;
  const tagCol = colMap['Discussion Tag'] + 1;
  const tagSumCol = colMap['Discussion Summary'] + 1;

  const titles = articlesSheet.getRange(2, titleCol, numRows, 1).getValues();
  const sources = articlesSheet.getRange(2, sourceCol, numRows, 1).getValues();
  const summaries = articlesSheet.getRange(2, summaryCol, numRows, 1).getValues();
  const scores = articlesSheet.getRange(2, scoreCol, numRows, 1).getValues();
  const tags = articlesSheet.getRange(2, tagCol, numRows, 1).getValues();
  const tagSums = articlesSheet.getRange(2, tagSumCol, numRows, 1).getValues();

  const groupsByTag = new Map();
  for (let i = 0; i < numRows; i++) {
    const summaryText = String(summaries[i][0] || '').trim();
    if (!summaryText || summaryText.startsWith('[ERROR]')) continue;

    const rowNumber = 2 + i;
    const title = String(titles[i][0] || '').trim();
    const source = String(sources[i][0] || '').trim();
    const tag = String(tags[i][0] || title).trim();
    const tagSum = String(tagSums[i][0] || summaryText).trim();
    const scoreVal = String(scores[i][0] ?? '').trim();

    const tagKey = tag.toLowerCase();
    if (!groupsByTag.has(tagKey)) {
      groupsByTag.set(tagKey, {
        topicTag: tag,
        topicSummary: tagSum,
        sources: new Set(),
        rowNumbers: [],
        needsScoring: false,
      });
    }
    const g = groupsByTag.get(tagKey);
    g.sources.add(source);
    g.rowNumbers.push(rowNumber);
    if (scoreVal === '') {
      g.needsScoring = true;
    }
  }

  const pendingClusters = Array.from(groupsByTag.values()).filter(g => g.needsScoring);
  if (pendingClusters.length === 0) {
    console.log('Step 1d (Score): 0 pending topic clusters to score.');
    return;
  }

  const activeTopicNames = topics.map(t => t.name);
  const systemInstruction = renderPromptTemplate(
    config.PROMPT_SCORE || DEFAULT_PROMPTS.PROMPT_SCORE,
    {
      SHORTLIST_ABOVE: config.SHORTLIST_ABOVE,
      TOPICS_JSON: JSON.stringify(topics),
    }
  );

  const budgets = computeTokenBudgets(config, modelLimits);

  console.log(
    'Step 1d (Score): scoring all ' +
      pendingClusters.length +
      ' topic cluster(s) together in 1 holistic call.'
  );

  // Score all topic clusters in a single holistic call so relative scoring sees the entire set
  runBatchWithBisect(
    pendingClusters,
    subBatch => {
      const idToCluster = new Map();
      const requiredIds = [];
      const payload = subBatch.map((g, i) => {
        const id = 's' + i;
        requiredIds.push(id);
        idToCluster.set(id, g);
        return {
          id,
          source: Array.from(g.sources).join(', '),
          title: g.topicTag,
          summary: g.topicSummary,
        };
      });

      const parsed = callGeminiJson({
        model: config.GEMINI_MODEL,
        apiKey,
        systemInstruction,
        userText: JSON.stringify({
          totalClusters: payload.length,
          requiredClusterIds: requiredIds,
          articles: payload,
        }),
        responseSchema: buildScoreResponseSchema(activeTopicNames),
        maxOutputTokens: budgets.maxOut,
        thinkingLevel: config.THINKING_LEVEL,
        deadlineAt,
      });

      return validateScoreResults(parsed, idToCluster, new Set(activeTopicNames));
    },
    validRecords => {
      for (const rec of validRecords) {
        const emailedStatus = rec.score > config.SHORTLIST_ABOVE ? 'PENDING' : 'SKIPPED';
        const matchedStr = rec.matchedTopics.join(', ');
        for (const r of rec.item.rowNumbers) {
          articlesSheet.getRange(r, colMap['Relevance Score'] + 1).setValue(rec.score);
          articlesSheet.getRange(r, colMap['Matched Topics'] + 1).setValue(matchedStr);
          articlesSheet.getRange(r, colMap['Score Reason'] + 1).setValue(rec.reason);
          articlesSheet.getRange(r, colMap['Emailed'] + 1).setValue(emailedStatus);
        }
      }
    },
    (poisonCluster, errMsg) => {
      const cleanErr = '[ERROR] ' + String(errMsg || 'Scoring failed').slice(0, 200);
      for (const r of poisonCluster.rowNumbers) {
        articlesSheet.getRange(r, colMap['Relevance Score'] + 1).setValue(0);
        articlesSheet.getRange(r, colMap['Score Reason'] + 1).setValue(cleanErr);
        articlesSheet.getRange(r, colMap['Emailed'] + 1).setValue('SKIPPED');
      }
    },
    ctx
  );
}

/**
 * Groups summarised articles in `Articles` across news agencies that cover the same
 * topic/discussion, tags them with a shared `Discussion Tag` (Col K), and writes a synthesised
 * `Discussion Summary` (Col L) generated from all related articles in the group.
 */
function clusterPendingArticles(ctx) {
  const { articlesSheet, colMap, config, apiKey, modelLimits, deadlineAt } = ctx;
  const lastRow = articlesSheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const titleCol = colMap['Title'] + 1;
  const sourceCol = colMap['Source'] + 1;
  const summaryCol = colMap['Summary'] + 1;
  const tagCol = colMap['Discussion Tag'] + 1;
  const tagSumCol = colMap['Discussion Summary'] + 1;
  const emailedCol = colMap['Emailed'] + 1;

  const titles = articlesSheet.getRange(2, titleCol, numRows, 1).getValues();
  const sources = articlesSheet.getRange(2, sourceCol, numRows, 1).getValues();
  const summaries = articlesSheet.getRange(2, summaryCol, numRows, 1).getValues();
  const tags = articlesSheet.getRange(2, tagCol, numRows, 1).getValues();
  const tagSums = articlesSheet.getRange(2, tagSumCol, numRows, 1).getValues();
  const emailedVals = articlesSheet.getRange(2, emailedCol, numRows, 1).getValues();

  const candidateItems = [];
  let hasUnclustered = false;

  for (let i = 0; i < numRows; i++) {
    const summaryText = String(summaries[i][0] || '').trim();
    if (!summaryText || summaryText.startsWith('[ERROR]')) continue;

    const emailed = String(emailedVals[i][0] || '').trim().toUpperCase();
    if (emailed === 'YES') continue;

    const existingTag = String(tags[i][0] || '').trim();
    const existingTagSum = String(tagSums[i][0] || '').trim();
    if (!existingTag || !existingTagSum) {
      hasUnclustered = true;
    }

    candidateItems.push({
      rowNumber: 2 + i,
      source: String(sources[i][0] || '').trim(),
      title: String(titles[i][0] || '').trim(),
      summary: summaryText,
    });
  }

  if (candidateItems.length === 0 || !hasUnclustered) {
    return;
  }

  const systemInstruction = config.PROMPT_CLUSTER || DEFAULT_PROMPTS.PROMPT_CLUSTER;
  const budgets = computeTokenBudgets(config, modelLimits);

  console.log(
    'Step 1d (Cluster & Synthesise): grouping all ' +
      candidateItems.length +
      ' article(s) across agencies together in 1 holistic call.'
  );

  // Cluster all candidate summaries in a single holistic call so cross-agency matches are never split across batches
  runBatchWithBisect(
    candidateItems,
    subBatch => {
      const idToItem = new Map();
      const requiredIds = [];
      const payload = subBatch.map((it, i) => {
        const id = 'c' + i;
        requiredIds.push(id);
        idToItem.set(id, it);
        return {
          id,
          source: it.source,
          title: it.title,
          summary: it.summary,
        };
      });

      const parsed = callGeminiJson({
        model: config.GEMINI_MODEL,
        apiKey,
        systemInstruction,
        userText: JSON.stringify({
          totalArticles: payload.length,
          requiredArticleIds: requiredIds,
          articles: payload,
        }),
        responseSchema: CLUSTER_RESPONSE_SCHEMA,
        maxOutputTokens: budgets.maxOut,
        thinkingLevel: config.THINKING_LEVEL,
        deadlineAt,
      });

      return validateClusterResults(parsed, idToItem);
    },
    validAssignments => {
      for (const assign of validAssignments) {
        const r = assign.item.rowNumber;
        articlesSheet.getRange(r, tagCol).setValue(assign.topicTag);
        articlesSheet.getRange(r, tagSumCol).setValue(assign.topicSummary);
      }
    },
    poisonItem => {
      const r = poisonItem.rowNumber;
      articlesSheet.getRange(r, tagCol).setValue(poisonItem.title);
      articlesSheet.getRange(r, tagSumCol).setValue(poisonItem.summary);
    },
    ctx
  );
}

/**
 * Validates Gemini summary output: supports the dynamic ID-keyed object `{ a0: "summary", ... }`
 * (as well as legacy `{ results: [{ id, summary }] }`).
 */
function validateSummaryResults(parsed, idToItem) {
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Invalid summary response: not an object');
  }
  const seen = new Set();
  const valid = [];

  if (Array.isArray(parsed.results)) {
    for (const entry of parsed.results) {
      if (!entry || typeof entry !== 'object') continue;
      const id = String(entry.id || '');
      const summary = typeof entry.summary === 'string' ? entry.summary.trim() : '';
      if (!idToItem.has(id) || seen.has(id) || !summary) continue;
      seen.add(id);
      valid.push({ item: idToItem.get(id), summary });
    }
  } else {
    for (const [id, item] of idToItem.entries()) {
      const val = parsed[id];
      const summary = typeof val === 'string' ? val.trim() : '';
      if (summary) {
        seen.add(id);
        valid.push({ item, summary });
      }
    }
  }

  if (valid.length === 0 && idToItem.size > 0) {
    throw new Error('Zero valid summary items returned for batch of ' + idToItem.size);
  }
  return valid;
}

/**
 * Validates Gemini score output: supports the dynamic ID-keyed object `{ s0: { score, matchedTopics, reason }, ... }`
 * (as well as legacy `{ results: [{ id, score, matchedTopics, reason }] }`).
 */
function validateScoreResults(parsed, idToItem, allowedTopicsSet) {
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Invalid score response: not an object');
  }
  const seen = new Set();
  const valid = [];

  const entries = Array.isArray(parsed.results)
    ? parsed.results
    : Array.from(idToItem.keys()).map(id => Object.assign({ id }, parsed[id] || {}));

  for (const entry of entries) {
    if (!entry || typeof entry !== 'object') continue;
    const id = String(entry.id || '');
    const score = entry.score;
    const reason = typeof entry.reason === 'string' ? entry.reason.trim() : '';
    if (!idToItem.has(id) || seen.has(id)) continue;
    if (!Number.isInteger(score) || score < 1 || score > 10) continue;
    if (!reason) continue;

    const rawTopics = Array.isArray(entry.matchedTopics) ? entry.matchedTopics : [];
    const matchedTopics = rawTopics
      .map(t => String(t || '').trim())
      .filter(t => t && (!allowedTopicsSet || allowedTopicsSet.size === 0 || allowedTopicsSet.has(t)));

    seen.add(id);
    valid.push({
      item: idToItem.get(id),
      score,
      matchedTopics,
      reason,
    });
  }

  if (valid.length === 0 && idToItem.size > 0) {
    throw new Error('Zero valid score items returned for batch of ' + idToItem.size);
  }
  return valid;
}

/**
 * Validates Gemini cluster output: ensures every assigned articleId belongs to the batch
 * and prevents duplicate assignments. Omitted items in multi-item batches are automatically
 * retried by `runBatchWithBisect`.
 */
function validateClusterResults(parsed, idToItem) {
  if (!parsed || !Array.isArray(parsed.clusters)) {
    throw new Error('Invalid cluster schema: missing clusters array');
  }
  const assignedIds = new Set();
  const assignments = [];

  for (const cluster of parsed.clusters) {
    if (!cluster || typeof cluster !== 'object') continue;
    const topicTag = typeof cluster.topicTag === 'string' ? cluster.topicTag.trim() : '';
    const topicSummary = typeof cluster.topicSummary === 'string' ? cluster.topicSummary.trim() : '';
    const articleIds = Array.isArray(cluster.articleIds) ? cluster.articleIds : [];
    if (!topicTag || !topicSummary || articleIds.length === 0) continue;

    for (const rawId of articleIds) {
      const id = String(rawId || '').trim();
      if (!idToItem.has(id) || assignedIds.has(id)) continue;
      assignedIds.add(id);
      assignments.push({
        item: idToItem.get(id),
        topicTag,
        topicSummary,
      });
    }
  }

  if (assignments.length === 0 && idToItem.size > 0) {
    throw new Error('Zero valid cluster assignments returned for batch of ' + idToItem.size);
  }

  if (idToItem.size === 1 && assignments.length === 0) {
    const item = Array.from(idToItem.values())[0];
    assignments.push({
      item,
      topicTag: item.title,
      topicSummary: item.summary,
    });
  }

  return assignments;
}

/**
 * Executes a batch call; if the model returns only a partial subset of items,
 * saves the returned items immediately and retries the remaining omitted items.
 * If a call throws or returns 0 valid items, recursively bisects the batch into halves.
 * Marks a single failing item as poison so it won't block future runs.
 */
function runBatchWithBisect(batch, callAndValidateFn, onValidItemsFn, onPoisonItemFn, ctx) {
  if (!batch || batch.length === 0) return;
  if (ctx && Date.now() >= ctx.deadlineAt) {
    ctx.deadlineTripped = true;
    return;
  }

  try {
    const validRecords = callAndValidateFn(batch);
    onValidItemsFn(validRecords);

    if (validRecords.length < batch.length) {
      const completedItems = new Set(validRecords.map(r => r.item));
      const missingBatch = batch.filter(it => !completedItems.has(it));
      if (missingBatch.length > 0 && missingBatch.length < batch.length) {
        console.warn(
          'Model returned ' +
            validRecords.length +
            '/' +
            batch.length +
            ' items; retrying remaining ' +
            missingBatch.length +
            ' omitted item(s).'
        );
        runBatchWithBisect(missingBatch, callAndValidateFn, onValidItemsFn, onPoisonItemFn, ctx);
      }
    }
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    console.warn('Batch of ' + batch.length + ' failed (' + msg + ').');
    if (batch.length === 1) {
      console.error('Handling single failing item.');
      onPoisonItemFn(batch[0], msg);
      return;
    }
    const mid = Math.ceil(batch.length / 2);
    const left = batch.slice(0, mid);
    const right = batch.slice(mid);
    runBatchWithBisect(left, callAndValidateFn, onValidItemsFn, onPoisonItemFn, ctx);
    runBatchWithBisect(right, callAndValidateFn, onValidItemsFn, onPoisonItemFn, ctx);
  }
}

/**
 * Calls Google AI Studio `v1beta/models/{model}:generateContent` with enforced JSON schema,
 * `x-goog-api-key` header, rate-limit backoff, `finishReason === 'STOP'` verification,
 * and automatic fallback if `thinkingLevel` is rejected by a 2.5-series model.
 */
function callGeminiJson(opts) {
  const {
    model,
    apiKey,
    systemInstruction,
    userText,
    responseSchema,
    maxOutputTokens,
    thinkingLevel,
    deadlineAt,
  } = opts;

  const cleanModel = String(model || 'gemini-3.8-flash').replace(/^models\//, '').trim();
  const url = GEMINI_API_BASE + encodeURIComponent(cleanModel) + ':generateContent';

  const generationConfig = {
    responseMimeType: 'application/json',
    responseSchema,
    maxOutputTokens: maxOutputTokens || 16384,
    temperature: 0.2,
  };

  if (thinkingLevel && String(thinkingLevel).trim()) {
    generationConfig.thinkingConfig = {
      thinkingLevel: String(thinkingLevel).trim(),
    };
  }

  const requestBody = {
    systemInstruction: {
      parts: [{ text: systemInstruction }],
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: userText }],
      },
    ],
    generationConfig,
  };

  const estIn = estimateTokens(systemInstruction) + estimateTokens(userText);

  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const payloadStr = JSON.stringify(requestBody);
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: payloadStr,
      muteHttpExceptions: true,
    });

    const status = res.getResponseCode();
    const resText = res.getContentText();

    if (status === 400 && generationConfig.thinkingConfig && /thinking/i.test(resText)) {
      console.warn('Model "' + cleanModel + '" rejected thinkingLevel; retrying without thinkingConfig.');
      delete generationConfig.thinkingConfig;
      continue;
    }

    if (status === 429 || status === 503) {
      const waitMs = parseRetryDelayMs_(resText, attempt);
      if (attempt < maxAttempts && (!deadlineAt || Date.now() + waitMs + 5000 < deadlineAt)) {
        console.warn('Gemini HTTP ' + status + ' (attempt ' + attempt + '); sleeping ' + waitMs + 'ms.');
        Utilities.sleep(waitMs);
        continue;
      }
      throw new Error('Gemini rate limit / overload HTTP ' + status + ': ' + resText.slice(0, 250));
    }

    if (status !== 200) {
      throw new Error('Gemini API HTTP ' + status + ': ' + resText.slice(0, 300));
    }

    const data = JSON.parse(resText);
    const usage = data.usageMetadata || {};
    console.log(
      'Gemini usage [' +
        cleanModel +
        ']: estIn=' +
        estIn +
        ', actualIn=' +
        (usage.promptTokenCount || 0) +
        ', out=' +
        (usage.candidatesTokenCount || 0) +
        ', thoughts=' +
        (usage.thoughtsTokenCount || 0)
    );

    const candidate = data.candidates && data.candidates[0];
    if (!candidate) {
      throw new Error('Gemini response contained no candidates: ' + resText.slice(0, 250));
    }

    const finishReason = candidate.finishReason;
    if (finishReason && finishReason !== 'STOP') {
      throw new Error('Gemini generation did not finish cleanly (finishReason=' + finishReason + ')');
    }

    const parts = (candidate.content && candidate.content.parts) || [];
    const textPart = parts.find(p => typeof p.text === 'string' && !p.thought);
    const rawJson = textPart ? textPart.text : (parts[0] && parts[0].text) || '';
    if (!rawJson) {
      throw new Error('Gemini returned empty text part');
    }

    return JSON.parse(rawJson);
  }

  throw new Error('Exhausted Gemini retry attempts');
}

/**
 * Extracts RetryInfo.retryDelay (e.g. "12s") from a 429 error payload, defaulting to exponential backoff.
 */
function parseRetryDelayMs_(resText, attempt) {
  try {
    const match = /"retryDelay"\s*:\s*"(\d+(?:\.\d+)?)s"/i.exec(resText);
    if (match) {
      return Math.min(30000, Math.ceil(parseFloat(match[1]) * 1000) + 500);
    }
  } catch (_) {}
  return Math.min(20000, 2000 * Math.pow(2, attempt - 1));
}
