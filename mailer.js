/**
 * HTML news digest builder and MailApp BCC sender (Step 1e).
 * Groups shortlisted articles by their shared `Discussion Tag` so each box in the digest
 * represents a topic/discussion, displays the synthesised `Discussion Summary`, and links
 * to every related article across news agencies.
 */

/**
 * Groups PENDING shortlisted articles (score > threshold) into Topic/Discussion boxes
 * keyed by `Discussion Tag` (falling back to article title), sorted by maxScore descending.
 * @param {Array<Object>} articles
 * @param {number} threshold Strictly greater than this integer (default 7 -> scores 8..10)
 * @return {{ totalCount: number, topicBoxes: Array<Object>, rowNumbers: Array<number> }}
 */
function groupShortlistedArticles(articles, threshold) {
  const minAbove = typeof threshold === 'number' ? threshold : 7;
  const shortlisted = (articles || []).filter(
    a =>
      Number.isInteger(a.score) &&
      a.score > minAbove &&
      String(a.emailed || '').trim().toUpperCase() === 'PENDING'
  );

  const byTagKey = new Map();
  const rowNumbers = [];

  for (const art of shortlisted) {
    rowNumbers.push(art.rowNumber);
    const rawTag = String(art.discussionTag || art.title || 'Untitled Topic').trim();
    const tagKey = rawTag.toLowerCase();
    const safeUrl = /^https:\/\//i.test(String(art.url || '').trim())
      ? String(art.url).trim()
      : '';

    const artMatched = String(art.matchedTopics || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    if (!byTagKey.has(tagKey)) {
      byTagKey.set(tagKey, {
        topicTag: rawTag,
        topicSummary: String(art.discussionSummary || art.summary || '').trim(),
        maxScore: art.score,
        matchedTopicsSet: new Set(artMatched),
        reason: String(art.reason || '').trim(),
        articles: [],
      });
    }

    const box = byTagKey.get(tagKey);
    if (art.score > box.maxScore) {
      box.maxScore = art.score;
      if (art.reason) box.reason = String(art.reason).trim();
    }
    if (!box.topicSummary && (art.discussionSummary || art.summary)) {
      box.topicSummary = String(art.discussionSummary || art.summary).trim();
    }
    for (const t of artMatched) {
      box.matchedTopicsSet.add(t);
    }

    box.articles.push({
      rowNumber: art.rowNumber,
      title: String(art.title || 'Untitled'),
      source: String(art.source || ''),
      publishedAt: String(art.publishedAt || ''),
      url: safeUrl,
      score: art.score,
    });
  }

  const topicBoxes = [];
  for (const box of byTagKey.values()) {
    box.articles.sort((a, b) => b.score - a.score);
    topicBoxes.push({
      topicTag: box.topicTag,
      topicSummary: box.topicSummary,
      maxScore: box.maxScore,
      matchedTopics: Array.from(box.matchedTopicsSet).join(', ') || 'General Highlights',
      reason: box.reason,
      articles: box.articles,
    });
  }

  topicBoxes.sort((a, b) => b.maxScore - a.maxScore || b.articles.length - a.articles.length);

  return {
    totalCount: shortlisted.length,
    topicBoxes,
    rowNumbers,
  };
}

/**
 * Step 1e: Collects PENDING shortlisted articles, renders `email_template.html`
 * (one box per tagged topic/discussion with links to each article), sends a single BCC email
 * to active recipients, and marks rows Emailed = 'YES'.
 */
function sendDigest(ctx) {
  const { articlesSheet, colMap, config, recipients } = ctx;
  if (!recipients || recipients.length === 0) {
    console.warn('Step 1e (Mailer): No active recipients in Recipients tab. Skipping email.');
    return { sent: false, count: 0 };
  }

  const lastRow = articlesSheet.getLastRow();
  if (lastRow < 2) return { sent: false, count: 0 };

  const numRows = lastRow - 1;
  const numCols = Object.keys(colMap).length;
  const eCol = colMap['Emailed'] + 1;
  const scoreCol = colMap['Relevance Score'] + 1;

  const emailedVals = articlesSheet.getRange(2, eCol, numRows, 1).getValues();
  const scoreVals = articlesSheet.getRange(2, scoreCol, numRows, 1).getValues();

  const candidateRowIndices = [];
  for (let i = 0; i < numRows; i++) {
    const status = String(emailedVals[i][0] || '').trim().toUpperCase();
    const score = Number(scoreVals[i][0]);
    if (status === 'PENDING' && Number.isInteger(score) && score > config.SHORTLIST_ABOVE) {
      candidateRowIndices.push(i);
    }
  }

  if (candidateRowIndices.length === 0) {
    console.log('Step 1e (Mailer): 0 PENDING articles with score > ' + config.SHORTLIST_ABOVE + '. Skipping email.');
    return { sent: false, count: 0 };
  }

  const articles = candidateRowIndices.map(i => {
    const rowNumber = 2 + i;
    const row = articlesSheet.getRange(rowNumber, 1, 1, numCols).getValues()[0];
    return {
      rowNumber,
      source: row[colMap['Source']],
      publishedAt: row[colMap['Published At']],
      title: row[colMap['Title']],
      url: row[colMap['URL']],
      summary: row[colMap['Summary']],
      score: Number(row[colMap['Relevance Score']]),
      matchedTopics: row[colMap['Matched Topics']],
      reason: row[colMap['Score Reason']],
      discussionTag: row[colMap['Discussion Tag']],
      discussionSummary: row[colMap['Discussion Summary']],
      emailed: row[colMap['Emailed']],
    };
  });

  const digest = groupShortlistedArticles(articles, config.SHORTLIST_ABOVE);
  if (digest.totalCount === 0) {
    return { sent: false, count: 0 };
  }

  const emailList = recipients.map(r => r.email);
  const remainingQuota = MailApp.getRemainingDailyQuota();
  if (remainingQuota < emailList.length) {
    console.error(
      'Step 1e (Mailer): Insufficient MailApp daily quota (' +
        remainingQuota +
        ' remaining, need ' +
        emailList.length +
        '). Leaving rows PENDING.'
    );
    return { sent: false, count: 0 };
  }

  const dateStr = Utilities.formatDate(new Date(), 'Asia/Singapore', 'dd MMM yyyy');
  const template = HtmlService.createTemplateFromFile('email_template');
  template.dateStr = dateStr;
  template.totalCount = digest.totalCount;
  template.threshold = config.SHORTLIST_ABOVE;
  template.topicBoxes = digest.topicBoxes;

  const htmlBody = template.evaluate().getContent();
  const subject =
    'News Digest (' +
    dateStr +
    ') — ' +
    digest.topicBoxes.length +
    ' Topic(s), ' +
    digest.totalCount +
    ' Article(s)';

  let toAddress = emailList[0];
  let bccAddresses = emailList.slice(1).join(',');
  try {
    const ownerEmail = Session.getEffectiveUser().getEmail();
    if (ownerEmail && ownerEmail.includes('@')) {
      toAddress = ownerEmail;
      bccAddresses = emailList.join(',');
    }
  } catch (_) {}

  const mailOpts = {
    to: toAddress,
    subject,
    htmlBody,
    name: 'GSheet News Digest',
  };
  if (bccAddresses) {
    mailOpts.bcc = bccAddresses;
  }

  MailApp.sendEmail(mailOpts);

  for (const r of digest.rowNumbers) {
    articlesSheet.getRange(r, eCol).setValue('YES');
  }

  console.log(
    'Step 1e (Mailer): Sent digest with ' +
      digest.topicBoxes.length +
      ' topic box(es) (' +
      digest.totalCount +
      ' articles) to ' +
      emailList.length +
      ' recipient(s).'
  );
  return { sent: true, count: digest.totalCount, topicsCount: digest.topicBoxes.length };
}
