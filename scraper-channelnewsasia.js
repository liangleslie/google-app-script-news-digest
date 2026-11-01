/**
 * Channel NewsAsia (CNA) scraper module (`scraper-channelnewsasia.js`).
 * Feed URLs (`CNA_RSS_FEEDS`) are defined in `constants.js`.
 */

/**
 * Scrapes new Channel NewsAsia articles published after `lastRunMs`,
 * pulling evenly across all configured CNA RSS feeds (Latest, Singapore, Asia, Business).
 * @param {Set<string>} existingUrls Normalised URLs already stored in Articles / Archived Articles
 * @param {number} maxItems Maximum new articles to return for this site
 * @param {number} maxChars Maximum article body length
 * @param {number} [lastRunMs] Epoch ms of `LAST_RUN_AT` — only items published after this are picked
 * @return {Array<{source: string, publishedAt: string, title: string, url: string, content: string}>}
 */
function scrapeChannelNewsAsia(existingUrls, maxItems, maxChars, lastRunMs) {
  const limit = maxItems || 40;
  const charCap = maxChars || 12000;

  const feedResponses = fetchAllSafe(CNA_RSS_FEEDS);
  const candidates = selectRssItemsRoundRobin(
    feedResponses,
    CNA_RSS_FEEDS,
    'CNA',
    existingUrls,
    limit,
    lastRunMs
  );

  if (candidates.length === 0) return [];

  const pageResponses = fetchAllSafe(candidates.map(c => c.url));
  return candidates.map((item, idx) => {
    let body = '';
    const res = pageResponses[idx];
    if (res && res.getResponseCode() === 200) {
      body = extractCnaBody_(res.getContentText(), charCap);
    }
    const content = body && body.length >= 80 ? body : (item.rssDescription || item.title);
    return {
      source: 'CNA',
      publishedAt: item.publishedAt,
      title: item.title,
      url: item.url,
      content: content.slice(0, charCap),
    };
  });
}

/**
 * CNA-specific body extractor: tries shared JSON-LD / <article> extractor first,
 * with explicit support for CNA's `.text-long` content blocks.
 */
function extractCnaBody_(html, maxChars) {
  const generic = extractArticleBody(html, maxChars);
  if (generic && generic.length >= 120) return generic;

  const textLongRegex = /<div[^>]+class=["'][^"']*text-long[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi;
  const blocks = [];
  let match;
  while ((match = textLongRegex.exec(html)) !== null) {
    const cleaned = decodeEntities(stripHtml(match[1])).replace(/\s+/g, ' ').trim();
    if (cleaned.length >= 40) blocks.push(cleaned);
  }
  return blocks.join('\n\n').slice(0, maxChars);
}
