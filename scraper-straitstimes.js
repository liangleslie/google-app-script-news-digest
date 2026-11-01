/**
 * The Straits Times scraper module (`scraper-straitstimes.js`).
 * Feed URLs (`ST_RSS_FEEDS`) are defined in `constants.js`.
 */

/**
 * Scrapes new Straits Times articles published after `lastRunMs`,
 * pulling evenly across all configured Straits Times RSS feeds (Singapore, Asia, Business, World).
 * @param {Set<string>} existingUrls Normalised URLs already stored in Articles / Archived Articles
 * @param {number} maxItems Maximum new articles to return for this site
 * @param {number} maxChars Maximum article body length
 * @param {number} [lastRunMs] Epoch ms of `LAST_RUN_AT` — only items published after this are picked
 * @return {Array<{source: string, publishedAt: string, title: string, url: string, content: string}>}
 */
function scrapeStraitsTimes(existingUrls, maxItems, maxChars, lastRunMs) {
  const limit = maxItems || 40;
  const charCap = maxChars || 12000;

  const feedResponses = fetchAllSafe(ST_RSS_FEEDS);
  const candidates = selectRssItemsRoundRobin(
    feedResponses,
    ST_RSS_FEEDS,
    'Straits Times',
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
      body = extractArticleBody(res.getContentText(), charCap);
    }
    const content = body && body.length >= 80 ? body : (item.rssDescription || item.title);
    return {
      source: 'Straits Times',
      publishedAt: item.publishedAt,
      title: item.title,
      url: item.url,
      content: content.slice(0, charCap),
    };
  });
}
