function normalizeTweetText(text) {
  return text.replace(/\s+/g, ' ').trim();
}

function extractTweetsFromPage() {
  const tweets = [];
  const articles = document.querySelectorAll('article[data-testid="tweet"]');

  for (const article of articles) {
    const textNode = article.querySelector('[data-testid="tweetText"]');
    if (!textNode) continue;

    const raw = textNode.innerText || textNode.textContent || '';
    const text = normalizeTweetText(raw);
    if (!text) continue;

    const timeEl = article.querySelector('time');
    const datetime = timeEl?.getAttribute('datetime') || null;

    tweets.push({ text, datetime });
  }

  const unique = [];
  const seen = new Set();
  for (const tweet of tweets) {
    const key = `${tweet.datetime || ''}-${tweet.text}`;
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(tweet);
    }
  }

  return unique;
}

chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request?.type !== 'EXTRACT_TWEETS') return;
  try {
    const tweets = extractTweetsFromPage();
    sendResponse({ ok: true, tweets });
  } catch (error) {
    sendResponse({ ok: false, error: String(error) });
  }
});
