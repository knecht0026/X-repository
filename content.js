const SETTINGS_DEFAULT = {
  autoDownload: false,
  minLikes: 100
};

const processedTweetNodes = new WeakSet();
let settingsCache = { ...SETTINGS_DEFAULT };
let scanTimer = null;

init();

async function init() {
  settingsCache = { ...SETTINGS_DEFAULT, ...(await chrome.storage.local.get(['autoDownload', 'minLikes'])) };
  installDownloadButtons();
  startScanner();
}

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== 'local') return;
  if (changes.autoDownload) settingsCache.autoDownload = changes.autoDownload.newValue;
  if (changes.minLikes) settingsCache.minLikes = Number(changes.minLikes.newValue) || SETTINGS_DEFAULT.minLikes;
});

function startScanner() {
  if (scanTimer) clearInterval(scanTimer);
  scanTimer = setInterval(() => {
    installDownloadButtons();
    if (settingsCache.autoDownload) {
      autoDownloadGoodTweets();
    }
  }, 1500);
}

function installDownloadButtons() {
  const tweets = document.querySelectorAll('article[data-testid="tweet"]');
  tweets.forEach((article) => {
    if (processedTweetNodes.has(article)) return;

    const actionBar = article.querySelector('div[role="group"]');
    if (!actionBar) return;

    const btn = document.createElement('button');
    btn.textContent = '下载DOC';
    btn.style.marginLeft = '8px';
    btn.style.padding = '4px 8px';
    btn.style.border = '1px solid #1d9bf0';
    btn.style.background = 'transparent';
    btn.style.color = '#1d9bf0';
    btn.style.borderRadius = '9999px';
    btn.style.cursor = 'pointer';
    btn.style.fontSize = '12px';

    btn.addEventListener('click', async (event) => {
      event.preventDefault();
      event.stopPropagation();
      const payload = extractTweetData(article);
      if (!payload) return;
      const res = await sendDownload(payload);
      btn.textContent = res?.skipped ? '已下载' : '已保存';
      btn.disabled = true;
      btn.style.opacity = '0.7';
    });

    actionBar.appendChild(btn);
    processedTweetNodes.add(article);
  });
}

async function autoDownloadGoodTweets() {
  const tweets = document.querySelectorAll('article[data-testid="tweet"]');
  for (const article of tweets) {
    const payload = extractTweetData(article);
    if (!payload) continue;
    if ((payload.likes || 0) < (Number(settingsCache.minLikes) || 0)) continue;
    await sendDownload(payload);
  }
}

function extractTweetData(article) {
  const statusLink = article.querySelector('a[href*="/status/"]');
  if (!statusLink) return null;

  const href = statusLink.getAttribute('href') || '';
  const match = href.match(/status\/(\d+)/);
  if (!match) return null;

  const tweetId = match[1];
  const baseUrl = location.hostname.includes('twitter.com') ? 'https://twitter.com' : 'https://x.com';
  const url = new URL(href, baseUrl).toString();

  const textNode = article.querySelector('[data-testid="tweetText"]');
  const authorNode = article.querySelector('[data-testid="User-Name"] span');
  const timeNode = article.querySelector('time');
  const likeNode = article.querySelector('[data-testid="like"]');

  return {
    tweetId,
    url,
    text: textNode ? textNode.innerText.trim() : '',
    author: authorNode ? authorNode.innerText.trim() : '',
    publishTime: timeNode?.getAttribute('datetime') || '',
    likes: parseLikes(likeNode)
  };
}

function parseLikes(likeNode) {
  if (!likeNode) return 0;
  const text = likeNode.getAttribute('aria-label') || likeNode.innerText || '';
  const normalized = text.toLowerCase().replace(/,/g, '').trim();
  const match = normalized.match(/([\d.]+)\s*([km]?)/);
  if (!match) return 0;
  let value = parseFloat(match[1]);
  if (Number.isNaN(value)) return 0;
  const unit = match[2];
  if (unit === 'k') value *= 1000;
  if (unit === 'm') value *= 1000000;
  return Math.round(value);
}

function sendDownload(payload) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: 'DOWNLOAD_TWEET', payload }, (response) => {
      resolve(response?.result || null);
    });
  });
}
