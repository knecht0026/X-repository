const DEFAULT_SETTINGS = {
  autoDownload: false,
  minLikes: 100,
  downloadedIds: []
};

chrome.runtime.onInstalled.addListener(async () => {
  const current = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  const next = { ...DEFAULT_SETTINGS, ...current };
  if (!Array.isArray(next.downloadedIds)) {
    next.downloadedIds = [];
  }
  await chrome.storage.local.set(next);
});

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === 'DOWNLOAD_TWEET') {
    handleDownload(message.payload)
      .then((result) => sendResponse({ ok: true, result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }
  return false;
});

async function handleDownload(tweet) {
  if (!tweet?.tweetId) {
    throw new Error('tweetId 缺失，无法下载。');
  }

  const { downloadedIds = [] } = await chrome.storage.local.get('downloadedIds');
  if (downloadedIds.includes(tweet.tweetId)) {
    return { skipped: true, reason: 'already_downloaded' };
  }

  const html = buildWordHtml(tweet);
  const dataUrl = 'data:application/msword;charset=utf-8,' + encodeURIComponent(html);
  const safeAuthor = sanitizeFileName(tweet.author || 'unknown');
  const safeId = sanitizeFileName(tweet.tweetId);
  const fileName = `x-tweets/${safeAuthor}-${safeId}.doc`;

  await chrome.downloads.download({
    url: dataUrl,
    filename: fileName,
    saveAs: false,
    conflictAction: 'uniquify'
  });

  downloadedIds.push(tweet.tweetId);
  await chrome.storage.local.set({ downloadedIds });

  return { skipped: false, fileName };
}

function buildWordHtml(tweet) {
  const now = new Date().toLocaleString('zh-CN', { hour12: false });
  const text = escapeHtml(tweet.text || '');
  const url = escapeHtml(tweet.url || '');
  const author = escapeHtml(tweet.author || '');
  const likes = Number.isFinite(tweet.likes) ? tweet.likes : 0;
  const publishTime = escapeHtml(tweet.publishTime || '');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <title>X Tweet ${escapeHtml(tweet.tweetId)}</title>
  <style>
    body { font-family: 'Calibri', 'Microsoft YaHei', sans-serif; line-height: 1.7; padding: 24px; }
    h1 { font-size: 22px; margin-bottom: 8px; }
    .meta { color: #555; margin-bottom: 14px; }
    .tweet { border-left: 4px solid #1d9bf0; padding-left: 12px; white-space: pre-wrap; }
    .small { font-size: 12px; color: #777; margin-top: 20px; }
  </style>
</head>
<body>
  <h1>X 推文备份</h1>
  <div class="meta"><b>作者：</b>${author}</div>
  <div class="meta"><b>点赞：</b>${likes}</div>
  <div class="meta"><b>发布时间：</b>${publishTime || '未知'}</div>
  <div class="meta"><b>链接：</b><a href="${url}">${url}</a></div>
  <div class="tweet">${text}</div>
  <div class="small">导出时间：${escapeHtml(now)}</div>
</body>
</html>`;
}

function sanitizeFileName(input) {
  return String(input).replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 80);
}

function escapeHtml(input) {
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
