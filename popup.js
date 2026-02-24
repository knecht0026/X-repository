const statusEl = document.getElementById('status');
const docLinkEl = document.getElementById('docLink');

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? '#f4212e' : '#00ba7c';
}

async function loadSettings() {
  const settings = await chrome.storage.local.get([
    'docTitle',
    'docId',
    'aiEndpoint',
    'aiModel',
    'aiApiKey'
  ]);

  document.getElementById('docTitle').value = settings.docTitle || '';
  document.getElementById('docId').value = settings.docId || '';
  document.getElementById('aiEndpoint').value = settings.aiEndpoint || '';
  document.getElementById('aiModel').value = settings.aiModel || '';
  document.getElementById('aiApiKey').value = settings.aiApiKey || '';
}

async function saveSettings() {
  await chrome.storage.local.set({
    docTitle: document.getElementById('docTitle').value.trim(),
    docId: document.getElementById('docId').value.trim(),
    aiEndpoint: document.getElementById('aiEndpoint').value.trim(),
    aiModel: document.getElementById('aiModel').value.trim(),
    aiApiKey: document.getElementById('aiApiKey').value.trim()
  });
  setStatus('设置已保存。');
}

async function exportTweets() {
  setStatus('正在读取当前页面推文...');
  docLinkEl.textContent = '';
  docLinkEl.removeAttribute('href');

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    setStatus('无法找到当前标签页。', true);
    return;
  }

  const extracted = await chrome.tabs.sendMessage(tab.id, { type: 'EXTRACT_TWEETS' });
  if (!extracted?.ok) {
    setStatus(extracted?.error || '读取推文失败。', true);
    return;
  }

  if (!extracted.tweets?.length) {
    setStatus('当前页面没有可导出的推文。', true);
    return;
  }

  setStatus(`已获取 ${extracted.tweets.length} 条推文，正在写入 Google Docs...`);
  const result = await chrome.runtime.sendMessage({
    type: 'EXPORT_TWEETS',
    tweets: extracted.tweets
  });

  if (!result?.ok) {
    setStatus(result?.error || '导出失败。', true);
    return;
  }

  setStatus('导出完成！');
  docLinkEl.href = result.docUrl;
  docLinkEl.textContent = '打开 Google Docs';
}

document.getElementById('saveSettings').addEventListener('click', saveSettings);
document.getElementById('exportBtn').addEventListener('click', exportTweets);
loadSettings();
