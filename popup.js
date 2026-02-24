const autoDownloadInput = document.getElementById('autoDownload');
const minLikesInput = document.getElementById('minLikes');
const saveBtn = document.getElementById('saveBtn');
const statusText = document.getElementById('status');

init();

async function init() {
  const { autoDownload = false, minLikes = 100 } = await chrome.storage.local.get(['autoDownload', 'minLikes']);
  autoDownloadInput.checked = autoDownload;
  minLikesInput.value = String(minLikes);
}

saveBtn.addEventListener('click', async () => {
  const nextAuto = autoDownloadInput.checked;
  const nextMinLikes = Math.max(0, parseInt(minLikesInput.value || '0', 10));

  await chrome.storage.local.set({
    autoDownload: nextAuto,
    minLikes: nextMinLikes
  });

  statusText.textContent = '设置已保存';
  setTimeout(() => {
    statusText.textContent = '';
  }, 1500);
});
