const DOCS_BASE_URL = 'https://docs.googleapis.com/v1/documents';

function formatYYMM(isoDate) {
  const d = isoDate ? new Date(isoDate) : new Date();
  const yy = String(d.getFullYear()).slice(2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${yy}/${mm}`;
}

function heuristicTopic(tweetText) {
  const cleaned = tweetText
    .replace(/https?:\/\/\S+/g, '')
    .replace(/[“”"'’]/g, '')
    .trim();

  const sentence = cleaned.split(/[。.!?！？\n]/).find(Boolean) || cleaned;
  return sentence.slice(0, 18) || '未命名主题';
}

async function generateTopicWithAI(tweetText, settings) {
  const { aiApiKey, aiEndpoint, aiModel } = settings || {};
  if (!aiApiKey || !aiEndpoint || !aiModel) {
    return heuristicTopic(tweetText);
  }

  const response = await fetch(aiEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${aiApiKey}`
    },
    body: JSON.stringify({
      model: aiModel,
      messages: [
        {
          role: 'system',
          content:
            '你是一个摘要助手。请为推文生成一个简短中文主题，不超过12个字，不加引号。'
        },
        {
          role: 'user',
          content: tweetText
        }
      ],
      temperature: 0.4
    })
  });

  if (!response.ok) {
    throw new Error(`AI 请求失败: ${response.status}`);
  }

  const data = await response.json();
  const topic =
    data?.choices?.[0]?.message?.content?.trim() || heuristicTopic(tweetText);
  return topic.slice(0, 20);
}

async function getAuthToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, token => {
      if (chrome.runtime.lastError || !token) {
        reject(
          new Error(
            chrome.runtime.lastError?.message || '无法获取 Google 授权令牌'
          )
        );
        return;
      }
      resolve(token);
    });
  });
}

async function createDocument(token, title) {
  const response = await fetch(DOCS_BASE_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ title })
  });

  if (!response.ok) {
    throw new Error(`创建文档失败: ${response.status}`);
  }

  return response.json();
}

async function batchUpdateDocument(token, documentId, requests) {
  const response = await fetch(`${DOCS_BASE_URL}/${documentId}:batchUpdate`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ requests })
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`写入文档失败: ${response.status} ${text}`);
  }

  return response.json();
}

function buildDocRequests(items) {
  let cursor = 1;
  const requests = [];

  for (const item of items) {
    const heading = `${item.yymm}-【${item.topic}-由AI生成】\n`;
    requests.push({
      insertText: {
        location: { index: cursor },
        text: heading
      }
    });

    requests.push({
      updateParagraphStyle: {
        range: {
          startIndex: cursor,
          endIndex: cursor + heading.length
        },
        paragraphStyle: {
          namedStyleType: 'HEADING_2'
        },
        fields: 'namedStyleType'
      }
    });

    cursor += heading.length;

    const body = `${item.text}\n\n`;
    requests.push({
      insertText: {
        location: { index: cursor },
        text: body
      }
    });

    requests.push({
      updateParagraphStyle: {
        range: {
          startIndex: cursor,
          endIndex: cursor + body.length
        },
        paragraphStyle: {
          namedStyleType: 'NORMAL_TEXT'
        },
        fields: 'namedStyleType'
      }
    });

    cursor += body.length;
  }

  return requests;
}

async function getSettings() {
  return chrome.storage.local.get([
    'docId',
    'docTitle',
    'aiApiKey',
    'aiEndpoint',
    'aiModel'
  ]);
}

chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request?.type !== 'EXPORT_TWEETS') return;

  (async () => {
    const settings = await getSettings();
    const token = await getAuthToken();

    let docId = settings.docId;
    if (!docId) {
      const created = await createDocument(
        token,
        settings.docTitle || 'X 推文导出合集'
      );
      docId = created.documentId;
      await chrome.storage.local.set({ docId });
    }

    const enriched = [];
    for (const tweet of request.tweets || []) {
      let topic;
      try {
        topic = await generateTopicWithAI(tweet.text, settings);
      } catch (_error) {
        topic = heuristicTopic(tweet.text);
      }

      enriched.push({
        ...tweet,
        topic,
        yymm: formatYYMM(tweet.datetime)
      });
    }

    const requests = buildDocRequests(enriched);
    await batchUpdateDocument(token, docId, requests);

    sendResponse({
      ok: true,
      docId,
      docUrl: `https://docs.google.com/document/d/${docId}/edit`
    });
  })().catch(error => {
    sendResponse({ ok: false, error: String(error) });
  });

  return true;
});
