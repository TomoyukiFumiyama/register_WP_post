/**
 * Spreadsheet columns (1-indexed)
 * A: 入稿ステータス
 * B: GoogleドキュメントURL
 * C: WP_STATUS
 * D: WP_SLUG
 * E: WP_CATEGORIES (comma separated)
 * F: WP_TAGS (comma separated)
 * G: WP_POST_ID
 * H: WP_POST_URL
 * I: 最終実行日時
 * J: 最終結果
 * K: エラーメッセージ
 */
const COL = {
  STATUS: 1,
  DOC_URL: 2,
  WP_STATUS: 3,
  WP_SLUG: 4,
  WP_CATEGORIES: 5,
  WP_TAGS: 6,
  WP_POST_ID: 7,
  WP_POST_URL: 8,
  LAST_EXECUTED_AT: 9,
  LAST_RESULT: 10,
  ERROR_MESSAGE: 11,
};

const HEADER_ROW = [
  '入稿ステータス',
  'GoogleドキュメントURL',
  'WP_STATUS',
  'WP_SLUG',
  'WP_CATEGORIES',
  'WP_TAGS',
  'WP_POST_ID',
  'WP_POST_URL',
  '最終実行日時',
  '最終結果',
  'エラーメッセージ',
];

const WAITING_STATUS = '入稿待ち';
const DEFAULT_WP_STATUS = 'draft';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WordPress連携')
    .addItem('1行目にヘッダーを設定', 'setupHeader')
    .addItem('入稿待ちを処理', 'processWaitingRows')
    .addToUi();
}

function setupHeader() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  sheet.setFrozenRows(1);
}

function processWaitingRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  const config = getConfig_();

  rows.forEach(function(row, index) {
    const rowNumber = index + 2;
    const ingestStatus = normalizeString_(row[COL.STATUS - 1]);
    if (ingestStatus !== WAITING_STATUS) {
      return;
    }

    try {
      const docUrl = normalizeString_(row[COL.DOC_URL - 1]);
      if (!docUrl) {
        throw new Error('GoogleドキュメントURL(B列)が空です。');
      }

      const docId = extractDocId_(docUrl);
      const doc = DocumentApp.openById(docId);
      const title = doc.getName();
      const content = doc.getBody().getText();

      const wpStatus = normalizeString_(row[COL.WP_STATUS - 1]) || DEFAULT_WP_STATUS;
      const slug = normalizeString_(row[COL.WP_SLUG - 1]);
      const categories = parseCsv_(row[COL.WP_CATEGORIES - 1]);
      const tags = parseCsv_(row[COL.WP_TAGS - 1]);
      const existingPostId = normalizeString_(row[COL.WP_POST_ID - 1]);

      const categoryIds = ensureTerms_(config, 'categories', categories);
      const tagIds = ensureTerms_(config, 'tags', tags);

      const payload = {
        title: title,
        content: content,
        status: wpStatus,
      };
      if (slug) {
        payload.slug = slug;
      }
      if (categoryIds.length > 0) {
        payload.categories = categoryIds;
      }
      if (tagIds.length > 0) {
        payload.tags = tagIds;
      }

      const result = upsertPost_(config, payload, existingPostId);

      sheet.getRange(rowNumber, COL.WP_POST_ID).setValue(result.id);
      sheet.getRange(rowNumber, COL.WP_POST_URL).setValue(result.link || '');
      sheet.getRange(rowNumber, COL.LAST_EXECUTED_AT).setValue(new Date());
      sheet.getRange(rowNumber, COL.LAST_RESULT).setValue('成功');
      sheet.getRange(rowNumber, COL.ERROR_MESSAGE).setValue('');
    } catch (error) {
      sheet.getRange(rowNumber, COL.LAST_EXECUTED_AT).setValue(new Date());
      sheet.getRange(rowNumber, COL.LAST_RESULT).setValue('失敗');
      sheet.getRange(rowNumber, COL.ERROR_MESSAGE).setValue(error.message || String(error));
    }
  });
}

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = normalizeString_(props.getProperty('WP_BASE_URL'));
  const username = normalizeString_(props.getProperty('WP_USERNAME'));
  const appPassword = normalizeString_(props.getProperty('WP_APP_PASSWORD'));

  if (!baseUrl || !username || !appPassword) {
    throw new Error('スクリプトプロパティに WP_BASE_URL / WP_USERNAME / WP_APP_PASSWORD を設定してください。');
  }

  return {
    baseUrl: baseUrl.replace(/\/$/, ''),
    authHeader: 'Basic ' + Utilities.base64Encode(username + ':' + appPassword),
  };
}

function parseCsv_(value) {
  return normalizeString_(value)
    .split(',')
    .map(function(v) { return normalizeString_(v); })
    .filter(function(v) { return v !== ''; });
}

function normalizeString_(value) {
  return String(value == null ? '' : value).trim();
}

function extractDocId_(url) {
  const match = url.match(/\/document\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error('GoogleドキュメントURL(B列)からドキュメントIDを取得できませんでした。');
  }
  return match[1];
}

function ensureTerms_(config, taxonomy, names) {
  return names.map(function(name) {
    const existing = findTermByName_(config, taxonomy, name);
    if (existing) {
      return existing.id;
    }
    return createTerm_(config, taxonomy, name).id;
  });
}

function findTermByName_(config, taxonomy, name) {
  const endpoint = '/wp-json/wp/v2/' + taxonomy + '?search=' + encodeURIComponent(name) + '&per_page=100';
  const res = wpRequest_(config, endpoint, 'get');
  const items = JSON.parse(res.getContentText());
  return items.find(function(item) {
    return normalizeString_(item.name) === name;
  }) || null;
}

function createTerm_(config, taxonomy, name) {
  const endpoint = '/wp-json/wp/v2/' + taxonomy;
  const res = wpRequest_(config, endpoint, 'post', { name: name });
  return JSON.parse(res.getContentText());
}

function upsertPost_(config, payload, postId) {
  const endpoint = postId
    ? '/wp-json/wp/v2/posts/' + encodeURIComponent(postId)
    : '/wp-json/wp/v2/posts';
  const method = postId ? 'post' : 'post';
  const res = wpRequest_(config, endpoint, method, payload);
  return JSON.parse(res.getContentText());
}

function wpRequest_(config, endpoint, method, payload) {
  const url = config.baseUrl + endpoint;
  const options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: config.authHeader,
    },
  };

  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('WordPress APIエラー: HTTP ' + code + ' / ' + response.getContentText());
  }

  return response;
}
