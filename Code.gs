/**
 * Spreadsheet columns (1-indexed)
 * A: 入稿ステータス
 * B: GoogleドキュメントURL
 * C: WP_POST_TYPE (post / page / カスタム投稿タイプのスラッグ。空欄なら post)
 * D: WP_STATUS
 * E: WP_SLUG
 * F: WP_CATEGORIES (comma separated)
 * G: WP_TAGS (comma separated)
 * H: WP_POST_ID
 * I: WP_POST_URL
 * J: 最終実行日時
 * K: 最終結果
 * L: エラーメッセージ
 */
const COL = {
  STATUS: 1,
  DOC_URL: 2,
  WP_POST_TYPE: 3,
  WP_STATUS: 4,
  WP_SLUG: 5,
  WP_CATEGORIES: 6,
  WP_TAGS: 7,
  WP_POST_ID: 8,
  WP_POST_URL: 9,
  LAST_EXECUTED_AT: 10,
  LAST_RESULT: 11,
  ERROR_MESSAGE: 12,
};

const HEADER_ROW = [
  '入稿ステータス',
  'GoogleドキュメントURL',
  'WP_POST_TYPE',
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

// 旧レイアウト（C列が WP_STATUS）の判定に使うヘッダー名
const LEGACY_HEADER_AT_POST_TYPE_COL = 'WP_STATUS';

const WAITING_STATUS = '入稿待ち';
const DONE_STATUS = '入稿済み';
const DEFAULT_WP_STATUS = 'draft';
const DEFAULT_POST_TYPE = {
  slug: 'post',
  name: '投稿',
  restBase: 'posts',
  restNamespace: 'wp/v2',
  taxonomies: ['category', 'post_tag'],
};

// Apps Script の実行時間上限(6分)で行の途中に強制終了されないよう、新しい行の処理を止める目安
const MAX_RUNTIME_MS = 4.5 * 60 * 1000;
// 検出した REST API ルートのキャッシュ時間（CacheService の上限は 6 時間）
const API_ROOT_CACHE_TTL_SECONDS = 6 * 60 * 60;
// UrlFetchApp.fetchAll で同時に送るリクエスト数
const FETCH_BATCH_SIZE = 10;
const MEDIA_UPLOAD_BATCH_SIZE = 5;
// 画像ファイル名に付ける内容ハッシュの長さ（再入稿時に同じ画像を再利用するための目印）
const MEDIA_HASH_LENGTH = 12;
const IMAGE_PLACEHOLDER_PATTERN = /\[\[G_DOC_IMAGE_\d+\]\]/g;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WordPress連携')
    .addItem('1行目にヘッダーを設定', 'setupHeader')
    .addItem('入稿待ちを処理', 'processWaitingRows')
    .addToUi();
}

function setupHeader() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (isLegacyLayout_(sheet)) {
    // 旧レイアウトのデータを保ったまま、C列に WP_POST_TYPE 列を差し込む
    sheet.insertColumnBefore(COL.WP_POST_TYPE);
  }
  sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  sheet.setFrozenRows(1);
}

function isLegacyLayout_(sheet) {
  const header = sheet.getRange(1, COL.WP_POST_TYPE).getValue();
  return normalizeString_(header) === LEGACY_HEADER_AT_POST_TYPE_COL;
}

function processWaitingRows() {
  const startedAt = Date.now();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const values = sheet.getRange(1, 1, lastRow, HEADER_ROW.length).getValues();
  if (normalizeString_(values[0][COL.WP_POST_TYPE - 1]) === LEGACY_HEADER_AT_POST_TYPE_COL) {
    throw new Error(
      'シートが旧レイアウト（C列=WP_STATUS）のままです。' +
      'メニュー「WordPress連携」→「1行目にヘッダーを設定」を実行すると、データを保ったまま C列に WP_POST_TYPE 列が追加されます。'
    );
  }

  const waitingRows = [];
  for (let i = 1; i < values.length; i += 1) {
    if (normalizeString_(values[i][COL.STATUS - 1]) === WAITING_STATUS) {
      waitingRows.push({ rowNumber: i + 1, row: values[i] });
    }
  }
  if (waitingRows.length === 0) {
    notify_('「' + WAITING_STATUS + '」の行はありません。');
    return;
  }

  // 入稿待ちの行があるときだけ WordPress への接続設定を行う
  const config = getConfig_();
  let successCount = 0;
  let failureCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < waitingRows.length; i += 1) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      skippedCount = waitingRows.length - i;
      break;
    }
    const item = waitingRows[i];
    if (processRow_(sheet, config, item.rowNumber, item.row)) {
      successCount += 1;
    } else {
      failureCount += 1;
    }
  }

  let message = '成功 ' + successCount + '件 / 失敗 ' + failureCount + '件';
  if (skippedCount > 0) {
    message += ' / 実行時間の上限が近いため、残り ' + skippedCount + '件は処理していません。もう一度実行すると続きから処理します。';
  }
  notify_(message);
}

/**
 * 1行分を入稿し、結果をシートに書き込む。成功したら true を返す。
 * 成功した行は A列を「入稿済み」にして、次回以降の処理対象から外す。
 */
function processRow_(sheet, config, rowNumber, row) {
  if (!publishRow_(sheet, config, rowNumber, row)) {
    return false;
  }

  try {
    sheet.getRange(rowNumber, COL.STATUS).setValue(DONE_STATUS);
  } catch (error) {
    // 入力規則などで A列を変更できなくても、入稿と投稿IDの保存は済んでいるので成功として扱う
    sheet.getRange(rowNumber, COL.ERROR_MESSAGE).setValue(
      '入稿は成功しましたが、A列を「' + DONE_STATUS + '」に変更できませんでした（A列の入力規則を確認してください）: ' +
      (error.message || String(error))
    );
  }
  return true;
}

function publishRow_(sheet, config, rowNumber, row) {
  try {
    const docUrl = normalizeString_(row[COL.DOC_URL - 1]);
    if (!docUrl) {
      throw new Error('GoogleドキュメントURL(B列)が空です。');
    }

    // ドキュメントの変換や画像アップロードの前に、シートの設定値を検証する
    const postType = resolvePostType_(config, row[COL.WP_POST_TYPE - 1]);
    const wpStatus = normalizeString_(row[COL.WP_STATUS - 1]) || DEFAULT_WP_STATUS;
    const slug = normalizeString_(row[COL.WP_SLUG - 1]);
    const categories = parseCsv_(row[COL.WP_CATEGORIES - 1]);
    const tags = parseCsv_(row[COL.WP_TAGS - 1]);
    const existingPostId = normalizeString_(row[COL.WP_POST_ID - 1]);
    assertTaxonomySupported_(postType, 'category', categories, 'カテゴリー', 'F列(WP_CATEGORIES)');
    assertTaxonomySupported_(postType, 'post_tag', tags, 'タグ', 'G列(WP_TAGS)');

    const doc = DocumentApp.openById(extractDocId_(docUrl));
    const title = doc.getName();
    const docContent = convertGoogleDocBodyToHtml_(doc.getBody());

    const categoryIds = ensureTermIds_(config, 'categories', categories);
    const tagIds = ensureTermIds_(config, 'tags', tags);

    const contentWithImages = replaceImagePlaceholders_(
      docContent.html,
      uploadDocumentImages_(config, docContent.images, title)
    );
    const content = applyDecorations_(contentWithImages);

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

    const result = upsertPost_(config, postType, payload, existingPostId);

    // H〜L列をまとめて1回で書き込む
    sheet
      .getRange(rowNumber, COL.WP_POST_ID, 1, COL.ERROR_MESSAGE - COL.WP_POST_ID + 1)
      .setValues([[result.id, result.link || '', new Date(), '成功', '']]);
    return true;
  } catch (error) {
    sheet
      .getRange(rowNumber, COL.LAST_EXECUTED_AT, 1, COL.ERROR_MESSAGE - COL.LAST_EXECUTED_AT + 1)
      .setValues([[new Date(), '失敗', error.message || String(error)]]);
    return false;
  }
}

function notify_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'WordPress連携', 10);
  } catch (e) {
    // トリガー実行などで UI が無い場合は通知を省略する
  }
}

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = normalizeString_(props.getProperty('WP_BASE_URL'));
  const username = normalizeString_(props.getProperty('WP_USERNAME'));
  const appPassword = normalizeString_(props.getProperty('WP_APP_PASSWORD'));

  if (!baseUrl || !username || !appPassword) {
    throw new Error('スクリプトプロパティに WP_BASE_URL / WP_USERNAME / WP_APP_PASSWORD を設定してください。');
  }

  const normalizedBaseUrl = normalizeWpBaseUrl_(baseUrl);
  const authHeader = 'Basic ' + Utilities.base64Encode(username + ':' + appPassword);
  const apiRoot = resolveApiRoot_(normalizedBaseUrl, authHeader);

  return {
    baseUrl: normalizedBaseUrl,
    apiRoot: apiRoot,
    authHeader: authHeader,
    // 1回の実行内で使い回すキャッシュ
    postTypes: null,
    termCache: {},
  };
}

function normalizeWpBaseUrl_(baseUrl) {
  return baseUrl
    .replace(/\/$/, '')
    .replace(/\/wp-admin\/?.*$/i, '')
    .replace(/\/wp-json\/wp\/v2\/?$/i, '')
    .replace(/\/wp-json\/?$/i, '')
    .replace(/\/index\.php\?rest_route=.*$/i, '')
    .replace(/\/\?rest_route=.*$/i, '');
}

/**
 * REST API ルートの検出はトップページ取得＋複数URLの試行で時間がかかるため、結果をキャッシュする。
 */
function resolveApiRoot_(baseUrl, authHeader) {
  const cacheKey = apiRootCacheKey_(baseUrl);
  const cached = cacheGet_(cacheKey);
  if (cached) {
    return cached;
  }
  const apiRoot = detectApiRoot_(baseUrl, authHeader);
  cachePut_(cacheKey, apiRoot, API_ROOT_CACHE_TTL_SECONDS);
  return apiRoot;
}

function clearCachedApiRoot_(baseUrl) {
  cacheRemove_(apiRootCacheKey_(baseUrl));
}

function apiRootCacheKey_(baseUrl) {
  return 'WP_API_ROOT_' + computeHexDigest_(baseUrl);
}

function cacheGet_(key) {
  try {
    return CacheService.getScriptCache().get(key);
  } catch (e) {
    return null;
  }
}

function cachePut_(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, value, ttlSeconds);
  } catch (e) {
    // キャッシュは最適化のためだけに使うので、失敗しても処理は続ける
  }
}

function cacheRemove_(key) {
  try {
    CacheService.getScriptCache().remove(key);
  } catch (e) {
    // 同上
  }
}

function detectApiRoot_(baseUrl, authHeader) {
  const candidates = buildApiCandidates_(baseUrl);
  const uniqueCandidates = Array.from(new Set(candidates));
  const probeResults = [];

  for (let i = 0; i < uniqueCandidates.length; i += 1) {
    const candidate = uniqueCandidates[i];
    const response = UrlFetchApp.fetch(appendQuery_(candidate + '/posts', { per_page: 1, _fields: 'id' }), {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: authHeader,
      },
    });

    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return candidate;
    }
    probeResults.push(candidate + ' -> HTTP ' + code);
  }

  throw new Error(
    'WordPress REST API に接続できませんでした。WP_BASE_URL を確認してください。' +
    ' 例: https://example.com または https://example.com/wordpress' +
    '\n確認したURL: ' + probeResults.join(' / ') +
    '\n補足: WP_BASE_URL には投稿ページURLではなく、WordPress設置先のURLを指定してください。'
  );
}

function buildApiCandidates_(baseUrl) {
  const defaults = [
    baseUrl + '/wp-json/wp/v2',
    baseUrl + '/index.php?rest_route=/wp/v2',
    baseUrl + '/?rest_route=/wp/v2',
    baseUrl + '/wp/wp-json/wp/v2',
    baseUrl + '/wordpress/wp-json/wp/v2',
  ];
  const discovered = discoverApiCandidates_(baseUrl);
  return discovered.concat(defaults);
}

function discoverApiCandidates_(baseUrl) {
  const discovered = [];
  try {
    const response = UrlFetchApp.fetch(baseUrl, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
    });

    const linkHeader = response.getHeaders().Link || response.getHeaders().link || '';
    const linkMatches = String(linkHeader).match(/<([^>]+)>\s*;\s*rel="?https:\/\/api\.w\.org\/"?/i);
    if (linkMatches && linkMatches[1]) {
      discovered.push(toWpV2Root_(linkMatches[1]));
    }

    const html = response.getContentText();
    const relHrefPattern = /rel=["']https:\/\/api\.w\.org\/["'][^>]*href=["']([^"']+)["']/i;
    const relHrefMatches = html.match(relHrefPattern);
    if (relHrefMatches && relHrefMatches[1]) {
      discovered.push(toWpV2Root_(relHrefMatches[1]));
    }
  } catch (e) {
    // Discovery is best-effort.
  }
  return discovered.filter(function(v) { return normalizeString_(v) !== ''; });
}

function toWpV2Root_(apiRootOrIndex) {
  const value = String(apiRootOrIndex).trim();
  // 例: https://example.com/?rest_route=/ → https://example.com/?rest_route=/wp/v2
  const restRouteMatch = value.match(/^(.*?[?&]rest_route=)/i);
  if (restRouteMatch) {
    return restRouteMatch[1] + '/wp/v2';
  }
  const cleaned = value
    .replace(/\/$/, '')
    .replace(/\/wp-json\/?$/i, '');
  return cleaned + '/wp-json/wp/v2';
}

/**
 * apiRoot（.../wp-json/wp/v2 または ...?rest_route=/wp/v2）にルートとクエリを付けたURLを作る。
 * namespace を指定すると wp/v2 以外の名前空間（カスタム投稿タイプ用など）を使う。
 */
function buildApiUrl_(config, route, query, namespace) {
  let root = config.apiRoot;
  if (namespace && namespace !== 'wp/v2') {
    root = root.replace(/wp\/v2$/, namespace);
  }
  return appendQuery_(root + route, query);
}

function appendQuery_(url, query) {
  const keys = Object.keys(query || {});
  if (keys.length === 0) {
    return url;
  }
  const queryString = keys
    .map(function(key) { return encodeURIComponent(key) + '=' + encodeURIComponent(query[key]); })
    .join('&');
  // ?rest_route= 形式のURLでは、追加のパラメータを & でつなぐ必要がある
  return url + (url.indexOf('?') === -1 ? '?' : '&') + queryString;
}

/**
 * UrlFetchApp.fetch / fetchAll 共通のリクエストオブジェクトを作る。
 * options: { query, payload, namespace }
 */
function buildWpRequest_(config, method, route, options) {
  const opts = options || {};
  const request = {
    url: buildApiUrl_(config, route, opts.query, opts.namespace),
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: config.authHeader,
    },
  };
  if (opts.payload) {
    request.contentType = 'application/json';
    request.payload = JSON.stringify(opts.payload);
  }
  return request;
}

function wpFetch_(config, method, route, options) {
  const request = buildWpRequest_(config, method, route, options);
  const params = Object.assign({}, request);
  delete params.url;
  return UrlFetchApp.fetch(request.url, params);
}

function fetchAllInBatches_(requests, batchSize) {
  const responses = [];
  for (let i = 0; i < requests.length; i += batchSize) {
    const batch = UrlFetchApp.fetchAll(requests.slice(i, i + batchSize));
    Array.prototype.push.apply(responses, batch);
  }
  return responses;
}

function parseWpResponse_(config, response, label) {
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw createWpError_(config, response, label);
  }
  return JSON.parse(response.getContentText());
}

function createWpError_(config, response, label) {
  const code = response.getResponseCode();
  const rawBody = response.getContentText();
  const json = tryParseJson_(rawBody);
  let message;
  if (json && typeof json === 'object' && json.message) {
    message = label + ': HTTP ' + code + ' / ' + json.message + (json.code ? ' [' + json.code + ']' : '');
  } else {
    // REST API ではなく HTML などが返った場合は接続先が誤っている可能性が高いため、検出結果のキャッシュも破棄する
    clearCachedApiRoot_(config.baseUrl);
    message = label + ': HTTP ' + code + ' / ' + stripHtml_(rawBody).slice(0, 300) +
      '（WP_BASE_URL または REST API URL が誤っている可能性があります）';
  }
  const error = new Error(message);
  error.httpCode = code;
  error.wpCode = json && json.code ? json.code : '';
  return error;
}

function tryParseJson_(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

/**
 * C列の値から投稿タイプを決める。空欄・post・posts は追加の通信なしで「投稿」として扱う。
 * それ以外はスラッグ / REST ベース名 / 表示名（例: 固定ページ）で照合する。
 */
function resolvePostType_(config, rawValue) {
  const value = normalizeString_(rawValue);
  const key = value.toLowerCase();
  if (!value || key === DEFAULT_POST_TYPE.slug || key === DEFAULT_POST_TYPE.restBase) {
    return DEFAULT_POST_TYPE;
  }

  const postTypes = getPostTypes_(config);
  const found = postTypes.find(function(type) {
    return type.slug.toLowerCase() === key ||
      type.restBase.toLowerCase() === key ||
      type.name === value;
  });
  if (found) {
    return found;
  }

  const available = postTypes
    .filter(function(type) { return type.slug.indexOf('wp_') !== 0 && type.slug !== 'nav_menu_item'; })
    .map(function(type) { return type.slug + (type.name ? '（' + type.name + '）' : ''); })
    .join(', ');
  throw new Error(
    'C列の投稿タイプ「' + value + '」が WordPress REST API で見つかりません。指定できる値: ' + available +
    '。カスタム投稿タイプの場合は show_in_rest が有効か確認してください。'
  );
}

function getPostTypes_(config) {
  if (!config.postTypes) {
    // 投稿タイプ一覧は1回の実行で1度だけ取得する
    const json = parseWpResponse_(config, wpFetch_(config, 'get', '/types'), 'WordPress APIエラー');
    config.postTypes = Object.keys(json)
      .map(function(key) {
        const type = json[key] || {};
        return {
          slug: normalizeString_(type.slug) || key,
          name: decodeHtmlEntities_(normalizeString_(type.name)),
          restBase: normalizeString_(type.rest_base) || key,
          restNamespace: normalizeString_(type.rest_namespace) || 'wp/v2',
          taxonomies: Array.isArray(type.taxonomies) ? type.taxonomies : [],
        };
      })
      .filter(function(type) { return type.slug !== 'attachment'; });
  }
  return config.postTypes;
}

function assertTaxonomySupported_(postType, taxonomy, names, label, columnLabel) {
  if (names.length > 0 && postType.taxonomies.indexOf(taxonomy) === -1) {
    throw new Error(
      '投稿タイプ「' + postType.slug + '」は' + label + 'に対応していません。' + columnLabel + 'を空にしてください。'
    );
  }
}


function convertGoogleDocBodyToHtml_(body) {
  const images = [];
  const chunks = [];
  const total = body.getNumChildren();

  for (let i = 0; i < total; i += 1) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const html = convertParagraphToHtml_(paragraph, images);
      if (html) {
        chunks.push(html);
      }
      continue;
    }

    if (type === DocumentApp.ElementType.LIST_ITEM) {
      const listItems = [child.asListItem()];
      while (i + 1 < total) {
        const next = body.getChild(i + 1);
        if (next.getType() !== DocumentApp.ElementType.LIST_ITEM) {
          break;
        }
        listItems.push(next.asListItem());
        i += 1;
      }

      const listHtml = listItems
        .map(function(item) { return convertInlineElementsToHtml_(item, images); })
        .map(function(itemHtml) { return normalizeString_(itemHtml); })
        .filter(function(itemHtml) { return itemHtml !== ''; })
        .map(function(itemHtml) { return '<li>' + itemHtml + '</li>'; })
        .join('');

      if (listHtml) {
        chunks.push('<ul>' + listHtml + '</ul>');
      }
      continue;
    }
  }

  return {
    html: chunks.join('\n'),
    images: images,
  };
}

function convertParagraphToHtml_(paragraph, imageCollector) {
  const textHtml = normalizeString_(convertInlineElementsToHtml_(paragraph, imageCollector));
  if (!textHtml) {
    return '';
  }

  const headingTag = headingToTag_(paragraph.getHeading());
  if (headingTag) {
    return '<' + headingTag + '>' + textHtml + '</' + headingTag + '>';
  }

  return '<p>' + textHtml + '</p>';
}

function headingToTag_(heading) {
  if (heading === DocumentApp.ParagraphHeading.HEADING1) return 'h1';
  if (heading === DocumentApp.ParagraphHeading.HEADING2) return 'h2';
  if (heading === DocumentApp.ParagraphHeading.HEADING3) return 'h3';
  if (heading === DocumentApp.ParagraphHeading.HEADING4) return 'h4';
  if (heading === DocumentApp.ParagraphHeading.HEADING5) return 'h5';
  if (heading === DocumentApp.ParagraphHeading.HEADING6) return 'h6';
  return '';
}

function convertInlineElementsToHtml_(container, imageCollector) {
  const parts = [];
  const total = container.getNumChildren();

  for (let i = 0; i < total; i += 1) {
    const child = container.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      const inlineImage = child.asInlineImage();
      const token = createImagePlaceholderToken_(imageCollector.length);
      const altDescription = normalizeString_(inlineImage.getAltDescription());
      imageCollector.push({
        token: token,
        blob: inlineImage.getBlob(),
        alt: altDescription,
      });
      parts.push(token);
      continue;
    }

    if (type !== DocumentApp.ElementType.TEXT) {
      continue;
    }

    const textElement = child.asText();
    const text = textElement.getText();
    if (!text) {
      continue;
    }

    const indices = textElement.getTextAttributeIndices();
    for (let j = 0; j < indices.length; j += 1) {
      const start = indices[j];
      const end = j + 1 < indices.length ? indices[j + 1] : text.length;
      // 段落内の改行（Shift+Enter など）は <br> にする。装飾処理は「1行 = 1ブロック」を前提にしているため、生の改行を残さない
      const escaped = sanitizeInlineText_(text.substring(start, end)).replace(/\r\n|[\r\n\u000b]/g, '<br>');
      const isBold = textElement.isBold(start) === true;
      parts.push(isBold ? '<strong>' + escaped + '</strong>' : escaped);
    }
  }

  return parts.join('');
}

function createImagePlaceholderToken_(index) {
  return '[[G_DOC_IMAGE_' + index + ']]';
}

/**
 * ドキュメント内の画像を WordPress メディアにアップロードし、プレースホルダーとURLの対応を返す。
 * - 同じ内容の画像は1回だけアップロードする
 * - ファイル名に内容ハッシュを含め、再入稿時はアップロード済みの画像を再利用する
 * - アップロードは fetchAll で並列に行う
 */
function uploadDocumentImages_(config, images, title) {
  if (!images || images.length === 0) {
    return [];
  }

  const uploadsByHash = {};
  const uploads = [];
  const hashes = images.map(function(image, index) {
    const bytes = image.blob.getBytes();
    const hash = computeHexDigest_(bytes).slice(0, MEDIA_HASH_LENGTH);
    if (!uploadsByHash[hash]) {
      const upload = createMediaUpload_(image.blob, bytes, hash, title + '-image-' + (index + 1));
      uploadsByHash[hash] = upload;
      uploads.push(upload);
    }
    return hash;
  });

  findUploadedMedia_(config, uploads);

  const pending = uploads.filter(function(upload) { return !upload.url; });
  const responses = fetchAllInBatches_(
    pending.map(function(upload) { return buildMediaUploadRequest_(config, upload); }),
    MEDIA_UPLOAD_BATCH_SIZE
  );
  responses.forEach(function(response, index) {
    const media = parseWpResponse_(config, response, 'WordPressメディアアップロードエラー');
    pending[index].url = media.source_url || '';
  });

  return images.map(function(image, index) {
    return {
      token: image.token,
      url: uploadsByHash[hashes[index]].url,
      alt: image.alt,
    };
  });
}

function createMediaUpload_(blob, bytes, hash, fallbackBaseName) {
  const contentType = normalizeString_(blob.getContentType()) || 'application/octet-stream';
  const nameParts = splitFileName_(normalizeString_(blob.getName()) || fallbackBaseName);
  const extension = contentTypeToExtension_(contentType) || nameParts.extension;
  return {
    hash: hash,
    bytes: bytes,
    contentType: contentType,
    filename: (nameParts.baseName + '-' + hash + extension).replace(/["\\\r\n]/g, ''),
    url: '',
  };
}

/**
 * ファイル名にハッシュを含むメディアを検索し、見つかったものはアップロードを省略する。
 * 検索は最適化のためのものなので、失敗した場合は通常どおりアップロードする。
 */
function findUploadedMedia_(config, uploads) {
  if (uploads.length === 0) {
    return;
  }

  const responses = fetchAllInBatches_(
    uploads.map(function(upload) {
      return buildWpRequest_(config, 'get', '/media', {
        query: { search: upload.hash, per_page: 10, _fields: 'id,source_url' },
      });
    }),
    FETCH_BATCH_SIZE
  );

  responses.forEach(function(response, index) {
    const code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      return;
    }
    const items = tryParseJson_(response.getContentText());
    if (!Array.isArray(items)) {
      return;
    }
    const hash = uploads[index].hash;
    const match = items.find(function(item) {
      return String(item.source_url || '').indexOf(hash) !== -1;
    });
    if (match) {
      uploads[index].url = match.source_url;
    }
  });
}

function buildMediaUploadRequest_(config, upload) {
  return {
    url: buildApiUrl_(config, '/media', { _fields: 'id,source_url' }),
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Authorization: config.authHeader,
      'Content-Disposition': 'attachment; filename="' + upload.filename + '"',
      'Content-Type': upload.contentType,
    },
    payload: upload.bytes,
  };
}

function splitFileName_(filename) {
  const match = String(filename).match(/^(.*?)(\.[A-Za-z0-9]{1,5})?$/);
  return {
    baseName: match[1] || 'image',
    extension: match[2] || '',
  };
}

function contentTypeToExtension_(contentType) {
  if (contentType === 'image/jpeg') return '.jpg';
  if (contentType === 'image/png') return '.png';
  if (contentType === 'image/gif') return '.gif';
  if (contentType === 'image/webp') return '.webp';
  return '';
}

function computeHexDigest_(value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, value);
  let hex = '';
  for (let i = 0; i < digest.length; i += 1) {
    hex += ('0' + (digest[i] & 0xff).toString(16)).slice(-2);
  }
  return hex;
}

function replaceImagePlaceholders_(html, uploadedImages) {
  const imagesByToken = {};
  uploadedImages.forEach(function(image) {
    imagesByToken[image.token] = image;
  });

  // プレースホルダーは必ず <p> / <li> / 見出しの中にあるため、<img> だけに置き換える（<p> の入れ子を作らない）
  return String(html).replace(IMAGE_PLACEHOLDER_PATTERN, function(token) {
    const image = imagesByToken[token];
    if (!image || !image.url) {
      return token;
    }
    return '<img src="' + escapeAttribute_(image.url) + '" alt="' + escapeAttribute_(image.alt || '') + '">';
  });
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

/**
 * カテゴリー / タグ名を ID に変換する（無ければ作成する）。
 * - 同じ実行内で一度解決した名前はキャッシュから返す
 * - 検索・作成は fetchAll で並列に行う
 */
function ensureTermIds_(config, taxonomy, names) {
  if (names.length === 0) {
    return [];
  }

  const cache = config.termCache[taxonomy] || (config.termCache[taxonomy] = {});
  const uniqueNames = names.filter(function(name, index) { return names.indexOf(name) === index; });
  const uncached = uniqueNames.filter(function(name) {
    return !Object.prototype.hasOwnProperty.call(cache, name);
  });

  if (uncached.length > 0) {
    const searchResponses = fetchAllInBatches_(
      uncached.map(function(name) {
        return buildWpRequest_(config, 'get', '/' + taxonomy, {
          query: { search: name, per_page: 100, _fields: 'id,name' },
        });
      }),
      FETCH_BATCH_SIZE
    );

    const notFound = [];
    searchResponses.forEach(function(response, index) {
      const name = uncached[index];
      const items = parseWpResponse_(config, response, 'WordPress APIエラー');
      const match = items.find(function(item) {
        return decodeHtmlEntities_(normalizeString_(item.name)) === name;
      });
      if (match) {
        cache[name] = match.id;
      } else {
        notFound.push(name);
      }
    });

    if (notFound.length > 0) {
      const createResponses = fetchAllInBatches_(
        notFound.map(function(name) {
          return buildWpRequest_(config, 'post', '/' + taxonomy, {
            query: { _fields: 'id' },
            payload: { name: name },
          });
        }),
        FETCH_BATCH_SIZE
      );

      createResponses.forEach(function(response, index) {
        const name = notFound[index];
        const body = tryParseJson_(response.getContentText());
        if (body && body.code === 'term_exists' && body.data && body.data.term_id) {
          // 検索で一致しなかったが同名のタームが既にある場合は、それを使う
          cache[name] = body.data.term_id;
          return;
        }
        cache[name] = parseWpResponse_(config, response, 'WordPress APIエラー').id;
      });
    }
  }

  return uniqueNames.map(function(name) { return cache[name]; });
}

function upsertPost_(config, postType, payload, postId) {
  const route = '/' + postType.restBase + (postId ? '/' + encodeURIComponent(postId) : '');
  const response = wpFetch_(config, 'post', route, {
    namespace: postType.restNamespace,
    // 本文を含む大きなレスポンスを受け取らないよう、必要な項目だけに絞る
    query: { _fields: 'id,link' },
    payload: payload,
  });

  if (postId && response.getResponseCode() === 404) {
    const error = createWpError_(config, response, 'WordPress APIエラー');
    error.message += ' / 補足: H列の WP_POST_ID（' + postId + '）が、C列の投稿タイプ「' +
      postType.slug + '」の投稿として存在するか確認してください。';
    throw error;
  }

  return parseWpResponse_(config, response, 'WordPress APIエラー');
}

function stripHtml_(text) {
  return String(text)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const BOX_DEFS = {
  POINT: { icon: '✅', title: 'ポイント・メリット', className: 'swell-box-point' },
  CAUTION: { icon: '⚠️', title: '注意点・デメリット', className: 'swell-box-caution' },
  NOTE: { icon: '💡', title: '補足説明・備考', className: 'swell-box-note' },
  SUMMARY: { icon: '📝', title: '要約・まとめ', className: 'swell-box-summary' },
};
const DEFAULT_CAPTION_TITLE = '補足情報';
// 段落の先頭にある開始タグ（[POINT] など / [CAPTION:タイトル]）。太字の開始タグや改行が前にあっても検出する
const BOX_OPEN_PATTERN = /^((?:<strong>|<br>|\s)*)\[(?:(POINT|CAUTION|NOTE|SUMMARY)|CAPTION(?::([^\]]*))?)\]/;

/**
 * 本文HTMLに装飾を適用する。
 * 本文は convertGoogleDocBodyToHtml_ が作る「1行 = 1ブロック（<p> / <hN> / <ul>）」の形式なので、
 * 装飾もブロック単位で行い、出力も1ブロック1行の形式を保つ。
 */
function applyDecorations_(content) {
  let decorated = normalizeString_(content);
  if (!decorated) {
    return decorated;
  }

  decorated = applyMarkerDecoration_(decorated);
  decorated = applyBoxDecorations_(decorated);
  decorated = applyFaqDecoration_(decorated);

  return decorated;
}

/**
 * ==テキスト== を <mark> に変換する。段落をまたいだり、<img> などのタグ内を書き換えたりしないよう、
 * 1ブロック（1行）の中で <strong> 以外のタグを避けて変換する。
 */
function applyMarkerDecoration_(html) {
  return html.split('\n').map(function(line) {
    const tags = [];
    const masked = line.replace(/<(?!\/?strong>)[^>]*>/g, function(tag) {
      tags.push(tag);
      return '\u0000' + (tags.length - 1) + '\u0001';
    });
    const marked = masked.replace(/==([^\u0000\u0001]+?)==/g, function(_, text) {
      return '<mark>' + text.trim() + '</mark>';
    });
    return marked.replace(/\u0000(\d+)\u0001/g, function(_, index) {
      return tags[Number(index)];
    });
  }).join('\n');
}

/**
 * [POINT] / [CAUTION] / [NOTE] / [SUMMARY] / [CAPTION:タイトル] 〜 [/…] をボックスに変換する。
 * 開始タグは段落の先頭、終了タグは段落の末尾に書く。1段落の中でも、複数の段落（箇条書き・見出し・画像を含む）にまたがってもよい。
 */
function applyBoxDecorations_(html) {
  const blocks = html.split('\n');
  const output = [];
  let box = null;

  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    let inner = getParagraphInner_(block);

    if (!box) {
      const open = inner === null ? null : inner.match(BOX_OPEN_PATTERN);
      if (!open) {
        output.push(block);
        continue;
      }
      box = { type: open[2] || 'CAPTION', captionTitle: open[3] || '', body: [], startIndex: i };
      inner = open[1] + inner.slice(open[0].length);
    }

    if (inner === null) {
      box.body.push(block);
      continue;
    }

    let close = inner.match(new RegExp('\\[/' + box.type + '\\]((?:</strong>|<br>|\\s)*)$'));
    if (close && inner.slice(0, close.index).indexOf('[' + box.type) !== -1) {
      // 「前文[POINT]…[/POINT]」のように、同じ段落内の別の開始タグと対になる終了タグでは閉じない
      close = null;
    }
    if (close) {
      inner = inner.slice(0, close.index) + close[1];
    }
    pushParagraph_(box.body, inner);
    if (close) {
      output.push(renderBox_(box));
      box = null;
    }
  }

  if (box) {
    // 終了タグが無い開始タグは変換せずに残し、その後ろの段落はあらためて変換する
    output.push(blocks[box.startIndex]);
    output.push(applyBoxDecorations_(blocks.slice(box.startIndex + 1).join('\n')));
  }

  return output.filter(function(block) { return block !== ''; }).join('\n');
}

function renderBox_(box) {
  const body = box.body.join('');
  if (box.type === 'CAPTION') {
    const title = normalizeString_(stripTags_(box.captionTitle)) || DEFAULT_CAPTION_TITLE;
    return '<div class="caption-box">' +
      '<p class="caption-box__title">' + title + '</p>' +
      '<div class="caption-box__body">' + body + '</div>' +
      '</div>';
  }

  const def = BOX_DEFS[box.type];
  return '<div class="swell-box ' + def.className + '">' +
    '<p class="swell-box__title">' + def.icon + ' ' + def.title + '</p>' +
    '<div class="swell-box__body">' + body + '</div>' +
    '</div>';
}

/**
 * 「よくある質問」見出し（H1〜H6、または「## よくある質問」と書いた段落）から、
 * 次の同じレベル以上の見出しまでにある Q: / A: をFAQブロックに変換する。
 */
function applyFaqDecoration_(html) {
  const blocks = html.split('\n');
  const output = [];
  let i = 0;

  while (i < blocks.length) {
    const heading = getHeading_(blocks[i]);
    if (!heading || !/^よくある質問/.test(normalizeString_(stripTags_(heading.inner)))) {
      output.push(blocks[i]);
      i += 1;
      continue;
    }

    let end = i + 1;
    while (end < blocks.length) {
      const next = getHeading_(blocks[end]);
      if (next && next.level <= heading.level) {
        break;
      }
      end += 1;
    }

    const faqBlocks = buildFaqBlocks_(blocks.slice(i + 1, end));
    if (faqBlocks) {
      output.push('<h' + heading.level + '>' + heading.inner + '</h' + heading.level + '>');
      Array.prototype.push.apply(output, faqBlocks);
    } else {
      Array.prototype.push.apply(output, blocks.slice(i, end));
    }
    i = end;
  }

  return output.join('\n');
}

/**
 * FAQセクション内のブロックを変換する。Q/A が1組も無ければ null を返す。
 * - Q: の段落から A: の前までを質問、A: から次の Q: の前までを回答とする
 * - 最初の Q: より前の段落や、A: の無い Q: はそのまま残す
 */
function buildFaqBlocks_(blocks) {
  const items = [];
  let entry = null;
  let entryCount = 0;

  function closeEntry() {
    if (!entry) {
      return;
    }
    if (entry.answered) {
      items.push({ entry: entry });
      entryCount += 1;
    } else {
      items.push({ raw: entry.source });
    }
    entry = null;
  }

  blocks.forEach(function(block) {
    const qa = getQaInner_(block);
    if (qa && qa.kind === 'Q') {
      closeEntry();
      entry = { question: [qa.inner], questionExtra: [], answer: [], answerExtra: [], answered: false, source: [block] };
      return;
    }
    if (!entry) {
      items.push({ raw: [block] });
      return;
    }

    entry.source.push(block);
    if (qa && qa.kind === 'A' && !entry.answered) {
      entry.answered = true;
      entry.answer.push(qa.inner);
      return;
    }

    // 続きの段落は同じ Q / A の <p> 内に改行でつなぎ、箇条書きなどの段落以外のブロックはその後ろに置く
    const inner = getParagraphInner_(block);
    if (!entry.answered) {
      if (inner !== null && entry.questionExtra.length === 0) {
        entry.question.push(inner);
      } else {
        entry.questionExtra.push(block);
      }
    } else if (inner !== null && entry.answerExtra.length === 0) {
      entry.answer.push(inner);
    } else {
      entry.answerExtra.push(block);
    }
  });
  closeEntry();

  if (entryCount === 0) {
    return null;
  }

  const output = [];
  let group = [];
  function flushGroup() {
    if (group.length > 0) {
      output.push('<div class="faq-block">' + group.join('') + '</div>');
      group = [];
    }
  }
  items.forEach(function(item) {
    if (item.entry) {
      group.push(renderFaqItem_(item.entry));
      return;
    }
    flushGroup();
    Array.prototype.push.apply(output, item.raw);
  });
  flushGroup();

  return output;
}

function renderFaqItem_(entry) {
  const joinLines = function(parts) {
    return parts.map(cleanInlineHtml_).filter(function(part) { return part !== ''; }).join('<br>');
  };
  return '<div class="faq-block__item">' +
    '<p class="faq-block__q">Q. ' + joinLines(entry.question) + '</p>' + entry.questionExtra.join('') +
    '<p class="faq-block__a">A. ' + joinLines(entry.answer) + '</p>' + entry.answerExtra.join('') +
    '</div>';
}

/**
 * 段落または見出しが Q: / A:（全角の Ｑ・Ａ・：も可）で始まる場合、その種類と、接頭辞を除いた中身を返す。
 */
function getQaInner_(block) {
  let inner = getParagraphInner_(block);
  if (inner === null) {
    const heading = block.match(/^<h([1-6])>([\s\S]*)<\/h\1>$/);
    if (!heading) {
      return null;
    }
    inner = heading[2];
  }

  const match = normalizeString_(stripTags_(inner)).match(/^([QＱAＡ])[:：]/);
  if (!match) {
    return null;
  }
  return {
    kind: match[1] === 'Q' || match[1] === 'Ｑ' ? 'Q' : 'A',
    inner: inner.replace(/^((?:<[^>]+>|\s)*)[QＱAＡ][:：]\s*/, '$1'),
  };
}

/**
 * 見出しブロックのレベルと中身を返す。<hN> のほか、「## 見出し」のように書いた段落も見出しとして扱う。
 */
function getHeading_(block) {
  const heading = block.match(/^<h([1-6])>([\s\S]*)<\/h\1>$/);
  if (heading) {
    return { level: Number(heading[1]), inner: heading[2] };
  }

  const inner = getParagraphInner_(block);
  if (inner === null) {
    return null;
  }
  const markdown = normalizeString_(stripTags_(inner)).match(/^(#{1,6})(?:\s+|(?=よくある質問))/);
  if (!markdown) {
    return null;
  }
  return {
    level: markdown[1].length,
    inner: cleanInlineHtml_(inner.replace(/^((?:<[^>]+>|\s)*)#{1,6}\s*/, '$1')),
  };
}

function getParagraphInner_(block) {
  const match = block.match(/^<p>([\s\S]*)<\/p>$/);
  return match ? match[1] : null;
}

function pushParagraph_(blocks, inner) {
  const cleaned = cleanInlineHtml_(inner);
  if (cleaned) {
    blocks.push('<p>' + cleaned + '</p>');
  }
}

/**
 * 装飾タグを取り除いた後に残る空の <strong> / <mark> と、前後の空白・<br> を取り除く。
 */
function cleanInlineHtml_(inner) {
  return String(inner)
    .replace(/<(strong|mark)>\s*<\/\1>/g, '')
    .replace(/^(?:\s|<br>)+|(?:\s|<br>)+$/g, '');
}

function stripTags_(html) {
  return String(html).replace(/<[^>]*>/g, '');
}

function sanitizeInlineText_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeAttribute_(text) {
  return sanitizeInlineText_(text).replace(/"/g, '&quot;').replace(/[\r\n\u000b]+/g, ' ');
}

function decodeHtmlEntities_(text) {
  return String(text)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, '\'')
    .replace(/&amp;/g, '&');
}
