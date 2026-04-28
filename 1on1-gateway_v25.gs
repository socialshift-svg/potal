// ============================================================
// 1on1 Management System - GAS Central Gateway
// n8nの代替として全APIをGASに集約
// F-11/F-12（Drive Push通知）のみn8n継続
// ============================================================

// ── 設定 ──────────────────────────────────────────────────
// ★ セキュリティ: 機密値は全てScriptPropertiesから取得
// GASエディタ → プロジェクトの設定 → スクリプトプロパティ に以下を登録:
//   JWT_SECRET      : ランダムな32文字以上の文字列
//   GAS_SECRET      : ランダムな32文字以上の文字列
//   ADMIN_EMAIL     : 管理者メールアドレス
//   SPREADSHEET_ID  : スプレッドシートID
// 設定後、下記 _PROPS._init() が自動で読み込む

var _PROPS = {
  _cache: null,
  _init: function() {
    if (this._cache) return this._cache;
    var p = PropertiesService.getScriptProperties().getProperties();
    this._cache = p;
    return p;
  },
  get: function(key) { return this._init()[key] || ''; }
};

const CONFIG = {
  get SPREADSHEET_ID()           { return _PROPS.get('SPREADSHEET_ID')           || '1aOKg0jQZptb-2iKJoOTDSe5MCefDwO-axDJ_H7KO_QM'; },
  get JWT_SECRET()               { return _PROPS.get('JWT_SECRET')               || ''; },
  get ADMIN_EMAIL()              { return _PROPS.get('ADMIN_EMAIL')              || 'kohei.umeda@agent-network.com'; },
  get GAS_SECRET()               { return _PROPS.get('GAS_SECRET')               || ''; },
  JWT_EXPIRES_HOURS:        24,
  SHARED_DRIVE_ID:          '1xduD0ziGDXz64z6Iva29k65Lvgl48jPl',
  RECORDINGS_FOLDER_ID:     '1sCxZ7TFkGbAzfHZHF8k0LfqGmCqnjygs',
  // ★ test.admin@socialshift.work のマイドライブ「Meet Recordings」フォルダ
  MEET_RECORDINGS_FOLDER_ID:'1ZE3fyf2hd4dmzm5tUEO6vna8Q0zaW2th',
  // ★ 01_個人フォルダ（文字起こし・録音ファイルの格納先）
  INDIVIDUAL_FOLDER_ROOT_ID:'1sNZfZqgyMq9ZA7RZoS1tRVsZ1UgqDTxv',
  // ★ PERSONAL_FOLDER_ROOT_ID は INDIVIDUAL_FOLDER_ROOT_ID と同一（統一）
  PERSONAL_FOLDER_ROOT_ID:  '1sNZfZqgyMq9ZA7RZoS1tRVsZ1UgqDTxv',
};

// ── CORS ヘッダー ─────────────────────────────────────────
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,PATCH,PUT,DELETE,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Content-Type': 'application/json'
  };
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg, code) {
  return jsonResponse({ error: msg, code: code || 400 });
}

// GAS WebAppはCORSヘッダーを付与できないため、このラッパーはそのまま返すのみ（意図的）
// CORS対応はGASデプロイ設定「全員がアクセス可能」で行う
function addCorsHeaders(output) {
  return output;
}

// ── Spreadsheet ヘルパー ──────────────────────────────────
// ============================================================
// 日時ユーティリティ（JST統一）
// GASはUTCで動作するため、メール等の表示は全てJSTに変換する
// ============================================================

/**
 * DateオブジェクトまたはISO文字列をJST日時文字列に変換
 * @param {Date|string} dateOrStr
 * @param {string} format - 'full': 'YYYY/M/D HH:MM' / 'date': 'YYYY/M/D' / 'short': 'M/D HH:MM'
 * @return {string}
 */
function toJST_(dateOrStr, format) {
  var d = (dateOrStr instanceof Date) ? dateOrStr : new Date(dateOrStr);
  if (isNaN(d)) return String(dateOrStr || '—');
  // JSTはUTC+9
  var jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
  var Y   = jst.getUTCFullYear();
  var M   = jst.getUTCMonth() + 1;
  var D   = jst.getUTCDate();
  var h   = ('0' + jst.getUTCHours()).slice(-2);
  var m   = ('0' + jst.getUTCMinutes()).slice(-2);
  var wd  = ['日','月','火','水','木','金','土'][jst.getUTCDay()];
  if (format === 'date')  return Y + '/' + M + '/' + D;
  if (format === 'short') return M + '/' + D + ' (' + wd + ') ' + h + ':' + m;
  if (format === 'label') return M + '月' + D + '日（' + wd + '） ' + h + ':' + m;
  return Y + '/' + M + '/' + D + ' (' + wd + ') ' + h + ':' + m; // full（デフォルト）
}

/**
 * 「明日」の判定をJSTベースで行う
 * GASのnew Date()はUTCなので+9hでJST今日を算出
 */
function getTomorrowJST_() {
  var nowJST = new Date(Date.now() + 9 * 60 * 60 * 1000);
  // JSTの明日の日付文字列 (YYYY-MM-DD)
  var tomorrow = new Date(nowJST);
  tomorrow.setUTCDate(nowJST.getUTCDate() + 1);
  return tomorrow.getUTCFullYear() + '-'
    + ('0' + (tomorrow.getUTCMonth() + 1)).slice(-2) + '-'
    + ('0' + tomorrow.getUTCDate()).slice(-2);
}

/**
 * 今週・今月の判定用（JSTベース）
 */
function getNowJST_() {
  return new Date(Date.now() + 9 * 60 * 60 * 1000);
}


// SpreadsheetAppをリクエスト内でキャッシュ（openIdが重い呼び出しのため）
var _ss = null;
function getSpreadsheet_() {
  if (!_ss) _ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return _ss;
}
function getSheet(sheetName) {
  return getSpreadsheet_().getSheetByName(sheetName);
}

// ── CacheServiceによるシートデータキャッシュ ──────────────
// TTL: users/schedules=300秒, bookings=60秒, その他=120秒
var CACHE_TTL = {
  users:              300,
  mentor_schedules:   300,
  quick_links:        300,
  contents:           300,
  bookings:            60,
  leader_assignments: 120,
  call_reports:        60,
  pre_reports:         60,
  admin_memos:         60,
  mentor_reports:      60,
  surveys:             30, // ★ アンケートは短めにキャッシュ（回答後に即反映）
};

function getCachedSheet_(sheetName) {
  var ttl = CACHE_TTL[sheetName];
  if (!ttl) return null;

  var cache = CacheService.getScriptCache();
  var key   = 'sheet_' + sheetName;
  try {
    var hit = cache.get(key);
    if (!hit) return null;
    // チャンク分割されている場合はメタ情報から組み立て
    var meta = JSON.parse(hit);
    if (meta && meta._chunks) {
      var parts = [];
      for (var i = 0; i < meta._chunks; i++) {
        var chunk = cache.get(key + '_c' + i);
        if (!chunk) return null; // チャンク欠損はキャッシュミス扱い
        parts.push(chunk);
      }
      return JSON.parse(parts.join(''));
    }
    return meta; // チャンクなし（従来通り）
  } catch(e) { return null; }
}

function setCachedSheet_(sheetName, data) {
  var ttl = CACHE_TTL[sheetName];
  if (!ttl) return;
  var cache = CacheService.getScriptCache();
  var key   = 'sheet_' + sheetName;
  try {
    var json = JSON.stringify(data);
    if (json.length < 90000) {
      // 90KB未満はそのまま1エントリに格納
      cache.put(key, json, ttl);
    } else {
      // 90KB以上はチャンク分割して格納（各チャンク最大90000文字）
      var CHUNK = 90000;
      var chunks = [];
      for (var i = 0; i < json.length; i += CHUNK) {
        chunks.push(json.slice(i, i + CHUNK));
      }
      // チャンクを個別保存
      for (var j = 0; j < chunks.length; j++) {
        cache.put(key + '_c' + j, chunks[j], ttl);
      }
      // メタ情報（チャンク数）を保存
      cache.put(key, JSON.stringify({ _chunks: chunks.length }), ttl);
      Logger.log('setCachedSheet_: ' + sheetName + ' → ' + chunks.length + 'チャンク分割 (' + json.length + 'bytes)');
    }
  } catch(e) { Logger.log('setCachedSheet_ error: ' + e.message); }
}

// キャッシュを明示的に破棄（書き込み後に呼ぶ）
function invalidateCache_(sheetName) {
  var cache = CacheService.getScriptCache();
  var key   = 'sheet_' + sheetName;
  try {
    // チャンク分割キャッシュも合わせて削除
    var meta = cache.get(key);
    if (meta) {
      try {
        var m = JSON.parse(meta);
        if (m && m._chunks) {
          for (var i = 0; i < m._chunks; i++) {
            cache.remove(key + '_c' + i);
          }
        }
      } catch(pe) {}
    }
    cache.remove(key);
  } catch(e) {}
}

// キャッシュ対応版 sheetToObjects
function cachedSheetToObjects_(sheetName) {
  var cached = getCachedSheet_(sheetName);
  if (cached) return cached;
  var data = sheetToObjects(getSheet(sheetName));
  setCachedSheet_(sheetName, data);
  return data;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v === true)  { obj[h] = 'TRUE';  return; }
      if (v === false) { obj[h] = 'FALSE'; return; }
      // ★ Date型はISO文字列に変換（スプシが日付型として保存した場合の対策）
      if (v instanceof Date) { obj[h] = isNaN(v.getTime()) ? '' : v.toISOString(); return; }
      obj[h] = v === undefined ? '' : String(v);
    });
    return obj;
  });
}

function appendRow(sheetName, obj) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet   = getSheet(sheetName);
    if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row     = headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
    sheet.appendRow(row);
    invalidateCache_(sheetName); // キャッシュ破棄
  } finally {
    lock.releaseLock();
  }
}

function updateRowWhere(sheetName, matchCol, matchVal, updates) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet(sheetName);
    if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx  = headers.indexOf(matchCol);
    if (colIdx < 0) throw new Error('Column not found: ' + matchCol);

    // ★ 更新しようとしているキーのうちシートに存在しない列をログ出力
    var missingCols = Object.keys(updates).filter(function(k){ return headers.indexOf(k) < 0; });
    if (missingCols.length > 0) {
      Logger.log('updateRowWhere[' + sheetName + ']: 以下の列がシートに存在しないためスキップ: ' + missingCols.join(', '));
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIdx]) === String(matchVal)) {
        // ★ 行データをヘッダー長に合わせてパディング（列追加後の不一致を防ぐ）
        while (data[i].length < headers.length) { data[i].push(''); }
        Object.keys(updates).forEach(function(key) {
          var idx = headers.indexOf(key);
          if (idx >= 0) data[i][idx] = updates[key];
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
        Logger.log('updateRowWhere[' + sheetName + ']: 行' + (i+1) + ' 更新完了');
      }
    }
    invalidateCache_(sheetName);
  } catch(err) {
    Logger.log('updateRowWhere[' + sheetName + '] エラー: ' + err.message);
    throw err; // 再スロー（呼び出し元でエラーを把握できるよう）
  } finally {
    lock.releaseLock();
  }
}

// ── JWT (HMAC-SHA256 簡易実装) ────────────────────────────
function base64urlEncode(str) {
  return Utilities.base64EncodeWebSafe(str).replace(/=+$/, '');
}

function base64urlDecode(str) {
  var pad = str.length % 4;
  if (pad) str += '===='.slice(0, 4 - pad);
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(str)).getDataAsString();
}

function createJWT(payload) {
  var secret = CONFIG.JWT_SECRET;
  // ★ セキュリティ: JWT_SECRETが未設定の場合はトークン発行を拒否
  if (!secret) throw new Error('JWT_SECRET が設定されていません。ScriptPropertiesを確認してください。');
  var header = base64urlEncode(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  var now = Math.floor(Date.now() / 1000);
  payload.iat = now;
  payload.exp = now + CONFIG.JWT_EXPIRES_HOURS * 3600;
  var body = base64urlEncode(JSON.stringify(payload));
  var sigInput = header + '.' + body;
  var sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(sigInput, secret)
  ).replace(/=+$/, '');
  return header + '.' + body + '.' + sig;
}

function verifyJWT(token) {
  try {
    var secret = CONFIG.JWT_SECRET;
    if (!secret) return null; // JWT_SECRET未設定時は全トークン無効
    var parts = token.split('.');
    if (parts.length !== 3) return null;
    var payload = JSON.parse(base64urlDecode(parts[1]));
    if (payload.exp && Math.floor(Date.now() / 1000) > payload.exp) return null;
    var sigInput = parts[0] + '.' + parts[1];
    var expectedSig = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(sigInput, secret)
    ).replace(/=+$/, '');
    if (expectedSig !== parts[2]) return null;
    return payload;
  } catch(e) {
    return null;
  }
}

function decodeJWTNoVerify(token) {
  try {
    var parts = token.split('.');
    if (parts.length !== 3) return null;
    return JSON.parse(base64urlDecode(parts[1]));
  } catch(e) { return null; }
}

function getTokenFromRequest(e) {
  var auth = (e.parameter && e.parameter.Authorization) || '';
  if (!auth && e.postData) {
    try {
      // ★ 高速化: doPostで既にパース済みのbodyを再利用
      var body = e._parsedBody || JSON.parse(e.postData.contents || '{}');
      if (body._token) return body._token.replace(/^Bearer\s+/i, '').trim();
      auth = body._auth || body.token || '';
    } catch(err) {}
  }
  return auth.replace(/^Bearer\s+/i, '').trim();
}

function requireAuth(e, requiredRole) {
  var token = getTokenFromRequest(e);
  if (!token) return { error: 'TOKEN_MISSING', status: 401 };
  var payload = verifyJWT(token);
  // ★ セキュリティ修正: 検証失敗時はフォールバックせず即座にエラー
  // decodeJWTNoVerify へのフォールバックは署名改ざん・期限切れトークンを許容するため削除
  if (!payload) return { error: 'INVALID_TOKEN', status: 401 };
  if (requiredRole && payload.role !== requiredRole && payload.role !== 'admin') {
    return { error: 'FORBIDDEN', status: 403 };
  }
  return { payload: payload };
}

// ── SHA256 パスワードハッシュ ─────────────────────────────
// ※ 既存ログイン用（sha256: プレフィックス付き）
function sha256Hash(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str,
    Utilities.Charset.UTF_8
  );
  return 'sha256:' + bytes.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
}

// ※ user-management.html 送信値と一致させるための変換
// フロント: crypto.subtle で生成した 64桁の hex 文字列（プレフィックスなし）
// GAS内保存: "sha256:" + hex に統一して保存する
function normalizePasswordHash(rawHash) {
  // すでに "sha256:" 始まりなら何もしない
  if (typeof rawHash === 'string' && rawHash.indexOf('sha256:') === 0) return rawHash;
  // プレフィックスなし 64文字 hex → 付与して返す
  if (typeof rawHash === 'string' && /^[0-9a-f]{64}$/i.test(rawHash)) {
    return 'sha256:' + rawHash.toLowerCase();
  }
  // それ以外はそのまま（エラーは呼び出し元で処理）
  return rawHash || '';
}

// ── メール送信 ────────────────────────────────────────────
// ★ セキュリティ: メール送信先アドレスを検証してヘッダーインジェクションを防止
function sendMail(to, subject, htmlBody) {
  if (!to) return;
  // メールアドレス形式の簡易検証
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(to))) {
    Logger.log('sendMail: 無効なメールアドレスをスキップ: ' + to);
    return;
  }
  // subjectにヘッダーインジェクション文字が含まれていれば除去
  var safeSubject = String(subject || '').replace(/[\r\n]/g, ' ');
  GmailApp.sendEmail(to, safeSubject, '', { htmlBody: htmlBody, name: '1on1管理システム' });
}

// ── htmlEscape_: HTMLメール本文内でユーザー入力をエスケープ ──
function htmlEscape_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ── parseBody_: doPostで既にパース済みのbodyを再利用（二重パース防止）──
function parseBody_(e) {
  if (e && e._parsedBody) return e._parsedBody;
  try { return JSON.parse((e.postData && e.postData.contents) || '{}'); } catch(err) { return {}; }
}

// ── validateEmail_: メールアドレス形式検証 ──
function validateEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(email || ''));
}

// ── calcTenure_: 入社日から在籍期間を計算して「X年Yヶ月」形式で返す ──
// hire_date: 'YYYY-MM-DD' 形式の文字列
// 返値: { months: Number, label: String }  例: { months: 14, label: '1年2ヶ月' }
function calcTenure_(hireDate) {
  if (!hireDate) return { months: 0, label: '—' };
  var hire = new Date(hireDate);
  if (isNaN(hire.getTime())) return { months: 0, label: '—' };
  var nowJST = getNowJST_();
  var totalMonths = (nowJST.getUTCFullYear() - hire.getFullYear()) * 12
    + (nowJST.getUTCMonth() - hire.getMonth());
  if (nowJST.getUTCDate() < hire.getDate()) totalMonths--; // 日付補正
  if (totalMonths < 0) totalMonths = 0;
  var years  = Math.floor(totalMonths / 12);
  var months = totalMonths % 12;
  var label;
  if (years === 0) {
    label = months + 'ヶ月';
  } else if (months === 0) {
    label = years + '年';
  } else {
    label = years + '年' + months + 'ヶ月';
  }
  return { months: totalMonths, label: label };
}

// ============================================================
// doGet / doPost エントリーポイント
// ============================================================
function doGet(e) {
  try {
    var path = (e.parameter && e.parameter.path) || '';
    var output = routeRequest('GET', path, e);
    return addCorsHeaders(output);
  } catch(err) {
    return addCorsHeaders(jsonResponse({ error: err.message }, 500));
  }
}

function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents || '{}'); } catch(err) {}

    // パース済みbodyをeに付与してハンドラ内の再パースを削減
    e._parsedBody = body;

    // secret 認証ルーティング（F-11/F-12 + user-management API）
    if (body.secret) {
      var SECRET = PropertiesService.getScriptProperties().getProperty('N8N_SECRET') || CONFIG.GAS_SECRET;
      if (!SECRET || body.secret !== SECRET) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unauthorized' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var action = body.action || 'generateText';

      // ── n8n 連携アクション ──
      if (action === 'saveCallReport')    return jsonResponse(handleN8nSaveCallReport(body));
      if (action === 'saveMeetRecording') return jsonResponse(handleN8nSaveMeetRecording(body));

      // ── 既存 Drive/録音処理 ──
      if (action === 'organizeRecording') {
        return ContentService.createTextOutput(JSON.stringify(organizeRecording(body)))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (action === 'getFileContent') {
        return ContentService.createTextOutput(JSON.stringify(getFileContent(body)))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // generateText (Gemini)
      return ContentService.createTextOutput(JSON.stringify(generateTextGemini(body.prompt)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 通常 API ルーティング
    var path   = (e.parameter && e.parameter.path) || body.path || '';
    var method = (e.parameter && e.parameter._method) || body._method || 'POST';
    return addCorsHeaders(routeRequest(method.toUpperCase(), path, e));
  } catch(err) {
    return addCorsHeaders(jsonResponse({ error: err.message, code: 500 }));
  }
}

// Gemini 呼び出し
function generateTextGemini(prompt) {
  try {
    // ★ n8n経由でGemini AIを呼び出す
    // n8n Webhook → Secret認証 → Gemini AI → テキスト抽出（{ text }）→ 成功レスポンス
    var n8nUrl = 'https://umeda-0628.app.n8n.cloud/webhook/ai-summary';
    var payload = {
      secret: 'ai-proxy-secret-2026',
      prompt: prompt
    };
    var response = UrlFetchApp.fetch(n8nUrl, {
      method:          'post',
      contentType:     'application/json',
      payload:         JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('n8n Gemini エラー: HTTP ' + code + ' / ' + response.getContentText());
      return { success: false, error: 'HTTP ' + code };
    }
    var result = JSON.parse(response.getContentText());
    // n8nの「テキスト抽出」ノードが { text } を返す
    var text = result.text || '';
    if (!text) {
      Logger.log('n8n Gemini 空レスポンス: ' + response.getContentText());
      return { success: false, error: 'EMPTY_RESPONSE' };
    }
    return { success: true, text: text };
  } catch(err) {
    Logger.log('generateTextGemini エラー: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// ルートテーブル（O(1)ルーティング）
// if-elseチェーンの代わりにMapでルックアップ → パース時間を大幅削減
// ============================================================
var _ROUTE_TABLE = null;

function buildRouteTable_() {
  if (_ROUTE_TABLE) return _ROUTE_TABLE;
  _ROUTE_TABLE = {
    // ── Auth ──
    'POST auth/login':   handleLogin,
    'POST auth/verify':  handleVerify,
    'POST auth/refresh': handleRefresh,

    // ── Mentee ──
    'GET api/mentee/home':            handleMenteeHome,
    'GET api/mentee/notices':         handleMenteeNotices,
    'GET api/mentee/available-slots': handleAvailableSlots,
    'GET api/mentor/available-slots': handleMentorAvailableSlots,
    'GET api/mentee/reports':         handleMenteeReports,
    'POST api/mentee/pre-report':     handlePreReport,
    'POST api/mentee/pre-report/delete': handleDeletePreReport,      // ★ メンティー自身の削除
    'POST api/mentor/pre-report/delete': handleMentorDeletePreReport, // ★ メンターによる削除
    'PATCH api/mentee/profile':       handleUpdateProfile,
    'POST api/mentee/booking':        handleCreateBooking,
    'GET api/mentee/leader-info':     handleLeaderInfo,
    'GET api/mentee/links':           handleMenteeLinks,
    'GET api/mentee/contents':        handleMenteeContents,
    'GET api/mentee/my-chat-url':     handleMenteeChatUrl,
    'POST api/mentee/booking/cancel': handleCancelBooking,
    'POST api/mentee/tel-start':      handleTelStart,

    // ── Mentor ──
    'GET api/mentor/my-mentees':              handleMentorMyMentees,
    'GET api/mentor/members':                 handleMentorMembers,    // ★ 全メンバー一覧（mentor用）
    'GET api/mentor/my-bookings':             handleMentorMyBookings,
    'GET api/mentor/bookings':                handleMentorMyBookings,
    'POST api/mentor/update-status':          handleMentorUpdateStatus,
    'GET api/mentor/schedule':                handleGetSchedule,
    'POST api/mentor/schedule':               handleSaveSchedule,
    'POST api/mentor/meeting-completed':      handleMeetingCompleted,
    'POST api/mentor/reports/publish':        handlePublishReport,
    'POST api/mentor/booking/cancel':         handleCancelBooking,
    'GET api/mentor/mentee-reports':          handleMentorMenteeReports,
    'POST api/mentor/booking/create':         handleMentorCreateBooking,
    'POST api/mentor/booking/update':         handleUpdateBooking,
    'POST api/mentor/tel-start':              handleTelStart,
    'GET api/mentor/weekly-calls':            handleWeeklyCalls,
    'POST api/mentor/tel-ai-summary':         handleTelAiSummary,
    'GET api/mentor/call-reports':            handleCallReports,
    'POST api/mentor/call-reports':           handleCreateCallReport,
    'POST api/mentor/call-reports/confirm':   handleConfirmCallReport,
    'GET api/mentor/mentor-reports':          handleMentorReportsList,
    'POST api/mentor/mentor-reports/save':    handleSaveMentorReport,
    'POST api/mentor/mentor-reports/publish': handlePublishMentorReport,
    'POST api/mentor/mentor-reports/unpublish': handleUnpublishMentorReport, // ★ 下書きに戻す
    'POST api/mentor/mentor-reports/regenerate-ai': handleRegenerateAi, // ★ AI再生成（1on1）
    'POST api/leader/call-reports/regenerate-ai':   handleRegenerateTelAi, // ★ AI再生成（TEL）
    'GET api/mentor/pre-reports':             handleMentorPreReports,
    'GET api/mentor/memos':                   handleGetMentorMemos,
    'POST api/mentor/memos':                  handleSaveMentorMemo,
    'GET api/mentor/all-reports':             handleMentorAllReports,
    'POST api/mentor/set-mentee-goals':       handleMentorSetMenteeGoals,
    'GET api/mentor/contents':               handleMenteeContents,       // ★ 研修コンテンツ（menteeと共用）
    'GET api/mentor/surveys':                handleMentorSurveys,        // ★ メンター自身のアンケート履歴
    'GET api/mentee/pre-reports':            handleMenteePreReports,     // ★ メンティー自身の事前レポート一覧
    // ── サポート ──
    'GET api/support/list':                  handleSupportList,          // 一覧取得（全ログインユーザー）
    'POST api/support/detail':               handleSupportDetail,        // 詳細＋履歴取得（support_id をbodyで送る）
    'POST api/admin/support/update':         handleSupportUpdate,        // ステータス・対応内容更新（admin）
    'POST api/mentee/referral':              handleReferralSubmit,       // ★ リファラル候補者紹介

    // ── Admin ユーザー管理（JWT認証・user-management.html から呼ばれる）──
    'GET api/admin/users/list':            handleGetUsersJwt,
    'POST api/admin/users/create':         handleCreateUserJwt,
    'POST api/admin/users/update-profile': handleUpdateUserJwt,
    'POST api/admin/users/delete-user':    handleDeleteUserJwt,

    // ── Admin ──
    'GET api/admin/members':                   handleAdminMembers,
    'GET api/admin/pre-reports':               handleAdminPreReports,
    'GET api/admin/mentor-reports':            handleAdminMentorReports,
    'GET api/admin/call-reports':              handleAdminCallReports,
    'GET api/admin/users':                     handleAdminUsers,
    'POST api/admin/users/add':                handleAddUser,
    'GET api/admin/surveys':                   handleAdminSurveys,
    'GET api/admin/contents':                  handleAdminContents,
    'POST api/admin/contents':                 handleAddContent,
    'POST api/admin/contents/update':          handleUpdateContent,
    'POST api/admin/contents/delete':          handleDeleteContent,
    'POST api/admin/contents/bulk-import':     handleBulkImportContents,
    'GET api/admin/links':                     handleAdminLinks,
    'POST api/admin/links':                    handleAddLink,
    'POST api/admin/links/delete':             handleDeleteLink,
    'POST api/admin/notices':                  handleAdminNotice,
    'GET api/admin/notices/list':              handleAdminNoticeList,
    'POST api/admin/notices/delete':           handleDeleteNotice,
    'POST api/admin/notices/toggle':           handleToggleNotice,
    'GET api/admin/memos':                     handleGetAdminMemos,
    'POST api/admin/memos':                    handleSaveAdminMemo,
    'GET api/admin/leader-assignments':        handleGetAssignments,
    'POST api/admin/leader-assignments':       handleAddAssignment,
    'POST api/admin/leader-assignments/delete':handleDeleteAssignment,
    'GET api/admin/chat-webhook':              handleGetChatWebhook,
    'POST api/admin/chat-webhook/save':        handleSaveChatWebhook,
    'POST api/admin/chat-space/register':      handleRegisterChatSpace,
    'POST api/admin/create-chat-space':        handleCreateChatSpace,
    'GET api/admin/users/export':              handleAdminUsersExport,
    'POST api/admin/users/bulk-import':        handleAdminBulkImport,
    'POST api/admin/users/bulk-update':        handleAdminBulkUpdate,  // ★ 一括更新
    'POST api/admin/user/reset-password':      handleAdminResetPassword, // ★ パスワードリセット
    'POST api/admin/users/update':             handleAdminUpdateUser,
    'POST api/admin/set-tel-meet-url':         handleSetTelMeetUrl,
    'GET api/admin/tel-stock':                 handleGetTelStock,
    'POST api/admin/tel-stock/add':            handleAddTelStock,
    'POST api/admin/tel-stock/assign':         handleAssignTelStock,
    'POST api/admin/tel-stock/delete':         handleDeleteTelStock,
    'GET api/admin/ai-prompts':                handleGetAiPrompts,   // ★ AIプロンプト取得
    'POST api/admin/ai-prompts':               handleSaveAiPrompts,  // ★ AIプロンプト保存

    // ── Leader ──
    'GET api/leader/my-assignments': handleMyAssignments,

    // ── Analytics / Survey ──
    'POST api/analytics/page-view': handlePageView,
    'POST api/survey/submit':       handleSurveySubmit,
  };
  return _ROUTE_TABLE;
}

function routeRequest(method, path, e) {
  // ヘルスチェック（認証不要・最優先）
  if (path === 'health' || path === '') return handleHealth(e);

  // O(1)ルートルックアップ
  var key     = method + ' ' + path;
  var handler = buildRouteTable_()[key];
  if (handler) return handler(e);

  return errorResponse('Not Found: ' + method + ' ' + path, 404);
}

// ============================================================
// ヘルスチェック（死活監視・認証不要）
// GET /health → { ok, ts, version, sheetOk, cacheOk }
// ============================================================
function handleHealth(e) {
  var start   = Date.now();
  var sheetOk = false;
  var cacheOk = false;
  var version = '2.0.0'; // デプロイ時に更新

  try {
    var sheet = getSheet('users');
    sheetOk   = !!sheet;
  } catch(e) {}

  try {
    CacheService.getScriptCache().put('_health', '1', 10);
    cacheOk = CacheService.getScriptCache().get('_health') === '1';
  } catch(e) {}

  return jsonResponse({
    ok:      sheetOk && cacheOk,
    ts:      new Date().toISOString(),
    version: version,
    latency: Date.now() - start,
    sheetOk: sheetOk,
    cacheOk: cacheOk,
  });
}


// ============================================================
// F-Auth: 認証
// ============================================================

// ── ブルートフォース対策: ログイン失敗レート制限 ──
// CacheServiceで失敗回数を記録。5分で5回以上失敗したらロック
function checkLoginRateLimit_(email) {
  var cache = CacheService.getScriptCache();
  var key   = 'login_fail_' + email.replace(/[^a-z0-9]/g, '_').slice(0, 50);
  var data;
  try { data = JSON.parse(cache.get(key) || '{"count":0}'); } catch(e) { data = { count: 0 }; }
  return { count: data.count, key: key, cache: cache };
}

function recordLoginFailure_(email) {
  var r   = checkLoginRateLimit_(email);
  var cnt = (r.count || 0) + 1;
  try { r.cache.put(r.key, JSON.stringify({ count: cnt }), 300); } catch(e) {}
}

function clearLoginFailure_(email) {
  var r = checkLoginRateLimit_(email);
  try { r.cache.remove(r.key); } catch(e) {}
}

function handleLogin(e) {
  var body = parseBody_(e);

  var email = (body.email || '').trim().toLowerCase();

  // ★ セキュリティ: メールアドレス形式検証（不正入力を早期拒否）
  if (!email) return errorResponse('EMAIL_AND_PASSWORD_REQUIRED', 400);
  if (!validateEmail_(email)) return errorResponse('INVALID_EMAIL_FORMAT', 400);

  // ★ セキュリティ: ブルートフォース対策（5分で5回失敗でロック）
  var rateCheck = checkLoginRateLimit_(email);
  if (rateCheck.count >= 5) {
    return errorResponse('TOO_MANY_ATTEMPTS', 429);
  }

  // ★ セキュリティ強化: フロントでSHA-256ハッシュ化済みの値を受け取る
  // フロント(login.html)はcrypto.subtle で sha256Hex(password) してから送信
  // 後方互換: 旧クライアントが password をプレーンで送ってきた場合もGAS側でハッシュ化
  var rawHash     = body.password_hash || '';
  var rawPassword = body.password      || '';
  var inputHash;
  if (rawHash) {
    inputHash = normalizePasswordHash(rawHash);
  } else if (rawPassword) {
    inputHash = sha256Hash(rawPassword);
  } else {
    return errorResponse('EMAIL_AND_PASSWORD_REQUIRED', 400);
  }

  var users = cachedSheetToObjects_('users');
  var user = users.find(function(u) {
    return String(u.email || '').trim().toLowerCase() === email;
  });
  if (!user || !user.user_id) {
    recordLoginFailure_(email); // 失敗カウント増加
    return errorResponse('USER_NOT_FOUND', 401);
  }

  // status が deleted のユーザーはログイン不可
  if (String(user.status).toLowerCase() === 'deleted') {
    recordLoginFailure_(email);
    return errorResponse('USER_NOT_FOUND', 401);
  }

  if (user.password_hash !== inputHash) {
    recordLoginFailure_(email); // 失敗カウント増加
    return errorResponse('INVALID_PASSWORD', 401);
  }

  // ★ ログイン成功: 失敗カウントをリセット
  clearLoginFailure_(email);
  updateRowWhere('users', 'user_id', user.user_id, { updated_at: new Date().toISOString() });

  var token = createJWT({
    user_id: user.user_id,
    role: user.role,
    has_leader_role: user.has_leader_role === 'TRUE' || user.has_leader_role === 'true'
  });

  return jsonResponse({
    token: token,
    user: {
      user_id: user.user_id,
      name: user.name,
      email: user.email,
      role: user.role,
      has_leader_role: user.has_leader_role === 'TRUE' || user.has_leader_role === 'true'
    }
  });
}

function handleVerify(e) {
  var token = getTokenFromRequest(e);
  if (!token) return jsonResponse({ valid: false, error: 'NO_TOKEN' });
  var payload = verifyJWT(token);
  if (!payload) return jsonResponse({ valid: false, error: 'INVALID_OR_EXPIRED' });
  return jsonResponse({ valid: true, payload: payload });
}

function handleRefresh(e) {
  var token = getTokenFromRequest(e);
  if (!token) return jsonResponse({ valid: false, error: 'NO_TOKEN' });

  // ★ セキュリティ修正: まず正規検証を試みる。有効なトークンのみリフレッシュ可
  // 期限切れトークンは verifyJWT で弾く（decodeJWTNoVerify への無条件フォールバック廃止）
  var payload = verifyJWT(token);
  if (!payload) return jsonResponse({ valid: false, error: 'INVALID_OR_EXPIRED' });

  // DBでユーザーの存在・アクティブ状態を再確認
  var users = cachedSheetToObjects_('users');
  var user  = users.find(function(u) { return u.user_id === payload.user_id; });
  if (!user || !user.user_id || String(user.status).toLowerCase() === 'deleted') {
    return jsonResponse({ valid: false, error: 'USER_NOT_FOUND' });
  }

  var newToken = createJWT({
    user_id:        user.user_id,
    role:           user.role,
    has_leader_role: user.has_leader_role === 'TRUE' || user.has_leader_role === 'true'
  });
  return jsonResponse({ token: newToken, valid: true });
}

// ============================================================
// ユーザー管理 API JWT版ラッパー
// user-management.html から JWT 認証で呼ばれる（GAS_SECRET不要）
// ============================================================
function handleGetUsersJwt(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  try {
    var _r = getUsersRaw_();
    var users = _r.rows
      .filter(function(u) { return String(u.status).toLowerCase() !== 'deleted' && u.user_id; })
      .map(formatUserForClient_);
    return jsonResponse({ ok: true, users: users });
  } catch(err) {
    return errorResponse(err.message, 500);
  }
}

function handleCreateUserJwt(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  if (!body.name)          return jsonResponse({ ok: false, message: '氏名は必須です' });
  if (!body.email)         return jsonResponse({ ok: false, message: 'メールアドレスは必須です' });
  if (!validateEmail_(body.email)) return jsonResponse({ ok: false, message: 'メールアドレスの形式が正しくありません' });
  if (!body.role)          return jsonResponse({ ok: false, message: 'ロールは必須です' });
  if (['mentee','mentor','admin'].indexOf(String(body.role)) < 0) return jsonResponse({ ok: false, message: '無効なロールです' });
  if (!body.birthday)      return jsonResponse({ ok: false, message: '生年月日は必須です' });
  if (!body.password_hash) return jsonResponse({ ok: false, message: 'パスワードハッシュは必須です' });

  try {
    var _r = getUsersRaw_();
    var dup = _r.rows.find(function(u) {
      return String(u.email).toLowerCase() === String(body.email).toLowerCase()
        && String(u.status).toLowerCase() !== 'deleted' && u.user_id;
    });
    if (dup) return jsonResponse({ ok: false, message: 'このメールアドレスはすでに登録されています' });

    var user_id     = Utilities.getUuid();
    var now         = new Date().toISOString();
    var storedHash  = normalizePasswordHash(body.password_hash);

    appendRow('users', {
      user_id:         user_id,
      email:           String(body.email).trim().toLowerCase(),
      name:            String(body.name).trim(),
      role:            String(body.role),
      has_leader_role: body.has_leader_role ? 'TRUE' : 'FALSE',
      password_hash:   storedHash,
      mentor_id:       String(body.mentor_id  || ''),
      leader_id:       String(body.leader_id  || ''),
      phone_number:    String(body.phone       || '').trim(),
      workplace:       String(body.workplace   || '').trim(),
      work_status:     String(body.employment_type || ''),
      hourly_wage:     body.hourly_wage !== '' && body.hourly_wage != null ? Number(body.hourly_wage) || '' : '',
      status:          'active',
      created_at:      now,
      updated_at:      now,
      birthday:        String(body.birthday  || ''),
      hire_date:       String(body.hire_date || ''),
      chat_url:        String(body.chat_url         || '').trim(),
      chat_webhook_url:String(body.chat_webhook_url || '').trim(),
      tel_meet_url:    String(body.tel_meet_url     || '').trim(),
    });
    invalidateCache_('users');
    return jsonResponse({ ok: true, user_id: user_id, message: 'ユーザーを登録しました' });
  } catch(err) {
    return errorResponse(err.message, 500);
  }
}

function handleUpdateUserJwt(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  if (!body.user_id) return jsonResponse({ ok: false, message: 'user_id は必須です' });
  if (body.role !== undefined && ['mentee','mentor','admin'].indexOf(String(body.role)) < 0) {
    return jsonResponse({ ok: false, message: '無効なロールです' });
  }
  try {
    var _r      = getUsersRaw_();
    var target  = _r.rows.find(function(u) { return String(u.user_id) === String(body.user_id); });
    if (!target) return jsonResponse({ ok: false, message: 'ユーザーが見つかりません' });
    var now     = new Date().toISOString();
    var updates = {
      name:            body.name            !== undefined ? String(body.name).trim()                    : String(target.name            || ''),
      role:            body.role            !== undefined ? String(body.role)                           : String(target.role            || ''),
      has_leader_role: body.has_leader_role !== undefined ? (body.has_leader_role ? 'TRUE' : 'FALSE')  : String(target.has_leader_role || 'FALSE'),
      mentor_id:       body.mentor_id       !== undefined ? String(body.mentor_id  || '')              : String(target.mentor_id       || ''),
      leader_id:       body.leader_id       !== undefined ? String(body.leader_id  || '')              : String(target.leader_id       || ''),
      phone_number:    body.phone           !== undefined ? String(body.phone      || '').trim()        : String(target.phone_number    || ''),
      workplace:       body.workplace       !== undefined ? String(body.workplace  || '').trim()        : String(target.workplace       || ''),
      work_status:     body.employment_type !== undefined ? String(body.employment_type || '')          : String(target.work_status     || ''),
      hourly_wage:     body.hourly_wage     !== undefined ? (body.hourly_wage !== '' && body.hourly_wage != null ? Number(body.hourly_wage)||'' : '') : target.hourly_wage,
      birthday:        body.birthday        !== undefined ? String(body.birthday  || '')               : String(target.birthday        || ''),
      hire_date:       body.hire_date       !== undefined ? String(body.hire_date || '')               : String(target.hire_date       || ''),
      chat_url:        body.chat_url        !== undefined ? String(body.chat_url         || '').trim() : String(target.chat_url        || ''),
      chat_webhook_url:body.chat_webhook_url!== undefined ? String(body.chat_webhook_url|| '').trim() : String(target.chat_webhook_url|| ''),
      tel_meet_url:    body.tel_meet_url    !== undefined ? String(body.tel_meet_url     || '').trim() : String(target.tel_meet_url    || ''),
      calendar_email:  body.calendar_email  !== undefined ? String(body.calendar_email  || '').trim().toLowerCase() : String(target.calendar_email || ''),  // ★ カレンダー参照用メール
      personal_folder_id: body.personal_folder_id !== undefined ? String(body.personal_folder_id || '').trim() : String(target.personal_folder_id || ''),  // ★ 個人フォルダID
      updated_at:      now,
    };
    var currentRow = _r.headers.map(function(h) { return target[h] !== undefined ? target[h] : ''; });
    _r.headers.forEach(function(h, i) { if (updates.hasOwnProperty(h)) currentRow[i] = updates[h]; });
    _r.sheet.getRange(target.__rowIndex, 1, 1, _r.headers.length).setValues([currentRow]);
    invalidateCache_('users');
    return jsonResponse({ ok: true, message: 'ユーザー情報を更新しました' });
  } catch(err) {
    return errorResponse(err.message, 500);
  }
}

function handleDeleteUserJwt(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  if (!body.user_id) return jsonResponse({ ok: false, message: 'user_id は必須です' });
  try {
    var _r     = getUsersRaw_();
    var target = _r.rows.find(function(u) { return String(u.user_id) === String(body.user_id); });
    if (!target) return jsonResponse({ ok: false, message: 'ユーザーが見つかりません' });
    var now        = new Date().toISOString();
    var statusCol  = _r.headers.indexOf('status')     + 1;
    var updatedCol = _r.headers.indexOf('updated_at') + 1;
    if (statusCol  > 0) _r.sheet.getRange(target.__rowIndex, statusCol).setValue('deleted');
    if (updatedCol > 0) _r.sheet.getRange(target.__rowIndex, updatedCol).setValue(now);
    invalidateCache_('users');
    return jsonResponse({ ok: true, message: 'ユーザーを削除しました' });
  } catch(err) {
    return errorResponse(err.message, 500);
  }
}

// ============================================================
// ユーザー管理 API（user-management.html 連携）
// secret 認証ルート（doPost の secret 分岐）から呼ばれる
// ============================================================

// ── users シートのデータを取得（__rowIndex 付き）──
function getUsersRaw_() {
  var sheet = getSheet('users');
  if (!sheet) throw new Error('users シートが見つかりません');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).map(function(row, i) {
    var obj = {};
    headers.forEach(function(h, j) { obj[h] = row[j]; });
    obj.__rowIndex = i + 2;
    obj.__sheet = sheet;
    return obj;
  });
  return { sheet: sheet, headers: headers, rows: rows };
}

// ── ユーザーオブジェクトをフロント向けに整形 ──
function formatUserForClient_(u) {
  return {
    user_id:          String(u.user_id          || ''),
    email:            String(u.email            || ''),
    name:             String(u.name             || ''),
    role:             String(u.role             || ''),
    has_leader_role:  String(u.has_leader_role).toUpperCase() === 'TRUE',
    mentor_id:        String(u.mentor_id        || ''),
    leader_id:        String(u.leader_id        || ''),
    phone:            String(u.phone_number     || ''),
    workplace:        String(u.workplace        || ''),
    employment_type:  String(u.work_status      || ''),
    hourly_wage:      String(u.hourly_wage      || ''),
    status:           String(u.status           || 'active'),
    birthday:         String(u.birthday         || ''),
    hire_date:        String(u.hire_date        || ''), // ★ 入社日
    goal_work_6m:     String(u.goal_work_6m     || ''),
    goal_skill_6m:    String(u.goal_skill_6m    || ''),
    goal_start_month: normalizeYearMonth_(u.goal_start_month || ''), // ★ 目標期間 開始月
    goal_end_month:   normalizeYearMonth_(u.goal_end_month   || ''), // ★ 目標期間 終了月
    chat_url:         String(u.chat_url         || ''),
    chat_space_id:    String(u.chat_space_id    || ''),
    chat_webhook_url: String(u.chat_webhook_url || ''),
    created_at:       String(u.created_at       || ''),
    updated_at:       String(u.updated_at       || ''),
  };
}

// ── JWT でリクエスト元が admin か確認 ──
function requireAdminFromToken_(body, e) {
  var token = body.token || '';
  if (!token) token = getTokenFromRequest(e || {});
  if (!token) return { error: 'TOKEN_MISSING' };
  // ★ セキュリティ修正: 検証失敗時はフォールバックせず即座にエラー
  var payload = verifyJWT(token);
  if (!payload) return { error: 'INVALID_TOKEN' };
  if (payload.role !== 'admin') return { error: 'FORBIDDEN' };
  return { payload: payload };
}

// ────────────────────────────
// getUsers
// ────────────────────────────
function handleGetUsers(body, e) {
  var auth = requireAdminFromToken_(body, e);
  if (auth.error) return { ok: false, message: auth.error };

  try {
    var _r = getUsersRaw_();
    var users = _r.rows
      .filter(function(u) { return String(u.status).toLowerCase() !== 'deleted' && u.user_id; })
      .map(formatUserForClient_);
    return { ok: true, users: users };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

// ────────────────────────────
// createUser
// ────────────────────────────
function handleCreateUser(body, e) {
  var auth = requireAdminFromToken_(body, e);
  if (auth.error) return { ok: false, message: auth.error };

  if (!body.name)          return { ok: false, message: '氏名は必須です' };
  if (!body.email)         return { ok: false, message: 'メールアドレスは必須です' };
  if (!validateEmail_(body.email)) return { ok: false, message: 'メールアドレスの形式が正しくありません' };
  if (!body.role)          return { ok: false, message: 'ロールは必須です' };
  if (!body.birthday)      return { ok: false, message: '生年月日は必須です' };
  if (!body.password_hash) return { ok: false, message: 'パスワードハッシュは必須です' };
  // ★ セキュリティ: ロール値をホワイトリストで検証
  if (['mentee','mentor','admin'].indexOf(String(body.role)) < 0) return { ok: false, message: '無効なロールです' };

  try {
    var _r = getUsersRaw_();

    // メール重複チェック（deleted 以外）
    var dup = _r.rows.find(function(u) {
      return String(u.email).toLowerCase() === String(body.email).toLowerCase() &&
             String(u.status).toLowerCase() !== 'deleted' &&
             u.user_id;
    });
    if (dup) return { ok: false, message: 'このメールアドレスはすでに登録されています' };

    var user_id = Utilities.getUuid();
    var now = new Date().toISOString();
    // フロントから来る password_hash は "sha256:" なしの 64 桁 hex → 統一形式に変換して保存
    var storedHash = normalizePasswordHash(body.password_hash);

    // appendRow はヘッダー順にマッピングするので、新列（birthday/chat_url）があれば自動対応
    appendRow('users', {
      user_id:         user_id,
      email:           String(body.email).trim().toLowerCase(),
      name:            String(body.name).trim(),
      role:            String(body.role),
      has_leader_role: body.has_leader_role ? 'TRUE' : 'FALSE',
      password_hash:   storedHash,
      mentor_id:       String(body.mentor_id  || ''),
      leader_id:       String(body.leader_id  || ''),
      phone_number:    String(body.phone       || '').trim(),
      workplace:       String(body.workplace   || '').trim(),
      work_status:     String(body.employment_type || ''),
      hourly_wage:     body.hourly_wage !== '' && body.hourly_wage != null ? Number(body.hourly_wage) || '' : '',
      status:          'active',
      created_at:      now,
      updated_at:      now,
      birthday:        String(body.birthday   || ''),
      hire_date:       String(body.hire_date  || ''), // ★ 入社日
      chat_url:        String(body.chat_url         || '').trim(),
      chat_webhook_url:String(body.chat_webhook_url || '').trim(),
      tel_meet_url:    String(body.tel_meet_url     || '').trim(),
    });

    // リーダー割り当て処理（menteeの場合）
    if (String(body.role) === 'mentee' && body.leader_id) {
      var leaderCheck = setLeaderForMentee_(user_id, String(body.leader_id), '');
      if (!leaderCheck.ok) {
        // ユーザーは登録済みなのでエラーメッセージを返すがロールバックはしない
        return { ok: true, user_id: user_id, message: 'ユーザーを登録しました（リーダー設定エラー: ' + leaderCheck.error + '）' };
      }
      // leader_assignments にも登録
      var la_id = 'la-' + Date.now() + '-' + Math.random().toString(36).substring(2,7);
      var now2  = new Date().toISOString();
      appendRow('leader_assignments', {
        assignment_id: la_id,
        leader_id:     String(body.leader_id),
        mentee_id:     user_id,
        created_at:    now2,
        updated_at:    now2,
      });
    }

    return { ok: true, user_id: user_id, message: 'ユーザーを登録しました' };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

// ────────────────────────────
// updateUser
// ────────────────────────────
function handleUpdateUser(body, e) {
  var auth = requireAdminFromToken_(body, e);
  if (auth.error) return { ok: false, message: auth.error };
  if (!body.user_id) return { ok: false, message: 'user_id は必須です' };
  // ★ セキュリティ: ロール値をホワイトリストで検証
  if (body.role !== undefined && ['mentee','mentor','admin'].indexOf(String(body.role)) < 0) {
    return { ok: false, message: '無効なロールです' };
  }

  try {
    var _r = getUsersRaw_();
    var target = _r.rows.find(function(u) { return String(u.user_id) === String(body.user_id); });
    if (!target) return { ok: false, message: 'ユーザーが見つかりません' };

    var now = new Date().toISOString();

    // 更新フィールドを定義（undefined の場合は既存値を維持）
    var updates = {
      name:           body.name            !== undefined ? String(body.name).trim()            : String(target.name            || ''),
      role:           body.role            !== undefined ? String(body.role)                   : String(target.role            || ''),
      has_leader_role:body.has_leader_role !== undefined ? (body.has_leader_role ? 'TRUE' : 'FALSE') : String(target.has_leader_role || 'FALSE'),
      mentor_id:      body.mentor_id       !== undefined ? String(body.mentor_id  || '')       : String(target.mentor_id       || ''),
      leader_id:      body.leader_id       !== undefined ? String(body.leader_id  || '')       : String(target.leader_id       || ''),
      phone_number:   body.phone           !== undefined ? String(body.phone      || '').trim(): String(target.phone_number    || ''),
      workplace:      body.workplace       !== undefined ? String(body.workplace  || '').trim(): String(target.workplace       || ''),
      work_status:    body.employment_type !== undefined ? String(body.employment_type || '')  : String(target.work_status     || ''),
      hourly_wage:    body.hourly_wage     !== undefined ? (body.hourly_wage !== '' && body.hourly_wage != null ? Number(body.hourly_wage) || '' : '') : target.hourly_wage,
      birthday:       body.birthday        !== undefined ? String(body.birthday   || '')       : String(target.birthday        || ''),
      hire_date:      body.hire_date       !== undefined ? String(body.hire_date  || '')       : String(target.hire_date       || ''), // ★ 入社日
      chat_url:       body.chat_url        !== undefined ? String(body.chat_url         || '').trim(): String(target.chat_url        || ''),
      chat_webhook_url:body.chat_webhook_url !== undefined ? String(body.chat_webhook_url || '').trim(): String(target.chat_webhook_url || ''),
      tel_meet_url:   body.tel_meet_url    !== undefined ? String(body.tel_meet_url     || '').trim(): String(target.tel_meet_url    || ''),
      calendar_email: body.calendar_email  !== undefined ? String(body.calendar_email  || '').trim().toLowerCase() : String(target.calendar_email || ''),  // ★ カレンダー参照用メール
      personal_folder_id: body.personal_folder_id !== undefined ? String(body.personal_folder_id || '').trim() : String(target.personal_folder_id || ''),  // ★ 個人フォルダID
      updated_at:     now,
    };

    // シートの対象行を1回のAPI呼び出しでまとめて書き込み（バッチ化）
    var rowIndex = target.__rowIndex;
    var sheet    = _r.sheet;
    // 全列の現在値をコピーして更新フィールドだけ上書き
    var currentRow = _r.headers.map(function(h) { return target[h] !== undefined ? target[h] : ''; });
    _r.headers.forEach(function(h, colIdx) {
      if (updates.hasOwnProperty(h)) {
        currentRow[colIdx] = updates[h];
      }
    });
    sheet.getRange(rowIndex, 1, 1, _r.headers.length).setValues([currentRow]);

    // リーダー割り当て変更処理（menteeの場合）
    var newLeaderId     = updates.leader_id;
    var currentLeaderId = String(target.leader_id || '');

    if (String(updates.role || target.role) === 'mentee' && newLeaderId !== currentLeaderId) {
      var leaderCheck2 = setLeaderForMentee_(body.user_id, newLeaderId, currentLeaderId);
      if (!leaderCheck2.ok) {
        return { ok: false, message: leaderCheck2.error };
      }

      // leader_assignments の更新（古い割り当て削除→新規追加）
      var laSheet = getSheet('leader_assignments');
      if (laSheet) {
        var laData    = laSheet.getDataRange().getValues();
        var laHeaders = laData[0];
        var mCol = laHeaders.indexOf('mentee_id');
        var lCol = laHeaders.indexOf('leader_id');
        // 既存のこのメンティーの割り当てを削除（後ろから）
        for (var ri = laData.length - 1; ri >= 1; ri--) {
          if (String(laData[ri][mCol]) === String(body.user_id) &&
              String(laData[ri][lCol]) === currentLeaderId) {
            laSheet.deleteRow(ri + 1);
          }
        }
        // 新しいリーダーがある場合は追加
        if (newLeaderId) {
          var la_id2 = 'la-' + Date.now() + '-' + Math.random().toString(36).substring(2,7);
          var now3   = new Date().toISOString();
          appendRow('leader_assignments', {
            assignment_id: la_id2,
            leader_id:     newLeaderId,
            mentee_id:     body.user_id,
            created_at:    now3,
            updated_at:    now3,
          });
        }
      }
    }

    return { ok: true, message: 'ユーザー情報を更新しました' };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}

// ────────────────────────────
// deleteUser（論理削除: status = 'deleted'）
// ────────────────────────────
function handleDeleteUser(body, e) {
  var auth = requireAdminFromToken_(body, e);
  if (auth.error) return { ok: false, message: auth.error };
  if (!body.user_id) return { ok: false, message: 'user_id は必須です' };

  try {
    var _r = getUsersRaw_();
    var target = _r.rows.find(function(u) { return String(u.user_id) === String(body.user_id); });
    if (!target) return { ok: false, message: 'ユーザーが見つかりません' };

    var now        = new Date().toISOString();
    var rowIndex   = target.__rowIndex;
    var sheet      = _r.sheet;
    var statusCol  = _r.headers.indexOf('status')     + 1;
    var updatedCol = _r.headers.indexOf('updated_at') + 1;

    if (statusCol  > 0) sheet.getRange(rowIndex, statusCol).setValue('deleted');
    if (updatedCol > 0) sheet.getRange(rowIndex, updatedCol).setValue(now);

    return { ok: true, message: 'ユーザーを削除しました' };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}


// ============================================================
// コンテンツ削除
// ============================================================
function handleDeleteContent(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var content_id = body.content_id || '';
  if (!content_id) return errorResponse('MISSING_CONTENT_ID', 400);

  var sheet = getSheet('contents');
  if (!sheet) return errorResponse('SHEET_NOT_FOUND', 500);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf('content_id');
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIdx]) === String(content_id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return errorResponse('CONTENT_NOT_FOUND', 404);
}

// ============================================================
// リンク削除
// ============================================================
function handleDeleteLink(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var link_id = body.link_id || '';
  if (!link_id) return errorResponse('MISSING_LINK_ID', 400);

  var sheet = getSheet('quick_links');
  if (!sheet) return errorResponse('SHEET_NOT_FOUND', 500);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf('link_id');
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIdx]) === String(link_id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return errorResponse('LINK_NOT_FOUND', 404);
}

// ============================================================
// お知らせ一覧（管理者向け全件）
// ============================================================
function handleAdminNoticeList(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var all = cachedSheetToObjects_('admin_memos').filter(function(m) { return m.memo_id; });
  // is_notice=TRUE のものだけ（内部メモは除外）
  var notices = all
    .filter(function(m) { return String(m.is_notice).toUpperCase() === 'TRUE'; }) // ブール値対応
    .sort(function(a, b) { return (b.created_at || '').localeCompare(a.created_at || ''); })
    .slice(0, 100)
    .map(function(m) {
      return {
        memo_id:       m.memo_id,
        content:       m.content,
        created_at:    m.created_at,
        target_id:     m.target_id     || 'all',
        is_active:     m.is_active     || 'TRUE',
        display_from:  m.display_from  || '',
        display_until: m.display_until || ''
      };
    });
  return jsonResponse({ ok: true, notices: notices });
}

// ============================================================
// お知らせ削除
// ============================================================
function handleDeleteNotice(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var memo_id = body.memo_id || '';
  if (!memo_id) return errorResponse('MISSING_MEMO_ID', 400);

  var sheet = getSheet('admin_memos');
  if (!sheet) return errorResponse('SHEET_NOT_FOUND', 500);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf('memo_id');
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIdx]) === String(memo_id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return errorResponse('MEMO_NOT_FOUND', 404);
}

// ============================================================
// 管理者メモ取得（特定ユーザー向け）
// admin_memos シートの target_id = user_id のもの
// ============================================================
function handleGetAdminMemos(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var body = parseBody_(e);
  var user_id = body.user_id || '';

  if (!user_id) return errorResponse('MISSING_USER_ID', 400);

  var sheet = getSheet('admin_memos');
  if (!sheet) return jsonResponse({ ok: true, memos: [] });

  var all = sheetToObjects(sheet).filter(function(m) { return m.memo_id; });
  var memos = all
    .filter(function(m) { return m.target_id === user_id && String(m.is_notice || '').toUpperCase() !== 'TRUE'; })
    .sort(function(a, b) { return (b.created_at || '').localeCompare(a.created_at || ''); })
    .slice(0, 50)
    .map(function(m) {
      return {
        memo_id:    m.memo_id,
        content:    m.content,
        created_at: m.created_at,
        // ★ memo_type: 'memo'=管理者メモ / 'message'=メッセージ送信
        memo_type:  m.memo_type || 'memo',
      };
    });
  return jsonResponse({ ok: true, memos: memos });
}

// ============================================================
// 管理者メモ保存
// ============================================================
function handleSaveAdminMemo(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var target_id = body.target_id || '';
  var content   = body.content   || '';
  var memo_type = body.memo_type || 'memo'; // 'memo' or 'message'
  if (!target_id || !content) return errorResponse('MISSING_FIELDS', 400);

  var memo_id = 'memo-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var now = new Date().toISOString();

  var memoSheet = getSheet('admin_memos');
  if (!memoSheet) {
    memoSheet = getSpreadsheet_().insertSheet('admin_memos');
    memoSheet.appendRow(['memo_id','admin_id','target_id','content','memo_type','created_at','updated_at']);
  }

  try {
    appendRow('admin_memos', {
      memo_id:    memo_id,
      admin_id:   auth.payload.user_id,
      target_id:  target_id,
      content:    content,
      memo_type:  memo_type,
      created_at: now,
      updated_at: now
    });
  } catch(err) {
    return errorResponse('SHEET_ERROR: ' + err.message, 500);
  }
  return jsonResponse({ ok: true, memo_id: memo_id });
}


// ============================================================
// メンター: 担当メンティー一覧
// ============================================================
// ============================================================
// Mentor: 全メンバー一覧（自分の担当かどうかのフラグ付き）
// GET api/mentor/members
// ============================================================
function handleMentorMembers(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var mentor_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u.name; });

  var members = users
    .filter(function(u){ return u.role === 'mentee'; })
    .map(function(u) {
      var tenure = calcTenure_(u.hire_date || '');
      return {
        user_id:       u.user_id,
        name:          u.name          || '',
        role:          u.role          || '',
        status:        u.status        || 'green',
        workplace:     u.workplace     || '',
        employment_type: u.employment_type || '',
        mentor_id:     u.mentor_id     || '',
        mentor_name:   userMap[u.mentor_id] || '—',
        hire_date:     u.hire_date     || '',
        tenure_months: tenure.months,
        tenure_label:  tenure.label,
        is_my_mentee:  u.mentor_id === mentor_id, // ★ 自分の担当フラグ
      };
    });

  return jsonResponse({ ok: true, members: members });
}

function handleMentorMyMentees(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });

  // ★ 公開済みメンターレポートから最新AI要約を取得（メンティーID→最新レポートのマップ）
  var lastRepMap = {};
  cachedSheetToObjects_('mentor_reports')
    .filter(function(r){ return r.report_id && r.mentor_id === user_id && r.is_published === 'TRUE'; })
    .sort(function(a, b){ return (b.created_at || '').localeCompare(a.created_at || ''); })
    .forEach(function(r){
      if (!lastRepMap[r.mentee_id]) {
        lastRepMap[r.mentee_id] = { report_id: r.report_id, ai_summary: r.ai_summary || '' };
      }
    });

  var mentees = users.filter(function(u){
    return u.mentor_id === user_id && u.status !== 'deleted';
  }).map(function(u){
    var lastRep = lastRepMap[u.user_id] || null;
    return {
      user_id:          u.user_id,
      name:             u.name,
      email:            u.email,
      status:           u.status || 'green',
      workplace:        u.workplace || '',
      work_status:      u.work_status || '',
      hourly_wage:      u.hourly_wage || '',
      phone:            u.phone_number || '',
      hire_date:        u.hire_date || '',
      goal_work_6m:     u.goal_work_6m     || '',
      goal_skill_6m:    u.goal_skill_6m    || '',
      goal_start_month: normalizeYearMonth_(u.goal_start_month || ''),
      goal_end_month:   normalizeYearMonth_(u.goal_end_month   || ''),
      chat_url:         u.chat_url || '',
      chat_webhook_url: u.chat_webhook_url || '',
      tel_meet_url:     u.tel_meet_url || '',
      leader_id:        u.leader_id || '',
      last_report_id:   lastRep ? lastRep.report_id  : '',
      last_ai_summary:  lastRep ? lastRep.ai_summary : '',
    };
  });
  return jsonResponse({ ok:true, mentees: mentees });
}

// ============================================================
// メンター: 自分のスケジュール（bookings）一覧
// ============================================================
function handleMentorMyBookings(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;
  var role    = auth.payload.role;

  var bookings = cachedSheetToObjects_('bookings').filter(function(b){ return b.booking_id; });
  var mine = bookings.filter(function(b){
    return (b.mentor_id === user_id || role === 'admin') && b.status !== 'cancelled';
  }).sort(function(a,b){ return (a.scheduled_at||'').localeCompare(b.scheduled_at||''); });

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var result = mine.map(function(b){
    var mentee = users.find(function(u){ return u.user_id === b.mentee_id; });
    return {
      booking_id:   b.booking_id,
      mentee_id:    b.mentee_id,
      mentee_name:  mentee ? mentee.name : '',
      mentor_id:    b.mentor_id,
      scheduled_at: b.scheduled_at,
      meet_link:    b.meet_link || '',
      status:       b.status || 'confirmed',
    };
  });
  return jsonResponse({ ok:true, bookings: result });
}

// ============================================================
// メンター: メンティーのステータス更新
// ============================================================
function handleMentorUpdateStatus(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var target_id = body.target_id || '';
  var status    = body.status    || '';
  if (!target_id || !status) return errorResponse('MISSING_FIELDS', 400);

  // ★ updateRowWhere でバッチ書き込み（getRange個別呼び出しを排除）
  updateRowWhere('users', 'user_id', target_id, {
    status:     status,
    updated_at: new Date().toISOString()
  });
  invalidateCache_('users');
  return jsonResponse({ ok: true });
}

// ────────────────────────────────────────────────────────
// setupUsersSheetColumns
// ★ GASエディタから一度だけ手動実行してください
// birthday 列・chat_url 列をスプシに追加します
// ────────────────────────────────────────────────────────
function setupPreReportsSheet() {
  // pre_reports シートを手動セットアップする関数（GASエディタから一度だけ手動実行）
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName('pre_reports');
  if (sheet) { Logger.log('pre_reports シートは既に存在します'); return; }
  sheet = ss.insertSheet('pre_reports');
  sheet.getRange(1, 1, 1, 15).setValues([[
    'report_id', 'user_id', 'target_month', 'session_date',
    'current_project', 'goal_work_6m', 'goal_skill_6m',
    'project_result', 'study_hours', 'study_content',
    'good_points', 'improvement_points',
    'next_month_project_goal', 'next_month_study_goal',
    'submitted_at'
  ]]);
  Logger.log('pre_reports シートを作成しました');
}

function diagPreReportsSheet() {
  // GASエディタから手動実行 → Loggerで pre_reports シートの状態を確認
  var sheet = getSheet('pre_reports');
  if (!sheet) {
    Logger.log('❌ pre_reports シートが存在しません');
    return;
  }
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  Logger.log('✅ pre_reports シート存在');
  Logger.log('ヘッダー列数: ' + headers.length);
  Logger.log('ヘッダー: ' + JSON.stringify(headers));
  Logger.log('データ行数: ' + (data.length - 1));

  // コードが期待するヘッダーとの差分
  var expected = [
    'report_id', 'user_id', 'target_month', 'session_date',
    'current_project', 'goal_work_6m', 'goal_skill_6m',
    'project_result', 'study_hours', 'study_content',
    'good_points', 'improvement_points',
    'next_month_project_goal', 'next_month_study_goal',
    'submitted_at'
  ];
  var missing  = expected.filter(function(h){ return headers.indexOf(h) < 0; });
  var extra    = headers.filter(function(h){ return expected.indexOf(h) < 0 && h !== ''; });
  if (missing.length > 0)  Logger.log('❌ 不足列: ' + missing.join(', '));
  if (extra.length > 0)    Logger.log('⚠️ 余分な列: ' + extra.join(', '));
  if (missing.length === 0 && extra.length === 0) Logger.log('✅ ヘッダーは完全一致');

  // 最新5件のreport_idとsubmitted_atを表示
  Logger.log('--- 最新データ（最大5行）---');
  data.slice(1).slice(-5).forEach(function(row, i) {
    var obj = {};
    headers.forEach(function(h, j){ obj[h] = row[j]; });
    Logger.log('行' + (i+1) + ': report_id=' + obj.report_id
      + ' user_id=' + obj.user_id
      + ' submitted_at=' + obj.submitted_at);
  });
}

// pre_reports シートのヘッダーをコードの期待値に揃える（既存データは保持）
function fixPreReportsHeaders() {
  var sheet = getSheet('pre_reports');
  if (!sheet) { Logger.log('シートが存在しません。setupPreReportsSheet()を実行してください'); return; }

  var expected = [
    'report_id', 'user_id', 'target_month', 'session_date',
    'current_project', 'goal_work_6m', 'goal_skill_6m',
    'project_result', 'study_hours', 'study_content',
    'good_points', 'improvement_points',
    'next_month_project_goal', 'next_month_study_goal',
    'submitted_at'
  ];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 不足列を末尾に追加
  expected.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
      headers.push(col);
      Logger.log('列追加: ' + col + ' → 列' + newCol);
    }
  });
  Logger.log('fixPreReportsHeaders 完了');
}

function setupMentorSchedulesSheet() {
  // mentor_schedules シートを手動セットアップする関数（GASエディタから一度だけ手動実行）
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName('mentor_schedules');
  if (sheet) {
    Logger.log('mentor_schedules シートは既に存在します');
    return;
  }
  sheet = ss.insertSheet('mentor_schedules');
  sheet.getRange(1, 1, 1, 7).setValues([[
    'schedule_id', 'mentor_id', 'day_of_week',
    'start_time', 'end_time', 'is_active', 'created_at'
  ]]);
  Logger.log('mentor_schedules シートを作成しました');
}

function setupUsersSheetColumns() {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName('users');
  if (!sheet) { Logger.log('ERROR: users シートが見つかりません'); return; }

  function addColIfMissing(name) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes(name)) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(name);
      Logger.log('列を追加しました: ' + name + ' (列' + newCol + ')');
    } else {
      Logger.log('列は既存: ' + name);
    }
  }

  addColIfMissing('birthday');
  addColIfMissing('chat_url');
  addColIfMissing('hire_date');          // ★ 入社日
  addColIfMissing('goal_work_6m');       // ★ 半年業務目標
  addColIfMissing('goal_skill_6m');      // ★ 半年スキル目標
  addColIfMissing('goal_start_month');   // ★ 目標期間 開始月
  addColIfMissing('goal_end_month');     // ★ 目標期間 終了月
  Logger.log('setupUsersSheetColumns 完了');
}

// bookings シートに survey_notified 列を追加するセットアップ関数
function setupBookingsSheetColumns() {
  var sheet = getSheet('bookings');
  if (!sheet) { Logger.log('bookings シートが存在しません'); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  function addColIfMissing(colName) {
    if (headers.indexOf(colName) < 0) {
      sheet.getRange(1, headers.length + 1).setValue(colName);
      headers.push(colName);
      Logger.log('bookings に列追加: ' + colName);
    }
  }
  addColIfMissing('survey_notified'); // ★ アンケート通知済みフラグ
  Logger.log('setupBookingsSheetColumns 完了');
}

// ★ 既存のcompleted予約に一括でsurvey_notified=TRUEを設定する
// （過去分のアンケートメール再送を防ぐために一度だけ手動実行する）
function markAllExistingBookingsAsNotified() {
  invalidateCache_('bookings');
  var bookings = cachedSheetToObjects_('bookings');
  var count = 0;
  bookings.forEach(function(b) {
    if (b.booking_id && b.status === 'completed'
        && String(b.survey_notified || '').toUpperCase() !== 'TRUE') {
      updateRowWhere('bookings', 'booking_id', b.booking_id, {
        survey_notified: 'TRUE', updated_at: new Date().toISOString()
      });
      count++;
    }
  });
  invalidateCache_('bookings');
  Logger.log('markAllExistingBookingsAsNotified: ' + count + '件にフラグを設定しました');
}

// ============================================================
// F-07: 閲覧ログ記録
// ============================================================
function handlePageView(e) {
  var body = parseBody_(e);
  var token = getTokenFromRequest(e);
  // ★ ここでは認証は行わず、ログ記録のみが目的のため decodeJWTNoVerify を使用
  // アクセスログはベストエフォート（トークン無効でも記録する）
  var payload = decodeJWTNoVerify(token) || {};

  var log_id = (payload.user_id || 'unknown') + '_' + Date.now();
  appendRow('access_logs', {
    log_id: log_id,
    user_id: payload.user_id || 'unknown',
    role: payload.role || 'unknown',
    page: (body.page || 'unknown').trim(),
    action: (body.action || '').trim(),
    user_agent: '',
    logged_at: new Date().toISOString()
  });
  return jsonResponse({ ok: true });
}

// ============================================================
// F-Mentee: ホームデータ
// ============================================================
function handleMenteeHome(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users');
  var bookings = cachedSheetToObjects_('bookings');

  var me = users.find(function(u) { return u.user_id === user_id; });
  var mentor = me && me.mentor_id ? users.find(function(u) { return u.user_id === me.mentor_id; }) : null;

  var now = new Date().toISOString();
  var myBookings = bookings.filter(function(b) {
    return b.booking_id && b.mentee_id === user_id && b.status !== 'cancelled';
  });
  var upcoming = myBookings
    .filter(function(b) { return b.scheduled_at >= now; })
    .sort(function(a, b) { return a.scheduled_at.localeCompare(b.scheduled_at); });
  var nextBooking = upcoming[0] || null;

  var leader = me && me.leader_id ? users.find(function(u){ return u.user_id === me.leader_id; }) : null;

  // ★ 未回答アンケート（completed & 未回答の最新1件）
  var answeredIds = {};
  cachedSheetToObjects_('surveys')
    .filter(function(s){ return s.booking_id && s.user_id === user_id; })
    .forEach(function(s){ answeredIds[s.booking_id] = true; });
  // ★ 1on1開始時刻（scheduled_at）を過ぎたら表示（Item8: 完了後→開始時刻に変更）
  var _surveyNow = new Date();
  var pendingSurvey = myBookings
    .filter(function(b){
      if (answeredIds[b.booking_id]) return false;
      if (b.status === 'cancelled') return false;
      return b.status === 'completed' ||
        ((b.status === 'confirmed' || b.status === 'in_progress') && new Date(b.scheduled_at) <= _surveyNow);
    })
    .sort(function(a, b){ return (b.scheduled_at||'').localeCompare(a.scheduled_at||''); })[0] || null;

  return jsonResponse({
    user: {
      user_id:          me ? me.user_id          : '',
      name:             me ? me.name             : '',
      role:             me ? me.role             : '',
      email:            me ? me.email            : '',
      chat_url:         me ? (me.chat_url         || '') : '',
      chat_space_id:    me ? (me.chat_space_id    || '') : '',
      workplace:        me ? (me.workplace        || '') : '',
      work_status:      me ? (me.work_status      || '') : '',
      hourly_wage:      me ? (me.hourly_wage      || '') : '',
      phone_number:     me ? (me.phone_number     || '') : '',
      mentor_id:        me ? (me.mentor_id        || '') : '',
      leader_id:        me ? (me.leader_id        || '') : '',
      hire_date:        me ? (me.hire_date        || '') : '',
    },
    goal_work_6m:     me ? (me.goal_work_6m     || '') : '',
    goal_skill_6m:    me ? (me.goal_skill_6m    || '') : '',
    goal_start_month: normalizeYearMonth_(me ? (me.goal_start_month || '') : ''), // ★ 目標期間 開始月
    goal_end_month:   normalizeYearMonth_(me ? (me.goal_end_month   || '') : ''), // ★ 目標期間 終了月
    current_project: me ? (me.current_project  || '') : '',
    mentor: mentor ? { user_id: mentor.user_id, name: mentor.name, email: mentor.email, chat_url: mentor.chat_url||'' } : null,
    leader: leader ? {
      user_id:      leader.user_id,
      name:         leader.name,
      phone_number: leader.phone_number || '',
      chat_url:     leader.chat_url || '',
    } : null,
    next_booking: nextBooking ? {
      booking_id:        nextBooking.booking_id,
      scheduled_at:      nextBooking.scheduled_at,
      meet_link:         nextBooking.meet_link || '',
      calendar_event_id: nextBooking.calendar_event_id || '',
      calendar_url:      nextBooking.calendar_url      || ''
    } : null,
    upcoming_count: upcoming.length,
    all_upcoming_bookings: upcoming.map(function(b) {
      return {
        booking_id:        b.booking_id,
        scheduled_at:      b.scheduled_at,
        meet_link:         b.meet_link || '',
        duration_minutes:  b.duration_minutes || 60,
        calendar_event_id: b.calendar_event_id || '',
        calendar_url:      b.calendar_url      || ''
      };
    }),
    // ★ 未回答アンケート
    pending_survey: pendingSurvey ? {
      booking_id:   pendingSurvey.booking_id,
      scheduled_at: pendingSurvey.scheduled_at,
    } : null,
  });
}

// ============================================================
// F-Mentee: お知らせ
// ============================================================
function handleMenteeNotices(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;
  var now = new Date();

  var all = cachedSheetToObjects_('admin_memos').filter(function(m) { return m.memo_id; });
  var notices = all
    .filter(function(m) {
      // ① is_notice=TRUE のものだけ（管理者が明示的に作成したお知らせのみ）
      // 'TRUE'/'true'/true(boolean) 全て対応
      var isnRaw = m.is_notice;
      var isn = (isnRaw === true || isnRaw === 'TRUE' || isnRaw === 'true' || isnRaw === 'True')
                ? 'TRUE' : String(isnRaw || '').toUpperCase();
      if (isn !== 'TRUE') return false;
      // ② is_active=FALSE なら非表示
      // is_active が明示的に FALSE の場合のみ除外（空・undefined は表示する）
      var isa = String(m.is_active || '').toUpperCase();
      if (isa === 'FALSE') return false;
      // ③ 配信対象チェック（all / role:xxx / user_id）
      if (m.target_id !== 'all') {
        if (m.target_id.indexOf('role:') === 0) {
          // ロール指定: 'role:mentor' / 'role:mentee' / 'role:admin'
          var targetRole = m.target_id.split(':')[1];
          if (targetRole !== auth.payload.role) return false;
        } else {
          // 個人指定: user_id
          if (m.target_id !== user_id) return false;
        }
      }
      // ④ 表示期間チェック
      if (m.display_from) {
        var from = new Date(m.display_from);
        if (!isNaN(from) && now < from) return false;
      }
      if (m.display_until) {
        var until = new Date(m.display_until);
        if (!isNaN(until) && now > until) return false;
      }
      return true;
    })
    .sort(function(a, b) { return (b.created_at || '').localeCompare(a.created_at || ''); })
    .slice(0, 20)
    .map(function(m) {
      return {
        memo_id:       m.memo_id,
        content:       m.content,
        created_at:    m.created_at,
        target_id:     m.target_id,
        display_from:  m.display_from  || '',
        display_until: m.display_until || ''
      };
    });
  return jsonResponse({ notices: notices });
}

function handleAdminNotice(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  if (!body.content) return errorResponse('MISSING_CONTENT', 400);

  // admin_memosシートに必要なカラムがなければ自動追加（初回のみ）
  var memoSheet = getSheet('admin_memos');
  if (memoSheet) {
    var memoHeaders = memoSheet.getRange(1,1,1,memoSheet.getLastColumn()).getValues()[0];
    ['is_notice','is_active','display_from','display_until'].forEach(function(col){
      if (memoHeaders.indexOf(col) < 0) {
        memoSheet.getRange(1, memoSheet.getLastColumn()+1).setValue(col);
        Logger.log('handleAdminNotice: カラム自動追加 → ' + col);
      }
    });
  }

  var memo_id = 'notice-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var now     = new Date().toISOString();
  appendRow('admin_memos', {
    memo_id:       memo_id,
    admin_id:      auth.payload.user_id,
    target_id:     body.target_id     || 'all',
    content:       body.content,
    is_notice:     'TRUE',              // 公開お知らせフラグ
    is_active:     'TRUE',              // 表示ON
    display_from:  body.display_from  || '',  // 表示開始日時
    display_until: body.display_until || '',  // 表示終了日時
    created_at:    now,
    updated_at:    now
  });
  return jsonResponse({ ok: true, memo_id: memo_id });
}

// ============================================================
// Admin: お知らせの is_active 切り替え（表示ON/OFF）
// POST api/admin/notices/toggle { memo_id }
// ============================================================
function handleToggleNotice(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var memo_id = (body.memo_id || '').trim();
  if (!memo_id) return errorResponse('MISSING_MEMO_ID', 400);

  var all = sheetToObjects(getSheet('admin_memos')).filter(function(m){ return m.memo_id; });
  var target = all.find(function(m){ return m.memo_id === memo_id; });
  if (!target) return errorResponse('NOT_FOUND', 404);

  var newActive = String(target.is_active).toUpperCase() === 'TRUE' ? 'FALSE' : 'TRUE'; // ブール値対応
  updateRowWhere('admin_memos', 'memo_id', memo_id, {
    is_active:  newActive,
    updated_at: new Date().toISOString()
  });
  invalidateCache_('admin_memos');
  return jsonResponse({ ok: true, is_active: newActive });
}

// ============================================================
// F-Mentee: 空き時間取得
// ============================================================
function handleAvailableSlots(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users');
  var me = users.find(function(u) { return u.user_id === user_id; });
  if (!me) return errorResponse('USER_NOT_FOUND', 404);
  var mentor_id = me.mentor_id;
  if (!mentor_id) return jsonResponse({ error: 'NO_MENTOR', slots: {} });

  var schedules = cachedSheetToObjects_('mentor_schedules')
    .filter(function(s) { return s.schedule_id && s.mentor_id === mentor_id && s.is_active === 'TRUE'; });

  var availableConfig = {};
  if (schedules.length > 0) {
    schedules.forEach(function(s) {
      // ★ スプレッドシートの時刻型変換に対応して正規化
      availableConfig[String(s.day_of_week)] = {
        start: normalizeTimeStr_(s.start_time) || '10:00',
        end:   normalizeTimeStr_(s.end_time)   || '19:00'
      };
    });
  } else {
    ['1','2','3','4','5'].forEach(function(d) { availableConfig[d] = { start: '10:00', end: '19:00' }; });
  }

  // ── rangeStart/rangeEnd を先に定義（カレンダー取得より前） ──
  var now        = new Date();
  var rangeStart = new Date(now); rangeStart.setHours(0, 0, 0, 0);
  var rangeEnd   = new Date(rangeStart); rangeEnd.setDate(rangeStart.getDate() + 14);

  // ★ メンターの default_1on1_duration を取得
  // 未設定・30分以下の場合は60分（システムデフォルト）
  var mentor = users.find(function(u){ return u.user_id === mentor_id; });
  var durationMinutes = (mentor && parseInt(mentor.default_1on1_duration) > 30)
    ? parseInt(mentor.default_1on1_duration)
    : 60;

  // bookingsシートの既存1on1予約をブロック（JST基準でキー生成）
  var bookings = cachedSheetToObjects_('bookings')
    .filter(function(b) { return b.booking_id && b.mentor_id === mentor_id && b.status !== 'cancelled'; });
  var bookedSlots = {};
  bookings.forEach(function(b) {
    if (!b.scheduled_at) return;
    var bDuration = parseInt(b.duration_minutes) || durationMinutes;
    var dt  = new Date(b.scheduled_at);
    // 予約開始〜終了の全30分スロットをブロック
    for (var bi = 0; bi < bDuration; bi += 30) {
      var blockTime = new Date(dt.getTime() + bi * 60 * 1000);
      var jst = new Date(blockTime.getTime() + 9 * 60 * 60 * 1000);
      var key = jst.toISOString().slice(0,10) + 'T' +
        ('0'+jst.getUTCHours()).slice(-2) + ':' + ('0'+jst.getUTCMinutes()).slice(-2);
      bookedSlots[key] = true;
    }
  });

  // ★ Googleカレンダー連携：calendar_email があればそちらを優先使用
  // （例: @agent-network.com のカレンダーを @socialshift.work に共有している場合）
  try {
    var calendar = null;
    if (mentor) {
      var calEmail = (mentor.calendar_email || '').trim() || (mentor.email || '').trim();
      if (calEmail) {
        try {
          calendar = CalendarApp.getCalendarById(calEmail);
          if (!calendar) Logger.log('カレンダー取得失敗（null）: ' + calEmail);
        } catch(e) {
          Logger.log('カレンダー取得エラー（' + calEmail + '）: ' + e.message);
        }
      }
    }
    if (!calendar) calendar = CalendarApp.getDefaultCalendar();

    var events = calendar.getEvents(rangeStart, rangeEnd);
    events.forEach(function(ev) {
      if (ev.isAllDayEvent()) return;
      // ★ キャンセル済みイベント（タイトルに【キャンセル1on1】を含む）はスキップ
      if (ev.getTitle().indexOf('【キャンセル1on1】') !== -1) return;
      var evStart = ev.getStartTime();
      var evEnd   = ev.getEndTime();
      var cur = new Date(evStart);
      cur.setSeconds(0); cur.setMilliseconds(0);
      cur.setMinutes(cur.getMinutes() < 30 ? 0 : 30);
      while (cur < evEnd) {
        var jstCur = new Date(cur.getTime() + 9 * 60 * 60 * 1000);
        var dateS  = jstCur.toISOString().slice(0, 10);
        var timeS  = ('0'+jstCur.getUTCHours()).slice(-2) + ':' + ('0'+jstCur.getUTCMinutes()).slice(-2);
        bookedSlots[dateS + 'T' + timeS] = true;
        cur.setMinutes(cur.getMinutes() + 30);
      }
    });
    Logger.log('カレンダーイベント取得: ' + events.length + '件 （' + (mentor ? (mentor.calendar_email || mentor.email) : 'default') + '）');
  } catch(calErr) {
    Logger.log('カレンダー取得エラー（無視）: ' + calErr.message);
  }

  // スロット生成（JST基準・30分刻み・durationMinutes分の空きを確認）
  var result = {};
  for (var d = new Date(rangeStart); d < rangeEnd; d.setDate(d.getDate() + 1)) {
    var jstD   = new Date(d.getTime() + 9 * 60 * 60 * 1000);
    var dow    = String(jstD.getUTCDay());
    var config = availableConfig[dow];
    if (!config) continue;
    var dateStr    = jstD.toISOString().slice(0, 10);
    var slots      = [];
    var startParts = config.start.split(':').map(Number);
    var endParts   = config.end.split(':').map(Number);

    // ★ dateStr（JST基準の日付）+ 時刻でUTCのDateを作る
    // 例: dateStr='2026-04-20', start='10:00' → UTC '2026-04-20T01:00:00Z'（JST10:00）
    var current = new Date(dateStr + 'T'
      + ('0'+startParts[0]).slice(-2) + ':' + ('0'+startParts[1]).slice(-2) + ':00+09:00');
    var endTime = new Date(dateStr + 'T'
      + ('0'+endParts[0]).slice(-2) + ':' + ('0'+endParts[1]).slice(-2) + ':00+09:00');

    // ★ 終了時刻から durationMinutes 分引いた時刻までしかスロットを生成しない
    var lastSlotTime = new Date(endTime.getTime() - durationMinutes * 60 * 1000);

    while (current <= lastSlotTime) {
      if (current > now) {
        var jstCurrent = new Date(current.getTime() + 9 * 60 * 60 * 1000);
        var timeStr    = ('0'+jstCurrent.getUTCHours()).slice(-2) + ':' + ('0'+jstCurrent.getUTCMinutes()).slice(-2);
        var slotKey    = dateStr + 'T' + timeStr;

        // ★ 開始から durationMinutes 分の間に埋まりがないか30分刻みで確認
        var isAvailable = true;
        for (var sm = 0; sm < durationMinutes; sm += 30) {
          var checkTime = new Date(current.getTime() + sm * 60 * 1000);
          var jstCheck  = new Date(checkTime.getTime() + 9 * 60 * 60 * 1000);
          var checkKey  = jstCheck.toISOString().slice(0,10) + 'T'
            + ('0'+jstCheck.getUTCHours()).slice(-2) + ':' + ('0'+jstCheck.getUTCMinutes()).slice(-2);
          if (bookedSlots[checkKey]) { isAvailable = false; break; }
        }

        slots.push({ time: timeStr, datetime: new Date(current).toISOString(), available: isAvailable });
      }
      current.setMinutes(current.getMinutes() + 30);
    }
    if (slots.length > 0) result[dateStr] = slots;
  }
  return jsonResponse({ mentor_id: mentor_id, slots: result, slot_minutes: durationMinutes });
}


// ============================================================
// Mentor: 空き時間スロット取得（予約作成・変更時に使用）
// GET api/mentor/available-slots?mentee_id=xxx
// ============================================================
function handleMentorAvailableSlots(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var mentor_id = auth.payload.user_id;

  var body = parseBody_(e);
  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentor = users.find(function(u){ return u.user_id === mentor_id; });

  // ★ durationMinutes: メンターの設定値（30以下は60に統一）
  var durationMinutes = (mentor && parseInt(mentor.default_1on1_duration) > 30)
    ? parseInt(mentor.default_1on1_duration) : 60;

  var schedules = cachedSheetToObjects_('mentor_schedules')
    .filter(function(s){ return s.schedule_id && s.mentor_id === mentor_id && s.is_active === 'TRUE'; });

  var availableConfig = {};
  if (schedules.length > 0) {
    schedules.forEach(function(s){
      availableConfig[String(s.day_of_week)] = { start: s.start_time || '10:00', end: s.end_time || '19:00' };
    });
  } else {
    ['1','2','3','4','5'].forEach(function(d){ availableConfig[d] = { start: '10:00', end: '19:00' }; });
  }

  // ★ bookingsシートの既存予約をブロック（JST基準でキー生成）
  var bookings = cachedSheetToObjects_('bookings')
    .filter(function(b){ return b.booking_id && b.mentor_id === mentor_id && b.status !== 'cancelled'; });
  var bookedSlots = {};
  bookings.forEach(function(b){
    if (!b.scheduled_at) return;
    var bDuration = parseInt(b.duration_minutes) || durationMinutes;
    var dt = new Date(b.scheduled_at);
    for (var bi = 0; bi < bDuration; bi += 30) {
      var blockTime = new Date(dt.getTime() + bi * 60 * 1000);
      var jst = new Date(blockTime.getTime() + 9 * 60 * 60 * 1000);
      var key = jst.toISOString().slice(0,10) + 'T'
        + ('0'+jst.getUTCHours()).slice(-2) + ':' + ('0'+jst.getUTCMinutes()).slice(-2);
      bookedSlots[key] = true;
    }
  });

  var now        = new Date();
  var rangeStart = new Date(now); rangeStart.setHours(0,0,0,0);
  var rangeEnd   = new Date(rangeStart); rangeEnd.setDate(rangeStart.getDate() + 28); // ★ メンターは4週間分

  // ★ Googleカレンダーの既存予定もブロック（skip_mentor_calendar=true の場合はスキップ）
  var skipCalendar = (e.parameter && e.parameter.skip_mentor_calendar === 'true')
                  || (body && (body.skip_mentor_calendar === true || body.skip_mentor_calendar === 'true'));

  if (!skipCalendar) {
    try {
      var calEmail  = (mentor && mentor.calendar_email) ? mentor.calendar_email.trim() : '';
      var calendar  = calEmail ? (CalendarApp.getCalendarById(calEmail) || CalendarApp.getDefaultCalendar())
                               : CalendarApp.getDefaultCalendar();
      var events    = calendar.getEvents(rangeStart, rangeEnd);
      events.forEach(function(ev){
        if (ev.isAllDayEvent()) return;
        // ★ キャンセル済みイベント（タイトルに【キャンセル1on1】を含む）はスキップ
        if (ev.getTitle().indexOf('【キャンセル1on1】') !== -1) return;
        var cur = new Date(ev.getStartTime());
        cur.setSeconds(0); cur.setMilliseconds(0);
        cur.setMinutes(cur.getMinutes() < 30 ? 0 : 30);
        while (cur < ev.getEndTime()) {
          var jstCur = new Date(cur.getTime() + 9 * 60 * 60 * 1000);
          var dateS  = jstCur.toISOString().slice(0, 10);
          var timeS  = ('0'+jstCur.getUTCHours()).slice(-2) + ':' + ('0'+jstCur.getUTCMinutes()).slice(-2);
          bookedSlots[dateS + 'T' + timeS] = true;
          cur.setMinutes(cur.getMinutes() + 30);
        }
      });
      Logger.log('メンターカレンダー取得: ' + events.length + '件 (' + (calEmail || 'default') + ')');
    } catch(calErr) {
      Logger.log('カレンダー取得エラー（無視）: ' + calErr.message);
    }
  }

  // ★ スロット生成（JST基準・30分刻み・durationMinutes分の空きを確認）
  // メンターは時間帯制限なし（カレンダーブロックのみ）
  var result = {};
  for (var d = new Date(rangeStart); d < rangeEnd; d.setDate(d.getDate() + 1)) {
    var jstD   = new Date(d.getTime() + 9 * 60 * 60 * 1000);
    var dow    = String(jstD.getUTCDay());
    var config = availableConfig[dow];
    if (!config) continue;
    var dateStr    = jstD.toISOString().slice(0, 10);
    var slots      = [];
    var startParts = config.start.split(':').map(Number);
    var endParts   = config.end.split(':').map(Number);

    // ★ JST日付文字列 + タイムゾーン指定でDateを正確に生成
    var current = new Date(dateStr + 'T' + ('0'+startParts[0]).slice(-2) + ':' + ('0'+startParts[1]).slice(-2) + ':00+09:00');
    var endTime = new Date(dateStr + 'T' + ('0'+endParts[0]).slice(-2) + ':' + ('0'+endParts[1]).slice(-2) + ':00+09:00');
    var lastSlotTime = new Date(endTime.getTime() - durationMinutes * 60 * 1000);

    while (current <= lastSlotTime) {
      if (current > now) {
        var jstCurrent = new Date(current.getTime() + 9 * 60 * 60 * 1000);
        var timeStr    = ('0'+jstCurrent.getUTCHours()).slice(-2) + ':' + ('0'+jstCurrent.getUTCMinutes()).slice(-2);
        var slotKey    = dateStr + 'T' + timeStr;

        // ★ durationMinutes 分の空きを確認
        var isAvailable = true;
        for (var sm = 0; sm < durationMinutes; sm += 30) {
          var checkTime = new Date(current.getTime() + sm * 60 * 1000);
          var jstCheck  = new Date(checkTime.getTime() + 9 * 60 * 60 * 1000);
          var checkKey  = jstCheck.toISOString().slice(0,10) + 'T'
            + ('0'+jstCheck.getUTCHours()).slice(-2) + ':' + ('0'+jstCheck.getUTCMinutes()).slice(-2);
          if (bookedSlots[checkKey]) { isAvailable = false; break; }
        }
        slots.push({ time: timeStr, datetime: new Date(current).toISOString(), available: isAvailable });
      }
      current.setMinutes(current.getMinutes() + 30);
    }
    if (slots.length > 0) result[dateStr] = slots;
  }
  return jsonResponse({ mentor_id: mentor_id, slots: result, slot_minutes: durationMinutes });
}


// ============================================================
// F-01: 予約作成
// ============================================================
function handleCreateBooking(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var mentee_id    = auth.payload.user_id;
  var mentor_id    = (body.mentor_id || '').trim();
  var scheduled_at = (body.scheduled_at || '').trim();

  if (!mentor_id || !scheduled_at) return errorResponse('MISSING_FIELDS', 400);

  var users  = cachedSheetToObjects_('users');
  var mentor = users.find(function(u) { return u.user_id === mentor_id; });
  var mentee = users.find(function(u) { return u.user_id === mentee_id; });
  if (!mentor) return errorResponse('MENTOR_NOT_FOUND', 404);
  if (!mentee) return errorResponse('MENTEE_NOT_FOUND', 404);

  // ★ duration_minutes: フロントから明示送信された値を優先
  // 未送信の場合はメンターの default_1on1_duration を使用
  // メンター未設定・30分以下の場合は60分（システムデフォルト）
  var duration_minutes = parseInt(body.duration_minutes) > 0
    ? parseInt(body.duration_minutes)
    : (mentor && parseInt(mentor.default_1on1_duration) > 30
        ? parseInt(mentor.default_1on1_duration)
        : 60);

  var booking_id = 'bk-' + Date.now() + '-' + Math.random().toString(36).substring(2, 9);
  var now = new Date().toISOString();

  appendRow('bookings', {
    booking_id: booking_id,
    mentee_id: mentee_id,
    mentor_id: mentor_id,
    scheduled_at: scheduled_at,
    duration_minutes: duration_minutes,
    status: 'scheduled',
    meet_link: '',
    recording_url: '',
    created_at: now,
    updated_at: now
  });

  var dateStr = toJST_(scheduled_at); // JST変換

  // カレンダー登録
  var calResult = addCalendarEvent_({
    title:               '1on1: ' + mentee.name + ' × ' + mentor.name,
    startIso:            scheduled_at,
    durationMinutes:     duration_minutes,
    meetLink:            '',
    mentorEmail:         mentor.email,                          // @socialshift.work
    mentorCalendarEmail: mentor.calendar_email || '',           // @agent-network.com
    menteeEmail:         mentee.email,
    bookingId:           booking_id,
  });
  // calendar_event_id をbookingsシートに保存
  if (calResult.ok && calResult.eventId) {
    var updates1 = {
      calendar_event_id: calResult.eventId,
      calendar_url:      calResult.htmlLink || '',
      updated_at: new Date().toISOString()
    };
    if (calResult.meetLink) updates1.meet_link = calResult.meetLink;
    updateRowWhere('bookings', 'booking_id', booking_id, updates1);
  } else {
    Logger.log('カレンダー登録失敗（予約は完了）: ' + (calResult.error || ''));
  }

  // ★ 確認用カレンダーイベント作成（mentor.calendar_email がある場合のみ）
  if (mentor.calendar_email) {
    var subResult = addSubCalendarEvent_({
      title:               '1on1: ' + mentee.name + ' × ' + mentor.name,
      startIso:            scheduled_at,
      durationMinutes:     duration_minutes,
      mentorEmail:         mentor.email,
      mentorCalendarEmail: mentor.calendar_email,
      bookingId:           booking_id,
    });
    if (subResult.ok && subResult.eventId) {
      updateRowWhere('bookings', 'booking_id', booking_id, {
        sub_calendar_event_id: subResult.eventId,
        updated_at: new Date().toISOString()
      });
    }
  }

  var calInfo = calResult.ok ? '<li>📅 カレンダーに登録済み（招待メール送信済み）</li>' : '<li>⚠️ カレンダー登録に失敗しました</li>';

  sendMail(mentor.email,
    '【1on1予約完了】' + mentee.name + ' さんとの1on1が予約されました',
    '<h2>1on1予約完了</h2><p>' + mentor.name + ' さん</p>' +
    '<p>' + mentee.name + ' さんとの1on1が予約されました。</p>' +
    '<ul><li>日時: ' + dateStr + '</li><li>時間: ' + duration_minutes + '分</li>' + calInfo + '</ul>'
  );
  sendMail(mentee.email,
    '【1on1予約完了】' + mentor.name + ' さんとの1on1が予約されました',
    '<h2>1on1予約完了</h2><p>' + mentee.name + ' さん</p>' +
    '<p>' + mentor.name + ' さんとの1on1が予約されました。</p>' +
    '<ul><li>日時: ' + dateStr + '</li><li>時間: ' + duration_minutes + '分</li>' + calInfo + '</ul>'
  );

  return jsonResponse({ ok: true, booking_id: booking_id, calendar_ok: calResult.ok });
}

// ============================================================
// F-05: 予約キャンセル
// ============================================================
function handleCancelBooking(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var booking_id = body.booking_id || (e.parameter && e.parameter.id) || '';
  if (!booking_id) return errorResponse('MISSING_BOOKING_ID', 400);

  var bookings = cachedSheetToObjects_('bookings');
  var booking = bookings.find(function(b) { return b.booking_id === booking_id; });
  if (!booking) return errorResponse('BOOKING_NOT_FOUND', 404);

  updateRowWhere('bookings', 'booking_id', booking_id, {
    status: 'cancelled',
    updated_at: new Date().toISOString()
  });

  var users = cachedSheetToObjects_('users');
  var mentor = users.find(function(u) { return u.user_id === booking.mentor_id; });
  var mentee = users.find(function(u) { return u.user_id === booking.mentee_id; });
  var reason = body.reason || '';

  sendMail(mentor ? mentor.email : null,
    '【1on1キャンセル通知】' + toJST_(booking.scheduled_at, 'date') + ' の予約がキャンセルされました',
    '<h2>1on1キャンセル通知</h2><p>以下の予約がキャンセルされました。</p>' +
    '<ul><li>日時: ' + toJST_(booking.scheduled_at) + '</li>' +
    (reason ? '<li>理由: ' + reason + '</li>' : '') + '</ul>'
  );
  sendMail(mentee ? mentee.email : null,
    '【1on1キャンセル通知】' + toJST_(booking.scheduled_at, 'date') + ' の予約がキャンセルされました',
    '<h2>1on1キャンセル通知</h2><p>以下の予約がキャンセルされました。</p>' +
    '<ul><li>日時: ' + toJST_(booking.scheduled_at) + '</li>' +
    (reason ? '<li>理由: ' + reason + '</li>' : '') + '</ul>'
  );

  // カレンダーイベント削除（本体 + 確認用）
  if (booking.calendar_event_id) {
    deleteCalendarEvent_(booking.calendar_event_id);
  }
  if (booking.sub_calendar_event_id) {
    deleteSubCalendarEvent_(booking.sub_calendar_event_id);
  }

  return jsonResponse({ ok: true });
}

// ============================================================
// F-Mentee: レポート一覧
// ============================================================
function handleMenteeReports(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var all = cachedSheetToObjects_('mentor_reports') || sheetToObjects(getSheet('mentor_reports'));
  all = all.filter(function(r) { return r.report_id; });
  var myReports = all
    .filter(function(r) { return r.mentee_id === user_id && r.is_published === 'TRUE'; })
    .sort(function(a, b) { return (b.created_at || '').localeCompare(a.created_at || ''); });

  return jsonResponse({ reports: myReports });
}

// ============================================================
// F-Mentee: 事前レポート提出
// ============================================================
function handlePreReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body       = parseBody_(e);
  var user_id    = auth.payload.user_id;
  var booking_id = (body.booking_id || '').trim();

  // pre_reports シートが存在しなければ自動作成
  var sheet = getSheet('pre_reports');
  if (!sheet) {
    var ss = getSpreadsheet_();
    sheet  = ss.insertSheet('pre_reports');
    sheet.getRange(1, 1, 1, 16).setValues([[
      'report_id', 'user_id', 'target_month', 'session_date',
      'current_project', 'goal_work_6m', 'goal_skill_6m',
      'project_result', 'study_hours', 'study_content',
      'good_points', 'improvement_points',
      'next_month_project_goal', 'next_month_study_goal',
      'submitted_at', 'booking_id'
    ]]);
    Logger.log('pre_reports シートを新規作成しました');
  } else {
    // ★ booking_id 列がなければ末尾に自動追加
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('booking_id') < 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('booking_id');
      invalidateCache_('pre_reports');
      Logger.log('pre_reports に booking_id 列を追加しました');
    }
  }

  // ★ Upsert: booking_id が指定されている場合、既存レポートを確認
  if (booking_id) {
    // 編集可能時間チェック（1on1開始から30分以内）
    var bookings = cachedSheetToObjects_('bookings').filter(function(b){ return b.booking_id; });
    var booking  = bookings.find(function(b){ return b.booking_id === booking_id; });
    if (booking && booking.scheduled_at) {
      var deadline = new Date(new Date(booking.scheduled_at).getTime() + 30 * 60 * 1000);
      if (new Date() > deadline) {
        return errorResponse('EDIT_DEADLINE_PASSED: 1on1開始から30分が経過しているため編集できません。', 403);
      }
    }
    // 既存レポートを検索して UPDATE
    var allReports = sheetToObjects(sheet).filter(function(r){ return r.report_id; });
    var existing   = allReports.find(function(r){ return r.user_id === user_id && r.booking_id === booking_id; });
    if (existing) {
      updateRowWhere('pre_reports', 'report_id', existing.report_id, {
        target_month:            body.target_month            || '',
        session_date:            body.session_date            || '',
        current_project:         body.current_project         || '',
        goal_work_6m:            body.goal_work_6m            || '',
        goal_skill_6m:           body.goal_skill_6m           || '',
        project_result:          body.project_result          || '',
        study_hours:             body.study_hours             || 0,
        study_content:           body.study_content           || '',
        good_points:             body.good_points             || '',
        improvement_points:      body.improvement_points      || '',
        next_month_project_goal: body.next_month_project_goal || '',
        next_month_study_goal:   body.next_month_study_goal   || '',
        submitted_at:            new Date().toISOString()
      });
      invalidateCache_('pre_reports');
      return jsonResponse({ ok: true, report_id: existing.report_id, updated: true });
    }
  }

  // INSERT
  var report_id = 'pr-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  appendRow('pre_reports', {
    report_id:               report_id,
    user_id:                 user_id,
    booking_id:              booking_id,
    target_month:            body.target_month            || '',
    session_date:            body.session_date            || '',
    current_project:         body.current_project         || '',
    goal_work_6m:            body.goal_work_6m            || '',
    goal_skill_6m:           body.goal_skill_6m           || '',
    project_result:          body.project_result          || '',
    study_hours:             body.study_hours             || 0,
    study_content:           body.study_content           || '',
    good_points:             body.good_points             || '',
    improvement_points:      body.improvement_points      || '',
    next_month_project_goal: body.next_month_project_goal || '',
    next_month_study_goal:   body.next_month_study_goal   || '',
    submitted_at:            new Date().toISOString()
  });
  invalidateCache_('pre_reports');
  return jsonResponse({ ok: true, report_id: report_id, updated: false });
}

// ============================================================
// F-Mentee: 事前レポート削除（自分の投稿のみ）
// POST api/mentee/pre-report/delete { report_id }
// ============================================================
function handleDeletePreReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  var user_id   = auth.payload.user_id;
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  var sheet = getSheet('pre_reports');
  if (!sheet) return errorResponse('SHEET_NOT_FOUND', 404);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var ridIdx  = headers.indexOf('report_id');
  var uidIdx  = headers.indexOf('user_id');
  if (ridIdx < 0) return errorResponse('INVALID_SHEET', 500);

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ridIdx]) !== report_id) continue;
    // ★ 自分の投稿のみ削除可（メンティー）
    if (uidIdx >= 0 && String(data[i][uidIdx]) !== user_id) {
      return errorResponse('FORBIDDEN', 403);
    }
    sheet.deleteRow(i + 1);
    invalidateCache_('pre_reports');
    Logger.log('handleDeletePreReport: ' + report_id + ' by ' + user_id);
    return jsonResponse({ ok: true });
  }
  return errorResponse('REPORT_NOT_FOUND', 404);
}

// ============================================================
// Mentor: 担当メンティーの事前レポート削除
// POST api/mentor/pre-report/delete { report_id }
// ============================================================
function handleMentorDeletePreReport(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  var sheet = getSheet('pre_reports');
  if (!sheet) return errorResponse('SHEET_NOT_FOUND', 404);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var ridIdx  = headers.indexOf('report_id');
  if (ridIdx < 0) return errorResponse('INVALID_SHEET', 500);

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ridIdx]) !== report_id) continue;
    sheet.deleteRow(i + 1);
    invalidateCache_('pre_reports');
    Logger.log('handleMentorDeletePreReport: ' + report_id + ' by mentor=' + auth.payload.user_id);
    return jsonResponse({ ok: true });
  }
  return errorResponse('REPORT_NOT_FOUND', 404);
}
// ============================================================
function handleMenteePreReports(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var preSheet = getSheet('pre_reports');
  var reports  = preSheet
    ? sheetToObjects(preSheet).filter(function(r){ return r.report_id && r.user_id === user_id; })
    : [];
  reports.sort(function(a, b){ return (b.submitted_at || '').localeCompare(a.submitted_at || ''); });

  // bookings と突合してスケジュール日時・ステータスを付加
  var bookingMap = {};
  cachedSheetToObjects_('bookings').filter(function(b){ return b.booking_id; })
    .forEach(function(b){ bookingMap[b.booking_id] = b; });

  var result = reports.map(function(r){
    var bk = r.booking_id ? bookingMap[r.booking_id] : null;

    // ★ target_month が Date型→ISO文字列に変換されている場合は YYYY/MM 形式に正規化
    var tm = r.target_month || '';
    if (tm) {
      // ISO形式 "2026-04-01T..." → "2026/04"
      var isoMatch = tm.match(/^(\d{4})-(\d{2})/);
      if (isoMatch) {
        tm = isoMatch[1] + '/' + isoMatch[2];
      }
    }
    return {
      report_id:               r.report_id,
      booking_id:              r.booking_id              || '',
      target_month:            tm,
      session_date:            r.session_date            || '',
      current_project:         r.current_project         || '',
      goal_work_6m:            r.goal_work_6m            || '',
      goal_skill_6m:           r.goal_skill_6m           || '',
      project_result:          r.project_result          || '',
      study_hours:             r.study_hours             || 0,
      study_content:           r.study_content           || '',
      good_points:             r.good_points             || '',
      improvement_points:      r.improvement_points      || '',
      next_month_project_goal: r.next_month_project_goal || '',
      next_month_study_goal:   r.next_month_study_goal   || '',
      submitted_at:            r.submitted_at            || '',
      scheduled_at:            bk ? (bk.scheduled_at || '') : '',
      booking_status:          bk ? (bk.status       || '') : '',
    };
  });
  return jsonResponse({ ok: true, reports: result });
}

// ============================================================
// Mentor: リファラル候補者紹介フォーム送信
// ============================================================
function handleReferralSubmit(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body    = parseBody_(e);
  var user_id = auth.payload.user_id;

  if (!body.referrer_name || !body.candidate_name) {
    return errorResponse('MISSING_REQUIRED_FIELDS: 紹介者名と候補者名は必須です', 400);
  }

  var REFERRAL_SS_ID = '1j-1XL-WEL3MsAWQYeLbE53NvNhaqffq0TzcsbvTndOA';
  try {
    var refSS    = SpreadsheetApp.openById(REFERRAL_SS_ID);
    var refSheet = refSS.getSheetByName('referrals');
    if (!refSheet) {
      refSheet = refSS.insertSheet('referrals');
      refSheet.getRange(1, 1, 1, 11).setValues([[
        'referral_id', 'submitted_by_user_id', 'referrer_name', 'candidate_name',
        'candidate_email', 'candidate_phone', 'relationship',
        'candidate_status', 'candidate_status_other', 'submitted_at', 'notes'
      ]]);
      Logger.log('referrals シートを新規作成しました');
    }

    var referral_id = 'ref-' + Date.now() + '-' + Math.random().toString(36).substring(2, 5);
    var now         = new Date().toISOString();

    refSheet.appendRow([
      referral_id,
      user_id,
      body.referrer_name          || '',
      body.candidate_name         || '',
      body.candidate_email        || '',
      body.candidate_phone        || '',
      body.relationship           || '',
      body.candidate_status       || '',
      body.candidate_status_other || '',
      now,
      body.notes                  || ''
    ]);

    try {
      var statusText = (body.candidate_status || '未選択')
        + (body.candidate_status_other ? '（' + body.candidate_status_other + '）' : '');
      var emailBody = [
        '【リファラル候補者紹介が届きました】',
        '',
        '応募ID　　: ' + referral_id,
        '受信日時　: ' + new Date().toLocaleString('ja-JP'),
        '',
        '■ 紹介者情報',
        '紹介者名　: ' + (body.referrer_name || '—'),
        '',
        '■ 候補者情報',
        '氏名　　　: ' + (body.candidate_name || '—'),
        'メール　　: ' + (body.candidate_email || '—'),
        '電話番号　: ' + (body.candidate_phone || '—'),
        '関係性　　: ' + (body.relationship || '—'),
        '現在の状況: ' + statusText,
        '',
        '確認後、候補者の方へのアプローチをお願いします。',
        '1on1管理システム リファラル機能'
      ].join('\n');
      GmailApp.sendEmail(
        'ss.pr@socialshift.work',
        '【リファラル紹介】' + (body.candidate_name || '候補者') + ' ／ 紹介: ' + (body.referrer_name || ''),
        emailBody,
        { name: 'SocialShift リファラルシステム' }
      );
    } catch(mailErr) {
      Logger.log('メール送信エラー（保存は完了）: ' + mailErr.message);
    }

    return jsonResponse({ ok: true, referral_id: referral_id });
  } catch(ssErr) {
    Logger.log('リファラルスプシ書き込みエラー: ' + ssErr.message);
    return errorResponse('SPREADSHEET_ERROR: ' + ssErr.message, 500);
  }
}

// ============================================================
// F-Mentee: プロフィール更新
// ============================================================
function handleUpdateProfile(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  updateRowWhere('users', 'user_id', auth.payload.user_id, {
    name: body.name || '',
    phone_number: body.phone_number || '',
    updated_at: new Date().toISOString()
  });
  return jsonResponse({ ok: true });
}

// ============================================================
// F-04: レポート公開通知
// ============================================================
function handlePublishReport(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var report_id = body.report_id || (e.parameter && e.parameter.id) || '';
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  var now = new Date().toISOString();
  updateRowWhere('mentor_reports', 'report_id', report_id, {
    is_published: 'TRUE',
    published_at: now
  });

  var reports = sheetToObjects(getSheet('mentor_reports'));
  var report = reports.find(function(r) { return r.report_id === report_id; });
  if (!report) return errorResponse('REPORT_NOT_FOUND', 404);

  var users = cachedSheetToObjects_('users');
  var mentee = users.find(function(u) { return u.user_id === report.mentee_id; });

  sendMail(mentee ? mentee.email : null,
    '【1on1レポート公開】メンターからのレポートが届きました',
    '<h2>1on1レポートが公開されました</h2>' +
    '<p>' + (mentee ? mentee.name : '') + ' さん</p>' +
    '<p>メンターからあなたの1on1レポートが公開されました。</p><hr>' +
    '<h3>AIサマリー</h3><p>' + (report.ai_summary || '') + '</p>' +
    '<h3>メンターからのアドバイス</h3><p>' + (report.ai_advice || '') + '</p>' +
    '<h3>次回の目標</h3><p>' + (report.next_goal || '') + '</p>'
  );
  return jsonResponse({ ok: true });
}

// ============================================================
// F-03: AIレポート生成 (meeting-completed)
// ============================================================
function handleMeetingCompleted(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var booking_id = body.booking_id || '';
  if (!booking_id) return errorResponse('MISSING_BOOKING_ID', 400);

  var bookings = cachedSheetToObjects_('bookings');
  var booking = bookings.find(function(b) { return b.booking_id === booking_id; });
  if (!booking) return errorResponse('BOOKING_NOT_FOUND', 404);

  var preReports = sheetToObjects(getSheet('pre_reports'));
  var preReport = preReports
    .filter(function(r) { return r.user_id === booking.mentee_id; })
    .sort(function(a, b) { return (b.submitted_at || '').localeCompare(a.submitted_at || ''); })[0] || {};

  var report_id = 'rep-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var now = new Date().toISOString();

  var aiResult = { ai_summary: '', ai_advice: '', next_goal: '' };
  try {
    var prompt = '以下の1on1セッションの事前レポートからメンターレポートを作成してください。\n\n' +
      '【事前レポート】\n今月の業務目標: ' + (preReport.current_project || '未記入') + '\n' +
      '6ヶ月業務目標: ' + (preReport.goal_work_6m || '未記入') + '\n' +
      '今月の業務結果: ' + (preReport.project_result || '未記入') + '\n' +
      '良かった点: ' + (preReport.good_points || '未記入') + '\n\n' +
      'JSON形式で返答: {"ai_summary": "...", "ai_advice": "...", "next_goal": "..."}';

    var gasResult = generateTextGemini(prompt);
    var text = gasResult.text || gasResult.content || '';
    var jsonMatch = text.match(/\{[\s\S]*?\}/);
    if (jsonMatch) {
      var parsed = JSON.parse(jsonMatch[0]);
      aiResult.ai_summary = parsed.ai_summary || '';
      aiResult.ai_advice = parsed.ai_advice || '';
      aiResult.next_goal = parsed.next_goal || '';
    }
  } catch(err) { Logger.log('AI generation error: ' + err.message); }

  appendRow('mentor_reports', {
    report_id: report_id,
    booking_id: booking_id,
    mentor_id: booking.mentor_id,
    mentee_id: booking.mentee_id,
    ai_summary: aiResult.ai_summary,
    ai_advice: aiResult.ai_advice,
    next_goal: aiResult.next_goal,
    mentor_edited: '',
    is_published: 'FALSE',
    created_at: now,
    published_at: ''
  });

  var users = cachedSheetToObjects_('users');
  var mentor = users.find(function(u) { return u.user_id === booking.mentor_id; });
  sendMail(mentor ? mentor.email : null,
    '【1on1レポート生成完了】AIレポートが作成されました',
    '<h2>AIレポートが生成されました</h2><p>内容を確認・編集してメンティーに公開してください。</p>' +
    '<h3>AIサマリー</h3><p>' + aiResult.ai_summary + '</p>'
  );

  return jsonResponse({ ok: true, report_id: report_id });
}

// ============================================================
// Mentor: スケジュール保存
// ============================================================
function handleSaveSchedule(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var schedules = (body.schedules || []);
  var mentor_id = auth.payload.user_id;
  var now       = new Date().toISOString();

  // mentor_schedules シートが存在しなければ作成
  var sheet = getSheet('mentor_schedules');
  if (!sheet) {
    var ss = getSpreadsheet_();
    sheet  = ss.insertSheet('mentor_schedules');
    sheet.getRange(1, 1, 1, 7).setValues([[
      'schedule_id', 'mentor_id', 'day_of_week',
      'start_time', 'end_time', 'is_active', 'created_at'
    ]]);
    Logger.log('mentor_schedules シートを新規作成しました');
  }

  // ★ start_time/end_time 列にスプレッドシートが時刻型変換しないようアポストロフィ付き文字列で保存
  // → setNumberFormat('@STRING@') で文字列列として書き込む
  var data0    = sheet.getDataRange().getValues();
  var headers0 = data0[0];
  var stIdx    = headers0.indexOf('start_time') + 1; // 1-based
  var etIdx    = headers0.indexOf('end_time')   + 1;
  // start_time / end_time 列全体を文字列フォーマットに設定
  if (stIdx > 0) sheet.getRange(1, stIdx, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
  if (etIdx > 0) sheet.getRange(1, etIdx, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');

  schedules.forEach(function(s) {
    var schedule_id = 'ms-' + mentor_id + '-' + s.day_of_week;
    // ★ 時刻文字列を確実に HH:MM 形式に正規化
    var startTime = normalizeTimeStr_(s.start_time || '10:00');
    var endTime   = normalizeTimeStr_(s.end_time   || '19:00');

    var row = {
      schedule_id: schedule_id,
      mentor_id:   mentor_id,
      day_of_week: String(s.day_of_week),
      start_time:  startTime,
      end_time:    endTime,
      is_active:   s.is_active !== false ? 'TRUE' : 'FALSE',
      created_at:  now
    };

    // UPSERT: schedule_id が既存なら更新、なければ追記
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var sidIdx  = headers.indexOf('schedule_id');
    var existing = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][sidIdx]) === schedule_id) { existing = i; break; }
    }

    if (existing >= 0) {
      var updRow = headers.map(function(h) {
        return row[h] !== undefined ? row[h] : data[existing][headers.indexOf(h)];
      });
      sheet.getRange(existing + 1, 1, 1, headers.length).setValues([updRow]);
    } else {
      var newRow = headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; });
      sheet.appendRow(newRow);
    }
  });

  invalidateCache_('mentor_schedules');

  // default_1on1_duration を users シートに更新
  if (body.default_duration) {
    updateRowWhere('users', 'user_id', mentor_id, {
      default_1on1_duration: String(body.default_duration)
    });
    invalidateCache_('users');
  }

  return jsonResponse({ ok: true });
}

// ── 時刻文字列正規化ヘルパー ──
// スプレッドシートが時刻型に変換した場合（0〜1の小数）や
// "10:00:00" 形式を "HH:MM" に正規化する
function normalizeTimeStr_(val) {
  if (!val) return '10:00';
  var s = String(val).trim();
  // スプレッドシートの時刻シリアル値（0〜1の小数）が来た場合
  var num = parseFloat(s);
  if (!isNaN(num) && num >= 0 && num < 1) {
    var totalMin = Math.round(num * 24 * 60);
    var h = Math.floor(totalMin / 60);
    var m = totalMin % 60;
    return ('0'+h).slice(-2) + ':' + ('0'+m).slice(-2);
  }
  // "HH:MM:SS" → "HH:MM"
  var parts = s.split(':');
  if (parts.length >= 2) {
    return ('0'+parseInt(parts[0])).slice(-2) + ':' + ('0'+parseInt(parts[1])).slice(-2);
  }
  return s;
}

// ★ 年月文字列正規化: スプレッドシートが Date型に変換した値を "YYYY-MM" 形式に戻す
// 例: "Wed Apr 01 2026 00:00:00 GMT+0900 (日本標準時)" → "2026-04"
// 例: Date オブジェクト → "2026-04"
// 例: "2026-04" → "2026-04"（そのまま）
// 例: "2026/4/1" → "2026-04"
function normalizeYearMonth_(val) {
  if (!val) return '';
  var s = String(val).trim();
  if (!s) return '';
  // すでに YYYY-MM 形式
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  // YYYY/M/D 形式
  var slashMatch = s.match(/^(\d{4})\/(\d{1,2})/);
  if (slashMatch) {
    return slashMatch[1] + '-' + ('0' + slashMatch[2]).slice(-2);
  }
  // Date.toString() 形式 or ISO 形式 → Date オブジェクト経由で変換
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      // JSTで年月を取得（UTC+9）
      var jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
      var y   = jst.getUTCFullYear();
      var m   = ('0' + (jst.getUTCMonth() + 1)).slice(-2);
      return y + '-' + m;
    }
  } catch(e) {}
  return s;
}

// ============================================================
// F-06: アンケート提出＋ギャップ検知
// ============================================================
function handleSurveySubmit(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var booking_id = (body.booking_id || '').trim();
  var progress_rating = parseInt(body.progress_rating);
  var mental_rating = parseInt(body.mental_rating);
  var comment = (body.comment || '').trim();

  if (!booking_id) return errorResponse('MISSING_BOOKING_ID', 400);

  var survey_id = 'sv-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var now = new Date().toISOString();

  appendRow('surveys', {
    survey_id: survey_id,
    booking_id: booking_id,
    user_id: auth.payload.user_id,
    role: auth.payload.role,
    progress_rating: progress_rating,
    mental_rating: mental_rating,
    comment: comment,
    submitted_at: now
  });
  // ★ キャッシュを即時無効化（次のhome取得で最新surveyが反映されるように）
  invalidateCache_('surveys');

  var surveys = sheetToObjects(getSheet('surveys')).filter(function(s) { return s.booking_id === booking_id; });
  var mentorSurvey = surveys.find(function(s) { return s.role === 'mentor'; });
  var menteeSurvey = surveys.find(function(s) { return s.role === 'mentee'; });

  if (mentorSurvey && menteeSurvey) {
    var progressGap = Math.abs(parseInt(mentorSurvey.progress_rating) - parseInt(menteeSurvey.progress_rating));
    var mentalGap = Math.abs(parseInt(mentorSurvey.mental_rating) - parseInt(menteeSurvey.mental_rating));
    var hasGap = progressGap >= 2 || mentalGap >= 2;

    if (hasGap) {
      var bookings = cachedSheetToObjects_('bookings');
      var booking = bookings.find(function(b) { return b.booking_id === booking_id; });
      if (booking) {
        updateRowWhere('users', 'user_id', booking.mentee_id, { status: 'yellow', updated_at: now });
        sendMail(CONFIG.ADMIN_EMAIL,
          '【ギャップ検知アラート】認識のズレが検出されました',
          '<h2>アンケートギャップ検知</h2>' +
          '<p>メンターとメンティーの評価に大きな差があります。</p>' +
          '<ul><li>進捗ギャップ: ' + progressGap + '</li>' +
          '<li>メンタルギャップ: ' + mentalGap + '</li>' +
          '<li>booking_id: ' + booking_id + '</li></ul>'
        );
      }
    }
  }

  return jsonResponse({ ok: true, survey_id: survey_id });
}

// ============================================================
// Admin: メンバー一覧
// ============================================================
function handleAdminMembers(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u.name; });

  var members = users.map(function(u) {
    var tenure = calcTenure_(u.hire_date); // ★ 在籍期間計算
    return {
      user_id:              u.user_id,
      name:                 u.name,
      role:                 u.role || '',
      status:               u.status || 'green',
      workplace:            u.workplace || '',
      work_status:          u.work_status || '',
      mentor_id:            u.mentor_id || '',
      mentor_name:          userMap[u.mentor_id] || '—',
      has_leader_role:      u.has_leader_role || 'FALSE',
      goal_work_6m:         u.goal_work_6m || '',
      goal_skill_6m:        u.goal_skill_6m || '',
      current_project:      u.current_project || '',
      default_1on1_duration:u.default_1on1_duration || '60',
      chat_url:             u.chat_url         || '',
      chat_space_id:        u.chat_space_id    || '',
      chat_webhook_url:     u.chat_webhook_url || '',
      tel_meet_url:         u.tel_meet_url     || '',
      hire_date:            u.hire_date        || '', // ★ 入社日
      tenure_months:        tenure.months,            // ★ 在籍月数（数値）
      tenure_label:         tenure.label,             // ★ 在籍期間（表示用）
    };
  });
  return jsonResponse({ members: members });
}

// ============================================================
// Admin: ユーザー一覧（既存エンドポイント・後方互換）
// ============================================================
function handleAdminUsers(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var users = cachedSheetToObjects_('users')
    .filter(function(u) { return u.user_id && String(u.status).toLowerCase() !== 'deleted'; });
  var result = users.map(function(u) {
    return {
      user_id: u.user_id, name: u.name, email: u.email,
      role: u.role, has_leader_role: u.has_leader_role,
      workplace: u.workplace || '', work_status: u.work_status || '',
      birthday: u.birthday || '', chat_url: u.chat_url || '',
      phone: u.phone_number || '', employment_type: u.work_status || '',
      hourly_wage: u.hourly_wage || '', status: u.status || 'active',
      mentor_id: u.mentor_id || '', leader_id: u.leader_id || '',
      created_at: u.created_at || '', updated_at: u.updated_at || ''
    };
  });
  return jsonResponse({ users: result });
}

// ============================================================
// Admin: ユーザー追加（既存エンドポイント・後方互換）
// ============================================================
function handleAddUser(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  if (!body.email || !body.name || !body.role || !body.password) {
    return errorResponse('MISSING_FIELDS', 400);
  }

  var user_id = 'u-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var now = new Date().toISOString();
  appendRow('users', {
    user_id:              user_id,
    email:                body.email.toLowerCase().trim(),
    name:                 body.name,
    role:                 body.role,
    has_leader_role:      body.has_leader_role ? 'TRUE' : 'FALSE',
    password_hash:        sha256Hash(body.password),
    mentor_id:            body.mentor_id || '',
    leader_id:            '',
    phone_number:         body.phone_number || '',
    workplace:            body.workplace || '',
    work_status:          body.work_status || '',
    hourly_wage:          body.hourly_wage || '',
    status:               'active',
    created_at:           now,
    updated_at:           now,
    birthday:             body.birthday || '',
    chat_url:             body.chat_url || '',
    goal_work_6m:         body.goal_work_6m || '',
    goal_skill_6m:        body.goal_skill_6m || '',
    current_project:      body.current_project || '',
    default_1on1_duration:body.default_1on1_duration || '60',
  });
  return jsonResponse({ ok: true, user_id: user_id });
}

// ============================================================
// Admin: アンケート集計
// ============================================================
function handleAdminSurveys(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var surveys = sheetToObjects(getSheet('surveys')).filter(function(s) { return s.survey_id; });
  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u.name; });

  var pairMap = {};
  surveys.forEach(function(s) {
    var bid = s.booking_id;
    if (!pairMap[bid]) pairMap[bid] = { booking_id: bid, mentor: null, mentee: null };
    if (s.role === 'mentor') pairMap[bid].mentor = s;
    else if (s.role === 'mentee') pairMap[bid].mentee = s;
  });

  var gaps = Object.values(pairMap)
    .filter(function(p) { return p.mentor && p.mentee; })
    .map(function(p) {
      var pg = Math.abs(parseInt(p.mentor.progress_rating) - parseInt(p.mentee.progress_rating));
      var mg = Math.abs(parseInt(p.mentor.mental_rating) - parseInt(p.mentee.mental_rating));
      return {
        booking_id: p.booking_id,
        progress_gap: pg, mental_gap: mg,
        has_gap: pg >= 2 || mg >= 2,
        mentor_name: userMap[p.mentor.user_id] || '',
        mentee_name: userMap[p.mentee.user_id] || '',
        submitted_at: p.mentee.submitted_at || ''
      };
    })
    .sort(function(a, b) { return (b.submitted_at || '').localeCompare(a.submitted_at || ''); });

  return jsonResponse({ surveys: surveys, gaps: gaps });
}

// ============================================================
// Mentor: 自分のアンケート回答一覧（履歴＋未回答）
// ============================================================
function handleMentorSurveys(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var surveys = sheetToObjects(getSheet('surveys')).filter(function(s){
    return s.survey_id && s.user_id === user_id;
  }).sort(function(a, b){ return (b.submitted_at || '').localeCompare(a.submitted_at || ''); });

  var bookingMap = {};
  var allBookings = cachedSheetToObjects_('bookings').filter(function(b){ return b.booking_id; });
  allBookings.forEach(function(b){ bookingMap[b.booking_id] = b; });

  var userMap = {};
  cachedSheetToObjects_('users').filter(function(u){ return u.user_id; })
    .forEach(function(u){ userMap[u.user_id] = u.name; });

  var result = surveys.map(function(s){
    var bk = s.booking_id ? bookingMap[s.booking_id] : null;
    var menteeName = bk ? (userMap[bk.mentee_id] || '') : '';
    return {
      survey_id:        s.survey_id,
      booking_id:       s.booking_id       || '',
      user_id:          s.user_id          || '',
      role:             s.role             || '',
      mentee_id:        bk ? (bk.mentee_id || '') : '',
      mentee_name:      menteeName,                          // ★ bookingのmentee_idから名前を取得
      progress_rating:  s.progress_rating  || 0,
      mental_rating:    s.mental_rating    || 0,
      message_to_admin: s.message_to_admin || '',
      submitted_at:     s.submitted_at     || '',
      scheduled_at:     bk ? (bk.scheduled_at || '') : '',
    };
  });

  // ★ 未回答チェック: 開始時刻を過ぎた自分のbookingで未回答のもの
  var answeredSet = {};
  surveys.forEach(function(s){ if (s.booking_id) answeredSet[s.booking_id] = true; });
  var now     = new Date();
  var pending = allBookings.filter(function(b){
    if (b.mentor_id !== user_id) return false;
    if (answeredSet[b.booking_id]) return false;
    if (b.status === 'cancelled') return false;
    return b.status === 'completed' ||
      ((b.status === 'confirmed' || b.status === 'in_progress') && new Date(b.scheduled_at) <= now);
  }).sort(function(a, b){ return (b.scheduled_at || '').localeCompare(a.scheduled_at || ''); });

  return jsonResponse({ ok: true, surveys: result, pending: pending });
}

// ============================================================
// Admin: コンテンツ
// ============================================================
// ============================================================
// Mentee: 研修コンテンツ一覧（ロール制限なし・全ユーザー閲覧可）
// ============================================================
function handleMenteeContents(e) {
  var auth = requireAuth(e); // ロール制限なし（ログインのみ必要）
  if (auth.error) return errorResponse(auth.error, auth.status);
  var contents = cachedSheetToObjects_('contents')
    .filter(function(c) { return c.content_id; })
    .sort(function(a, b) { return (b.created_at || '').localeCompare(a.created_at || ''); });
  return jsonResponse({ ok: true, contents: contents });
}


function handleAdminContents(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var contents = cachedSheetToObjects_('contents').filter(function(c) { return c.content_id; });
  return jsonResponse({ contents: contents });
}

function handleAddContent(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  if (!body.title || !body.url) return errorResponse('MISSING_FIELDS', 400);
  // ★ category 列が存在しなければ自動追加
  var cSheet = getSheet('contents');
  if (cSheet) {
    var ch = cSheet.getRange(1,1,1,cSheet.getLastColumn()).getValues()[0];
    if (ch.indexOf('category') < 0) {
      cSheet.getRange(1, ch.length+1).setValue('category');
      invalidateCache_('contents');
    }
  }
  var content_id = 'ct-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  appendRow('contents', {
    content_id: content_id, title: body.title || '',
    type: body.type || 'link', url: body.url || '',
    description: body.description || '', duration: body.duration || '',
    level: body.level || '', category: body.category || '',
    created_at: new Date().toISOString()
  });
  invalidateCache_('contents');
  return jsonResponse({ ok: true, content_id: content_id });
}

// ── コンテンツ更新 ──
function handleUpdateContent(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  if (!body.content_id) return errorResponse('MISSING_CONTENT_ID', 400);

  var sheet = getSheet('contents');
  // ★ category 列が存在しなければ自動追加
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  if (headers.indexOf('category') < 0) {
    sheet.getRange(1, headers.length+1).setValue('category');
    headers.push('category');
    invalidateCache_('contents');
  }
  var idIdx = headers.indexOf('content_id');
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(body.content_id)) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return errorResponse('CONTENT_NOT_FOUND', 404);

  var updates = {
    title:       body.title       !== undefined ? String(body.title)       : data[rowIdx][headers.indexOf('title')],
    type:        body.type        !== undefined ? String(body.type)        : data[rowIdx][headers.indexOf('type')],
    url:         body.url         !== undefined ? String(body.url)         : data[rowIdx][headers.indexOf('url')],
    description: body.description !== undefined ? String(body.description) : data[rowIdx][headers.indexOf('description')],
    duration:    body.duration    !== undefined ? String(body.duration)    : data[rowIdx][headers.indexOf('duration')],
    level:       body.level       !== undefined ? String(body.level)       : data[rowIdx][headers.indexOf('level')],
    category:    body.category    !== undefined ? String(body.category)    : data[rowIdx][headers.indexOf('category')] || '',
  };
  var updRow = headers.map(function(h) {
    return updates.hasOwnProperty(h) ? updates[h] : data[rowIdx][headers.indexOf(h)];
  });
  sheet.getRange(rowIdx + 1, 1, 1, headers.length).setValues([updRow]);
  invalidateCache_('contents');
  return jsonResponse({ ok: true });
}

// ── コンテンツ一括インポート（CSV）──
// rows: [{ title, type, url, description, duration, level, category }]
function handleBulkImportContents(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var rows = body.rows;
  if (!rows || !Array.isArray(rows) || rows.length === 0) return errorResponse('MISSING_ROWS', 400);

  // ★ category 列が存在しなければ自動追加
  var cSheet2 = getSheet('contents');
  if (cSheet2) {
    var ch2 = cSheet2.getRange(1,1,1,cSheet2.getLastColumn()).getValues()[0];
    if (ch2.indexOf('category') < 0) {
      cSheet2.getRange(1, ch2.length+1).setValue('category');
      invalidateCache_('contents');
    }
  }

  var added = 0;
  var now   = new Date().toISOString();
  rows.forEach(function(row) {
    if (!row.title || !row.url) return;
    var content_id = 'ct-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
    appendRow('contents', {
      content_id:  content_id,
      title:       String(row.title).trim(),
      type:        String(row.type || 'link'),
      url:         String(row.url).trim(),
      description: String(row.description || ''),
      duration:    String(row.duration    || ''),
      level:       String(row.level       || ''),
      category:    String(row.category    || ''), // ★ カテゴリ
      created_at:  now,
    });
    added++;
  });
  invalidateCache_('contents');
  return jsonResponse({ ok: true, added: added, total: rows.length });
}

// ============================================================
// Admin: リンク集
// ============================================================
function handleAdminLinks(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var links = cachedSheetToObjects_('quick_links').filter(function(l) { return l.link_id; });
  return jsonResponse({ links: links });
}

function handleMenteeLinks(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var links = cachedSheetToObjects_('quick_links').filter(function(l) { return l.link_id; });
  return jsonResponse({ links: links });
}

function handleAddLink(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var link_id = 'ql-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  appendRow('quick_links', {
    link_id: link_id, title: body.title || '',
    url: body.url || '', icon: body.icon || '🔗',
    created_at: new Date().toISOString()
  });
  return jsonResponse({ ok: true, link_id: link_id });
}

// ============================================================
// Leader 担当割り当て
// ============================================================
function handleGetAssignments(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a) { return a.assignment_id; });
  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u; });

  var enriched = assignments.map(function(a) {
    return Object.assign({}, a, {
      leader_name: userMap[a.leader_id] ? userMap[a.leader_id].name : a.leader_id,
      mentee_name: userMap[a.mentee_id] ? userMap[a.mentee_id].name : a.mentee_id
    });
  });

  var leaders = users.filter(function(u) { return u.has_leader_role === 'TRUE' || u.has_leader_role === 'true'; });
  var mentees = users.filter(function(u) { return u.role === 'mentee'; });

  return jsonResponse({ assignments: enriched, leaders: leaders, mentees: mentees });
}

function handleAddAssignment(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  if (!body.leader_id || !body.mentee_id) return errorResponse('MISSING_FIELDS', 400);

  // 重複チェック：同一メンティーへの複数リーダー割り当てを禁止
  var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a){ return a.assignment_id; });
  var existing = assignments.find(function(a){ return a.mentee_id === body.mentee_id; });
  if (existing) {
    var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    var existingLeader = users.find(function(u){ return u.user_id === existing.leader_id; });
    var existingName = existingLeader ? existingLeader.name : existing.leader_id;
    return errorResponse(
      'DUPLICATE_LEADER: このメンティーにはすでに「' + existingName + '」がリーダーとして割り当てられています。先に既存の割り当てを削除してください。',
      409
    );
  }

  // usersシートのleader_idも同期更新
  updateRowWhere('users', 'user_id', body.mentee_id, {
    leader_id:  body.leader_id,
    updated_at: new Date().toISOString(),
  });

  var assignment_id = 'la-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  appendRow('leader_assignments', {
    assignment_id: assignment_id,
    leader_id:     body.leader_id,
    mentee_id:     body.mentee_id,
    created_at:    new Date().toISOString()
  });
  return jsonResponse({ ok: true, assignment_id: assignment_id });
}

function handleDeleteAssignment(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var assignment_id = body.assignment_id || (e.parameter && e.parameter.id) || '';
  if (!assignment_id) return errorResponse('MISSING_ASSIGNMENT_ID', 400);

  var sheet = getSheet('leader_assignments');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var aCol = headers.indexOf('assignment_id');
  var mCol = headers.indexOf('mentee_id');
  var lCol = headers.indexOf('leader_id');
  var deleted_mentee_id = '';
  var deleted_leader_id = '';

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][aCol]) === String(assignment_id)) {
      deleted_mentee_id = String(data[i][mCol] || '');
      deleted_leader_id = String(data[i][lCol] || '');
      sheet.deleteRow(i + 1);
      break;
    }
  }

  // usersシートのleader_idも同期クリア（同じリーダーが設定されている場合のみ）
  if (deleted_mentee_id && deleted_leader_id) {
    var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    var target = users.find(function(u){ return u.user_id === deleted_mentee_id; });
    if (target && String(target.leader_id) === deleted_leader_id) {
      updateRowWhere('users', 'user_id', deleted_mentee_id, {
        leader_id:  '',
        updated_at: new Date().toISOString(),
      });
    }
  }

  return jsonResponse({ ok: true });
}

// ============================================================
// Leader: 担当 Mentee 一覧
// ============================================================
function handleMyAssignments(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var leader_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u; });

  // ① leader_assignments シートから取得
  var assignments = cachedSheetToObjects_('leader_assignments')
    .filter(function(a) { return a.assignment_id && a.leader_id === leader_id; });
  var assignedMenteeIds = assignments.map(function(a) { return a.mentee_id; });

  // ② users シートの leader_id からも取得（leader_assignmentsに未登録の場合を補完）
  var usersWithLeader = users.filter(function(u) {
    return u.role === 'mentee'
      && String(u.leader_id || '') === String(leader_id)
      && assignedMenteeIds.indexOf(u.user_id) < 0; // 重複除外
  });

  // ① と ② をマージ
  var menteeMap = {};
  assignments.forEach(function(a) {
    var u = userMap[a.mentee_id];
    if (u) menteeMap[u.user_id] = { assignment_id: a.assignment_id, user: u };
  });
  usersWithLeader.forEach(function(u) {
    if (!menteeMap[u.user_id]) {
      menteeMap[u.user_id] = { assignment_id: 'direct-' + u.user_id, user: u };
    }
  });

  var mentees = Object.values(menteeMap).map(function(item) {
    var u = item.user;
    return {
      assignment_id: item.assignment_id,
      user_id:       u.user_id,
      name:          u.name          || '',
      email:         u.email         || '',
      status:        u.status        || 'green',
      workplace:     u.workplace     || '',
      chat_url:      u.chat_url      || '',      // ★ 追加
      hire_date:     u.hire_date     || '',      // ★ 追加
      tel_meet_url:  u.tel_meet_url  || '',      // ★ 追加
      phone_number:  u.phone_number  || '',      // ★ 追加
    };
  });

  Logger.log('handleMyAssignments: leader=' + leader_id + ' mentees=' + mentees.length
    + ' (assignments=' + assignments.length + ' users_direct=' + usersWithLeader.length + ')');

  return jsonResponse({ mentees: mentees });
}

// ============================================================
// Mentee: 担当 Leader 情報
// ============================================================
function handleLeaderInfo(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var mentee_id = auth.payload.user_id;

  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var me = users.find(function(u) { return u.user_id === mentee_id; });

  // leader_id はusersシートのleader_idカラムを優先、なければleader_assignmentsを参照
  var leader_id = me && me.leader_id ? me.leader_id : null;
  if (!leader_id) {
    var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a) { return a.assignment_id; });
    var myAssignment = assignments.find(function(a) { return a.mentee_id === mentee_id; });
    if (myAssignment) leader_id = myAssignment.leader_id;
  }
  if (!leader_id) return jsonResponse({ leader: null });

  var leader = users.find(function(u) { return u.user_id === leader_id; });
  return jsonResponse({
    leader: leader ? {
      user_id:      leader.user_id,
      name:         leader.name,
      email:        leader.email,
      phone_number: leader.phone_number || '',
      chat_url:     leader.chat_url     || '',
    } : null,
    tel_meet_url: me ? (me.tel_meet_url || '') : '',  // メンティー自身のTEL用Meet URL
  });
}

// ============================================================
// TEL開始（mentor/mentee共通）
// POST api/mentor/tel-start / api/mentee/tel-start
// ① weekly_callsに記録
// ② メンティーのチャットに通知（リーダー発信時）
// ③ リーダーにメール通知（メンティー発信時）
// ④ tel_meet_urlを返す（クライアントでMeetを開く）
// ============================================================
function handleTelStart(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var caller_id   = auth.payload.user_id;
  var caller_role = auth.payload.role;
  var mentee_id   = (body.mentee_id || caller_id).trim();

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === mentee_id; });
  if (!mentee) return errorResponse('USER_NOT_FOUND', 404);

  var tel_meet_url = mentee.tel_meet_url || '';
  if (!tel_meet_url) return errorResponse('TEL_MEET_URL_NOT_SET: このメンティーにTEL用Meet URLが設定されていません', 400);

  // リーダー特定
  var leader = caller_role === 'mentee'
    ? users.find(function(u){ return u.user_id === mentee.leader_id; })
    : users.find(function(u){ return u.user_id === caller_id; });

  var now     = new Date().toISOString();
  var call_id = 'wc-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);

  // ① weekly_callsに記録
  appendRow('weekly_calls', {
    call_id:      call_id,
    leader_id:    leader ? leader.user_id : '',
    mentee_id:    mentee_id,
    meet_url:     tel_meet_url,
    call_date:    now,
    initiated_by: caller_role,
    created_at:   now,
  });

  // ② メンティーのチャットWebhookに通知（リーダー発信時）
  var notifyResult = { ok: false, reason: 'skipped' };
  if (caller_role !== 'mentee' && mentee.chat_webhook_url) {
    try {
      var leaderName = leader ? leader.name : 'リーダー';
      var msg = '📞 *' + leaderName + 'さんからTELが来ています！*\n'
        + '今すぐ入室してください👇\n' + tel_meet_url;
      var wRes = UrlFetchApp.fetch(mentee.chat_webhook_url, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ text: msg }),
        muteHttpExceptions: true
      });
      notifyResult = { ok: wRes.getResponseCode() === 200 };
    } catch(wErr) {
      notifyResult = { ok: false, reason: wErr.message };
    }
  }

  // ③ リーダーにメール通知（メンティー発信時）
  var mailResult = { ok: false, reason: 'skipped' };
  if (caller_role === 'mentee' && leader && leader.email) {
    try {
      sendMail(leader.email,
        '【TEL着信】' + mentee.name + ' さんからTELが来ています',
        '<h2>📞 ' + mentee.name + ' さんからTELが来ています！</h2>'
        + '<p>' + (leader.name || '') + ' さん、今すぐ入室してください。</p>'
        + '<p><a href="' + tel_meet_url + '" style="font-size:18px;font-weight:bold;color:#5dafa7">'
        + '▶ Meetに入室する</a></p>'
        + '<hr><p style="font-size:12px;color:#999">このメールはシステムから自動送信されました</p>'
      );
      mailResult = { ok: true };
    } catch(mErr) {
      mailResult = { ok: false, reason: mErr.message };
    }
  }

  Logger.log('TEL開始記録: ' + mentee.name + ' call_id=' + call_id + ' by=' + caller_role);
  return jsonResponse({
    ok:            true,
    call_id:       call_id,
    tel_meet_url:  tel_meet_url,
    mentee_name:   mentee.name,
    notify_result: notifyResult,
    mail_result:   mailResult,
  });
}

// ============================================================
// weekly_calls / call_reports
// ============================================================
function handleWeeklyCalls(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;
  var role = auth.payload.role;

  var all = sheetToObjects(getSheet('weekly_calls')).filter(function(c) { return c.call_id; });
  var calls = role === 'admin' ? all : all.filter(function(c) {
    return c.leader_id === user_id || c.mentee_id === user_id;
  });
  return jsonResponse({ calls: calls });
}

function handleCallReports(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;
  var role    = auth.payload.role;

  if (role === 'mentee') return errorResponse('FORBIDDEN', 403);

  var all = sheetToObjects(getSheet('call_reports')).filter(function(r) { return r.report_id; });
  var reports = role === 'admin' ? all : all.filter(function(r) {
    return r.leader_id === user_id;
  });

  // ★ mentee_name / leader_name を付与
  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u.name || u.user_id; });

  var enriched = reports.map(function(r) {
    return {
      report_id:          r.report_id,
      call_id:            r.call_id            || '',
      leader_id:          r.leader_id          || '',
      leader_name:        userMap[r.leader_id] || r.leader_id || '',
      mentee_id:          r.mentee_id          || '',
      mentee_name:        userMap[r.mentee_id] || r.mentee_id || '',
      transcript:         r.transcript         || '',
      ai_summary:         r.ai_summary         || '',
      ai_mentee_status:   r.ai_mentee_status   || 'green',
      next_action:        r.next_action        || '',
      talk_content:       r.talk_content       || '',
      concerns:           r.concerns           || '',
      good_points:        r.good_points        || '',
      memo:               r.memo               || '',
      meet_url:           r.meet_url           || '',
      meet_count:         r.meet_count         || '',
      is_confirmed:       r.is_confirmed       || 'FALSE',
      recording_url:      r.recording_url      || '',
      recording_file_id:  r.recording_file_id  || '',
      transcript_file_url:r.transcript_file_url|| '',
      created_at:         r.created_at         || '',
    };
  });

  return jsonResponse({ reports: enriched });
}

function handleCreateCallReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var role = auth.payload.role;
  // メンティーは作成不可
  if (role === 'mentee') return errorResponse('FORBIDDEN', 403);

  var body = parseBody_(e);

  var report_id    = 'cr-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
  var meet_url     = (body.meet_url || '').trim();
  // is_confirmed: boolean true/false または文字列 'true'/'false' を正規化
  var _ic = body.is_confirmed;
  var is_confirmed = (_ic === false || _ic === 'false' || _ic === 'FALSE') ? 'FALSE' : 'TRUE';

  // このメンティーのMeet実施回数を集計
  var existing = sheetToObjects(getSheet('call_reports'))
    .filter(function(r){ return r.mentee_id === body.mentee_id && r.meet_url; });
  var meet_count = existing.length + (meet_url ? 1 : 0);

  appendRow('call_reports', {
    report_id:         report_id,
    call_id:           body.call_id || '',
    leader_id:         auth.payload.user_id,
    mentee_id:         body.mentee_id || '',
    transcript:        body.transcript || '',
    ai_summary:        body.ai_summary || '',
    ai_mentee_status:  body.ai_mentee_status || '',
    next_action:       body.next_action || '',
    talk_content:      body.talk_content || '',
    concerns:          body.concerns || '',
    good_points:       body.good_points || '',
    memo:              body.memo || '',
    meet_url:          meet_url,
    meet_count:        meet_count,
    is_confirmed:      is_confirmed,
    recording_url:     body.recording_url || '',
    recording_file_id: body.recording_file_id || '',
    created_at:        new Date().toISOString()
  });
  return jsonResponse({ ok: true, report_id: report_id, meet_count: meet_count });
}

// ============================================================
// F-02: 前日リマインド（Time-based トリガーで呼び出す）
// ============================================================
// ============================================================
// 録画・録音アップロードリマインド（毎日10時）
// ① 1on1完了後24h以上 & recording_urlが空 → メンターにリマインド
// ② TELレポート未確定（is_confirmed=FALSE）& 24h以上 → リーダーにリマインド
// ============================================================
function recordingUploadReminder() {
  var now     = new Date();
  var cutoff  = new Date(now.getTime() - 24 * 60 * 60 * 1000); // 24時間前
  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u; });

  var FOLDER_URL = 'https://drive.google.com/drive/folders/' + CONFIG.RECORDINGS_FOLDER_ID;

  // ── ① 1on1録画未アップロードチェック ──
  var bookings = sheetToObjects(getSheet('bookings')).filter(function(b){ return b.booking_id; });
  var missingRecordings = bookings.filter(function(b) {
    if (b.status !== 'completed') return false;
    if (b.recording_url) return false; // アップロード済み
    var completedAt = new Date(b.scheduled_at);
    return completedAt < cutoff; // 24時間以上前に完了
  });

  // メンター別にまとめてメール送信
  var mentorBookings = {};
  missingRecordings.forEach(function(b) {
    var mid = b.mentor_id;
    if (!mentorBookings[mid]) mentorBookings[mid] = [];
    mentorBookings[mid].push(b);
  });

  Object.keys(mentorBookings).forEach(function(mentorId) {
    var mentor = userMap[mentorId];
    if (!mentor || !mentor.email) return;
    var items  = mentorBookings[mentorId];
    var listHtml = items.map(function(b) {
      var mentee  = userMap[b.mentee_id];
      var menteeName = mentee ? mentee.name : b.mentee_id;
      var dtStr = toJST_(b.scheduled_at, 'short'); // JST変換
      return '<li>' + menteeName + ' さんとの1on1（' + dtStr + '実施）</li>';
    }).join('');

    sendMail(mentor.email,
      '【リマインド】1on1録画のアップロードをお願いします（' + items.length + '件）',
      '<h2>📹 1on1録画のアップロードリマインド</h2>'
      + '<p>' + mentor.name + ' さん</p>'
      + '<p>以下の1on1の録画がまだアップロードされていません。<br>'
      + 'Google Meetの録画を下記フォルダにアップロードしてください。</p>'
      + '<ul>' + listHtml + '</ul>'
      + '<p><a href="' + FOLDER_URL + '" style="background:#1a73e8;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;display:inline-block;margin-top:8px">📁 録画フォルダを開く →</a></p>'
      + '<hr><p style="font-size:12px;color:#666">ファイルをフォルダに入れると自動的にリネーム・整理・AI要約が行われます。<br>'
      + '※ Aさん/Bさんのようなメンティー名フォルダがない場合は自動作成されます。</p>'
    );
    Logger.log('録画リマインド送信: ' + mentor.name + ' (' + items.length + '件)');
  });

  // ── ② TELレポート未確定チェック（is_confirmed=FALSE & 24時間以上） ──
  var callReports = sheetToObjects(getSheet('call_reports')).filter(function(r){ return r.report_id; });
  var unconfirmed = callReports.filter(function(r) {
    if (String(r.is_confirmed||'').toUpperCase() === 'TRUE' || r.is_confirmed === true) return false;
    var createdAt = new Date(r.created_at);
    return createdAt < cutoff;
  });

  // ── ③ TEL録音未アップロードチェック（recording_urlが空 & 24時間以上） ──
  var missingTelRecordings = callReports.filter(function(r) {
    if (r.recording_url) return false; // アップロード済み
    var createdAt = new Date(r.created_at);
    return createdAt < cutoff;
  });

  // リーダー別にまとめて送信（②③を1通にまとめる）
  var leaderIssues = {};
  unconfirmed.forEach(function(r) {
    var lid = r.leader_id; if (!lid) return;
    if (!leaderIssues[lid]) leaderIssues[lid] = { unconfirmed:[], missingRec:[] };
    leaderIssues[lid].unconfirmed.push(r);
  });
  missingTelRecordings.forEach(function(r) {
    var lid = r.leader_id; if (!lid) return;
    if (!leaderIssues[lid]) leaderIssues[lid] = { unconfirmed:[], missingRec:[] };
    // unconfirmedと重複しない場合のみ追加
    if (!leaderIssues[lid].missingRec.find(function(x){ return x.report_id===r.report_id; })) {
      leaderIssues[lid].missingRec.push(r);
    }
  });

  Object.keys(leaderIssues).forEach(function(leaderId) {
    var leader = userMap[leaderId];
    if (!leader || !leader.email) return;
    var issue = leaderIssues[leaderId];
    var bodyHtml = '<h2>📋 週次TEL 対応リマインド</h2><p>' + leader.name + ' さん</p>';

    // 録音未アップロード
    if (issue.missingRec.length > 0) {
      var recList = issue.missingRec.map(function(r) {
        var mentee = userMap[r.mentee_id];
        return '<li>' + (mentee ? mentee.name : r.mentee_id) + ' さんのTEL（'
          + toJST_(r.created_at, 'short').split(' ')[0] + '実施）</li>';
      }).join('');
      bodyHtml += '<h3>🎙 録音ファイルの未アップロード（' + issue.missingRec.length + '件）</h3>'
        + '<p>以下のハドル録音をDriveフォルダにアップロードしてください。</p>'
        + '<ul>' + recList + '</ul>'
        + '<p><a href="' + FOLDER_URL + '" style="background:#f9ab00;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;display:inline-block">📁 録音フォルダを開く →</a></p>'
        + '<p style="font-size:12px;color:#666">アップロード後10分以内に文字起こし・AI要約が自動実行されます</p><hr>';
    }

    // レポート未確定
    if (issue.unconfirmed.length > 0) {
      var repList = issue.unconfirmed.map(function(r) {
        var mentee  = userMap[r.mentee_id];
        var dtStr   = toJST_(r.created_at, 'short').split(' ')[0]; // JST変換
        var preview = (r.talk_content || r.ai_summary || '—').slice(0, 40);
        return '<li>' + (mentee ? mentee.name : r.mentee_id) + ' さん（'
          + dtStr + '）: ' + preview + '…</li>';
      }).join('');
      bodyHtml += '<h3>📝 AIレポートの未確定（' + issue.unconfirmed.length + '件）</h3>'
        + '<p>内容を確認・編集のうえ確定してください。</p>'
        + '<ul>' + repList + '</ul>'
        + '<p><a href="https://koheiumeda-arch.github.io/SS1on1/mentor.html#weeklyTel" style="background:#1a73e8;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;display:inline-block">📊 週次TEL画面を開く →</a></p>';
    }

    var totalCount = issue.missingRec.length + issue.unconfirmed.length;
    sendMail(leader.email,
      '【リマインド】週次TEL 未対応 ' + totalCount + '件（録音未UP/レポート未確定）',
      bodyHtml
    );
    Logger.log('TELリマインド送信: ' + leader.name
      + ' / 録音未UP:' + issue.missingRec.length
      + '件 / 未確定:' + issue.unconfirmed.length + '件');
  });

  Logger.log('recordingUploadReminder 完了'
    + ' / 1on1未UP:' + missingRecordings.length
    + '件 / TEL録音未UP:' + missingTelRecordings.length
    + '件 / TEL未確定:' + unconfirmed.length + '件');
}


function dailyReminder() {
  // JSTベースで「明日」の日付を取得
  var targetDate = getTomorrowJST_();

  var bookings = cachedSheetToObjects_('bookings').filter(function(b) {
    // scheduled_at をJSTに変換して日付部分で比較
    if (!b.booking_id || !b.scheduled_at) return false;
    var jstDate = toJST_(b.scheduled_at, 'date').replace(/\//g, '-');
    // YYYY-M-D → YYYY-MM-DD に正規化
    var parts = jstDate.split('-');
    var normalized = parts[0] + '-' + ('0'+parts[1]).slice(-2) + '-' + ('0'+parts[2]).slice(-2);
    return normalized === targetDate;
  });
  if (bookings.length === 0) return;

  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u; });

  bookings.forEach(function(b) {
    var mentor  = userMap[b.mentor_id];
    var mentee  = userMap[b.mentee_id];
    var dtStr   = toJST_(b.scheduled_at, 'label'); // JST表示 例: 3月24日（月） 12:30
    var meetLink= b.meet_link || '';
    var dur     = b.duration_minutes || 60;

    if (mentor) {
      sendMail(mentor.email,
        '【1on1リマインド】明日 ' + dtStr + ' の1on1があります',
        '<h2>1on1リマインダー</h2><p>' + mentor.name + ' さん</p>' +
        '<ul><li>メンティー: ' + (mentee ? mentee.name : '') + '</li>' +
        '<li>日時: ' + dtStr + '</li><li>時間: ' + dur + '分</li></ul>' +
        (meetLink ? '<p><a href="' + meetLink + '">Google Meetリンク</a></p>' : '')
      );
    }
    if (mentee) {
      sendMail(mentee.email,
        '【1on1リマインド】明日 ' + dtStr + ' の1on1があります',
        '<h2>1on1リマインダー</h2><p>' + mentee.name + ' さん</p>' +
        '<ul><li>メンター: ' + (mentor ? mentor.name : '') + '</li>' +
        '<li>日時: ' + dtStr + '</li><li>時間: ' + dur + '分</li></ul>' +
        (meetLink ? '<p><a href="' + meetLink + '">Google Meetリンク</a></p>' : '') +
        '<p><strong>事前レポートの提出をお忘れなく！</strong></p>'
      );
    }
  });
}

// ============================================================
// F-08: 週次 TEL リマインド（毎週金曜 9:00）
// ============================================================
function weeklyTelReminder() {
  var now = getNowJST_(); // JSTベース
  var day = now.getUTCDay();
  var monday = new Date(now);
  monday.setDate(now.getDate() - (day - 1));
  monday.setHours(0,0,0,0);
  var weekStart = monday.toISOString().split('T')[0];
  var friday = new Date(now);
  friday.setHours(23,59,59,999);
  var weekEnd = friday.toISOString().split('T')[0];

  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var leaders = users.filter(function(u) {
    return u.has_leader_role === 'TRUE' || u.has_leader_role === 'true';
  });
  var calls = sheetToObjects(getSheet('weekly_calls')).filter(function(c) { return c.call_id; });

  var doneIds = {};
  calls.filter(function(c) {
    var d = (c.call_date || '').split('T')[0];
    return d >= weekStart && d <= weekEnd && c.status !== 'cancelled';
  }).forEach(function(c) { doneIds[c.leader_id] = true; });

  var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a) { return a.assignment_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u; });

  leaders.filter(function(l) { return !doneIds[l.user_id]; }).forEach(function(leader) {
    var myMentees = assignments
      .filter(function(a) { return a.leader_id === leader.user_id; })
      .map(function(a) { return userMap[a.mentee_id]; })
      .filter(Boolean);
    var menteeList = myMentees.map(function(m) { return '・' + m.name + '（' + (m.workplace || '') + '）'; }).join('\n') || '（担当メンティーなし）';

    sendMail(leader.email,
      '【週次TELリマインド】今週（' + weekStart + '〜' + weekEnd + '）のTELが未実施です',
      '<h2>週次TELリマインダー</h2><p>' + leader.name + ' さん</p>' +
      '<p>今週の週次TELがまだ記録されていません。</p>' +
      '<h3>担当メンティー：</h3><pre>' + menteeList + '</pre>'
    );
  });
}

// ============================================================
// F-09: TEL レポート未提出アラート（毎週月曜 9:00）
// ============================================================
function weeklyTelReportAlert() {
  var now      = getNowJST_(); // JSTベース
  var day      = now.getUTCDay();
  var thisMonday = new Date(now);
  thisMonday.setUTCDate(now.getUTCDate() - (day === 0 ? 6 : day - 1));
  thisMonday.setHours(0,0,0,0);

  var lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);
  var lastSunday = new Date(thisMonday);
  lastSunday.setDate(thisMonday.getDate() - 1);

  var weekStart = lastMonday.toISOString().split('T')[0];
  var weekEnd = lastSunday.toISOString().split('T')[0];

  var calls = sheetToObjects(getSheet('weekly_calls')).filter(function(c) { return c.call_id; });
  var reports = sheetToObjects(getSheet('call_reports')).filter(function(r) { return r.report_id; });
  var users = cachedSheetToObjects_('users').filter(function(u) { return u.user_id; });
  var userMap = {};
  users.forEach(function(u) { userMap[u.user_id] = u; });

  var lastWeekCalls = calls.filter(function(c) {
    var d = (c.call_date || '').split('T')[0];
    return d >= weekStart && d <= weekEnd;
  });

  var reportedIds = {};
  reports.forEach(function(r) { reportedIds[r.call_id] = true; });

  var missing = lastWeekCalls.filter(function(c) { return !reportedIds[c.call_id]; });
  if (missing.length === 0) return;

  missing.forEach(function(c) {
    var leader = userMap[c.leader_id];
    var mentee = userMap[c.mentee_id];
    if (leader) {
      sendMail(leader.email,
        '【TELレポート未提出】' + weekStart + '週の週次TELレポートを提出してください',
        '<h2>週次TELレポート未提出アラート</h2>' +
        '<p>' + leader.name + ' さん</p>' +
        '<p>先週（' + weekStart + '〜' + weekEnd + '）の週次TELレポートが未提出です。</p>' +
        '<ul><li>メンティー: ' + (mentee ? mentee.name : '') + '</li>' +
        '<li>実施日: ' + c.call_date + '</li></ul>'
      );
    }
    sendMail(CONFIG.ADMIN_EMAIL,
      '【管理者アラート】週次TELレポート未提出 ' + c.call_id,
      '<h2>週次TELレポート未提出（管理者通知）</h2>' +
      '<ul><li>リーダー: ' + (leader ? leader.name + ' (' + leader.email + ')' : c.leader_id) + '</li>' +
      '<li>メンティー: ' + (mentee ? mentee.name : c.mentee_id) + '</li>' +
      '<li>実施日: ' + c.call_date + '</li>' +
      '<li>対象週: ' + weekStart + '〜' + weekEnd + '</li></ul>'
    );
  });
}

// ============================================================
// F-10: 月次実績集計（毎月 1 日 0:00）
// ============================================================
function monthlyAggregation() {
  var now = new Date();
  var firstOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastOfLastMonth  = new Date(now.getFullYear(), now.getMonth(), 0);
  var monthStart = firstOfLastMonth.toISOString().split('T')[0];
  var monthEnd   = lastOfLastMonth.toISOString().split('T')[0];
  var targetYm   = firstOfLastMonth.getFullYear() + '-' +
    ('0' + (firstOfLastMonth.getMonth() + 1)).slice(-2);

  var calls   = sheetToObjects(getSheet('weekly_calls')).filter(function(c) { return c.call_id; });
  var reports = sheetToObjects(getSheet('call_reports')).filter(function(r) { return r.report_id; });

  var monthCalls = calls.filter(function(c) {
    var d = (c.call_date || '').split('T')[0];
    return d >= monthStart && d <= monthEnd;
  });

  var reportedIds = {};
  reports.forEach(function(r) { reportedIds[r.call_id] = true; });

  var leaderStats = {};
  monthCalls.forEach(function(c) {
    var lid = c.leader_id || 'unknown';
    if (!leaderStats[lid]) leaderStats[lid] = { leader_id: lid, total: 0, completed: 0, reported: 0 };
    leaderStats[lid].total++;
    if (c.status === 'completed') leaderStats[lid].completed++;
    if (reportedIds[c.call_id]) leaderStats[lid].reported++;
  });

  var total = monthCalls.length;
  var completed = monthCalls.filter(function(c) { return c.status === 'completed'; }).length;
  var reported = monthCalls.filter(function(c) { return reportedIds[c.call_id]; }).length;
  var compRate = total > 0 ? Math.round(completed / total * 100) : 0;
  var repRate  = completed > 0 ? Math.round(reported / completed * 100) : 0;

  var statsStr = JSON.stringify(Object.values(leaderStats));
  var now2 = new Date().toISOString();

  appendRow('admin_memos', {
    memo_id: 'monthly-' + targetYm,
    admin_id: 'system',
    target_id: targetYm,
    content: '月次集計 ' + targetYm + ' | 総TEL:' + total + '件 完了:' + completed + '件(' + compRate + '%) レポート提出:' + reported + '件(' + repRate + '%) | leader_stats:' + statsStr,
    created_at: now2,
    updated_at: now2
  });

  sendMail(CONFIG.ADMIN_EMAIL,
    '【月次集計レポート】' + targetYm + ' 週次TEL実績',
    '<h2>' + targetYm + ' 月次TEL実績サマリー</h2>' +
    '<table border="1" cellpadding="8" style="border-collapse:collapse">' +
    '<tr><th>指標</th><th>実績</th></tr>' +
    '<tr><td>TEL予定件数</td><td>' + total + '件</td></tr>' +
    '<tr><td>実施件数</td><td>' + completed + '件 (' + compRate + '%)</td></tr>' +
    '<tr><td>レポート提出件数</td><td>' + reported + '件 (' + repRate + '%)</td></tr>' +
    '</table><h3>リーダー別実績</h3><pre>' + statsStr + '</pre>'
  );
}

// HTML 側から WebApp URL を取得するための関数
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================================
// handleRequest: 後方互換用エントリポイント
// ============================================================
function handleRequest(path, method, bodyStr, token) {
  try {
    var body = {};
    try { body = JSON.parse(bodyStr || '{}'); } catch(e) {}
    body._token = token;
    var e = {
      parameter: { path: path },
      postData: { contents: JSON.stringify(body) }
    };
    var output = routeRequest(method, path, e);
    // ContentService output → JSONオブジェクトに変換
    try {
      return JSON.parse(output.getContent());
    } catch(pe) {
      return { error: 'parse_error' };
    }
  } catch(err) {
    return { error: err.message };
  }
}

// ============================================================
// Time-based トリガー設定（初回 1 回だけ実行）
// ============================================================
// ============================================================
// ★ ポーリング: 録音・録画ファイルを自動検出して処理
// GASのtime-basedトリガー（10分おき）で実行
// ============================================================

// ── ポーリング処理の重複実行防止用ロック ──
function pollRecordingsMain() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('pollRecordings: 別インスタンスが実行中のためスキップ');
    return;
  }
  try {
    Logger.log('=== pollRecordingsMain 開始 ===');
    // ① 予定時刻を過ぎた scheduled → completed に自動更新
    var completedCount = autoCompleteBookings();
    Logger.log('自動完了処理: ' + completedCount + '件');
    // ② 未回答アンケート通知メール送信
    var notifiedCount = sendSurveyNotifications();
    Logger.log('アンケート通知: ' + notifiedCount + '件');
    // ③ 録画処理
    var result1on1 = pollMeetRecordings();
    var resultTel  = pollTelRecordings();
    Logger.log('1on1録画: ' + result1on1.processed + '件処理');
    Logger.log('TEL録音: '  + resultTel.processed  + '件処理');
  } catch(err) {
    Logger.log('pollRecordingsMain エラー: ' + err.message + '\n' + err.stack);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 予定時刻を過ぎた scheduled → completed に自動更新
// ============================================================
function autoCompleteBookings() {
  var completed = 0;
  try {
    var now      = new Date().toISOString();
    var bookings = cachedSheetToObjects_('bookings');
    bookings.forEach(function(b) {
      if (!b.booking_id || b.status !== 'scheduled') return;
      if (!b.scheduled_at) return;
      // 予定時刻 + duration_minutes を過ぎていたら完了扱い
      var endMs = new Date(b.scheduled_at).getTime() + (parseInt(b.duration_minutes)||60) * 60 * 1000;
      if (Date.now() > endMs) {
        updateRowWhere('bookings', 'booking_id', b.booking_id, {
          status:     'completed',
          updated_at: new Date().toISOString()
        });
        completed++;
      }
    });
    if (completed > 0) invalidateCache_('bookings');
  } catch(err) {
    Logger.log('autoCompleteBookings エラー: ' + err.message);
  }
  return completed;
}

// ============================================================
// 1on1完了後アンケート通知メール送信
// ・completed & 未回答 & 未通知 の予約にメール送信
// ・bookings.survey_notified = 'TRUE' で送信済みフラグ管理
// ============================================================
function sendSurveyNotifications() {
  var notified = 0;
  try {
    var GITHUB_BASE = 'https://koheiumeda-arch.github.io/SS1on1';

    // ★ キャッシュを使わず直接読む（通知済みフラグの更新を確実に反映）
    invalidateCache_('bookings');
    var bookings = cachedSheetToObjects_('bookings').filter(function(b) {
      return b.booking_id
        && b.status === 'completed'
        && String(b.survey_notified || '').toUpperCase() !== 'TRUE';
    });
    if (bookings.length === 0) return 0;

    var users   = cachedSheetToObjects_('users');
    var userMap = {};
    users.forEach(function(u) { userMap[u.user_id] = u; });

    // 既回答済みの booking_id 一覧（mentor・menteeそれぞれの回答状況）
    var surveysData = sheetToObjects(getSheet('surveys') || { getDataRange: function(){ return { getValues: function(){ return [[]]; } }; } });
    var menteeAnswered = {}; // menteeが回答済み
    var mentorAnswered = {}; // mentorが回答済み
    surveysData.forEach(function(s) {
      if (!s.booking_id) return;
      if (s.role === 'mentee') menteeAnswered[s.booking_id] = true;
      if (s.role === 'mentor') mentorAnswered[s.booking_id] = true;
    });

    bookings.forEach(function(b) {
      var bothAnswered = menteeAnswered[b.booking_id] && mentorAnswered[b.booking_id];

      // 両方回答済みならフラグだけ立てて終了（メール不要）
      if (bothAnswered) {
        updateRowWhere('bookings', 'booking_id', b.booking_id, {
          survey_notified: 'TRUE', updated_at: new Date().toISOString()
        });
        return;
      }

      var mentor = userMap[b.mentor_id];
      var mentee = userMap[b.mentee_id];
      var dtStr  = toJST_(b.scheduled_at, 'label');

      // ── Mentee へ通知（未回答の場合のみ）──
      if (!menteeAnswered[b.booking_id] && mentee && mentee.email) {
        var menteeUrl = GITHUB_BASE + '/mentee.html#survey';
        sendMail(mentee.email,
          '【1on1アンケートご回答のお願い】' + dtStr + ' の1on1について',
          '<h2>1on1アンケートにご回答ください</h2>'
          + '<p>' + mentee.name + ' さん</p>'
          + '<p>' + dtStr + ' に実施した' + (mentor ? '（' + mentor.name + ' さんとの）' : '') + '1on1についてアンケートにご回答ください。</p>'
          + '<p>所要時間は約1分です。</p>'
          + '<p style="margin:20px 0">'
          + '<a href="' + menteeUrl + '" style="background:#5dafa7;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block">📋 アンケートに回答する →</a>'
          + '</p>'
          + '<p style="font-size:12px;color:#666">※ログイン後、ホーム画面にアンケートが表示されます。</p>'
        );
      }

      // ── Mentor へ通知（未回答の場合のみ）──
      if (!mentorAnswered[b.booking_id] && mentor && mentor.email) {
        var mentorUrl = GITHUB_BASE + '/mentor.html';
        sendMail(mentor.email,
          '【1on1アンケートご回答のお願い】' + dtStr + ' の1on1について',
          '<h2>1on1アンケートにご回答ください</h2>'
          + '<p>' + mentor.name + ' さん</p>'
          + '<p>' + dtStr + ' に実施した' + (mentee ? '（' + mentee.name + ' さんとの）' : '') + '1on1についてアンケートにご回答ください。</p>'
          + '<p style="margin:20px 0">'
          + '<a href="' + mentorUrl + '" style="background:#5dafa7;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block">📋 アンケート画面を開く →</a>'
          + '</p>'
          + '<p style="font-size:12px;color:#666">※ログイン後、「アンケート」タブから回答できます。</p>'
        );
      }

      // ★ 通知済みフラグを即時更新（次回ポーリングで再送しないよう確実にフラグを立てる）
      updateRowWhere('bookings', 'booking_id', b.booking_id, {
        survey_notified: 'TRUE', updated_at: new Date().toISOString()
      });
      // ★ 書き込み直後にキャッシュを無効化（次回ポーリングで必ず最新を読む）
      invalidateCache_('bookings');
      notified++;
    });

  } catch(err) {
    Logger.log('sendSurveyNotifications エラー: ' + err.message);
  }
  return notified;
}

// ============================================================
// 手動実行用: Meet Recordingsフォルダの未処理ファイルを強制処理
// GASエディタから直接実行してください
// ============================================================
function forceProcessRecordings() {
  Logger.log('=== forceProcessRecordings 開始 ===');

  // ① processed_file_idsをリセット（強制再処理）
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('processed_file_ids');
  Logger.log('processed_file_ids をリセットしました');

  // ② フォルダの中身を確認
  var folder = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
  var files  = folder.getFiles();
  Logger.log('監視フォルダ: ' + folder.getName() + ' (ID: ' + CONFIG.MEET_RECORDINGS_FOLDER_ID + ')');
  Logger.log('格納先フォルダ: ' + CONFIG.RECORDINGS_FOLDER_ID);

  var fileList = [];
  while (files.hasNext()) {
    var f = files.next();
    fileList.push({ id: f.getId(), name: f.getName(), mime: f.getMimeType(), created: f.getDateCreated() });
    Logger.log('  ファイル: ' + f.getName() + ' | MimeType: ' + f.getMimeType() + ' | 作成: ' + f.getDateCreated());
  }
  Logger.log('合計 ' + fileList.length + ' ファイル');

  // ③ pollMeetRecordingsを実行
  var r1 = pollMeetRecordings();
  Logger.log('1on1処理結果: ' + r1.processed + '件');

  // ④ pollTelRecordingsを実行
  var r2 = pollTelRecordings();
  Logger.log('TEL処理結果: ' + r2.processed + '件');

  Logger.log('=== forceProcessRecordings 完了 ===');
}

// ============================================================
// 手動実行用: ファイルを指定して直接処理（デバッグ用）
// ============================================================
function debugOrganizeFile() {
  // ★ここのファイルIDを処理したいファイルのIDに書き換えてから実行
  var FILE_ID    = 'ここにファイルIDを入力';
  var MENTEE_NAME = 'テスト 花子'; // ★ メンティー名を入力
  var RECORD_TYPE = '1on1'; // '1on1' または 'TEL'

  Logger.log('debugOrganizeFile 開始: ' + FILE_ID);
  var result = organizeRecording({
    file_id:     FILE_ID,
    mentee_name: MENTEE_NAME,
    record_type: RECORD_TYPE,
    transcript:  ''
  });
  Logger.log('結果: ' + JSON.stringify(result));
}
// ============================================================
// 1on1 Google Meet録画の自動処理
// ・RECORDINGS_FOLDER_ID 内に新しく追加されたMP4を検知
// ・bookingsとマッチング → organizeRecording → 文字起こし → AI要約 → mentor_reports保存
// ============================================================
function pollMeetRecordings() {
  var processed = 0;
  try {
    // ★ 処理済みファイルIDをScriptPropertiesで管理
    var props        = PropertiesService.getScriptProperties();
    var processedStr = props.getProperty('processed_file_ids') || '{}';
    var processedMap = JSON.parse(processedStr);
    // 7日以上前のIDを削除（肥大化防止）
    var sevenDaysAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
    Object.keys(processedMap).forEach(function(id) {
      if (processedMap[id] < sevenDaysAgo) delete processedMap[id];
    });

    // ★ 監視対象フォルダ内の全ファイルを取得（日時フィルタなし）
    var rootFolder = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
    var files = rootFolder.getFiles();
    var allFiles = [];
    while (files.hasNext()) allFiles.push(files.next());

    // ★ 動画ファイル対象：
    //   パターンA: ファイル名に「1on1:」を含む（Meet生成の元ファイル名）
    //   パターンB: ファイル名に「_1on1」を含む（organizeRecordingでリネーム済み）
    var meetFiles = allFiles.filter(function(f) {
      var mime = f.getMimeType();
      var name = f.getName();
      if (mime === 'application/vnd.google-apps.document') return false;
      var isVideo = mime === 'video/mp4'
                 || name.toLowerCase().endsWith('.mp4')
                 || name.toLowerCase().endsWith('.webm');
      var is1on1  = name.indexOf('1on1:') >= 0
                 || name.indexOf('1on1：') >= 0
                 || name.indexOf('_1on1') >= 0; // ★ リネーム済みパターン追加
      return isVideo && is1on1;
    });

    // 処理済みIDでフィルタ
    meetFiles = meetFiles.filter(function(f) { return !processedMap[f.getId()]; });

    // 1回のポーリングで最大5件まで処理（タイムアウト防止）
    meetFiles = meetFiles.slice(0, 5);

    Logger.log('Meet録画候補（未処理・上限5件）: ' + meetFiles.length + '件 / フォルダ内合計: ' + allFiles.length + '件');

    var users    = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    invalidateCache_('bookings');
    var bookings = cachedSheetToObjects_('bookings')
      .filter(function(b){ return b.booking_id && b.status !== 'cancelled'; });

    meetFiles.forEach(function(file) {
      try {
        var fileCreated = file.getDateCreated().getTime();
        var fileName    = file.getName();

        // ── ファイル名からメンティー名を抽出（2パターン対応）──
        // パターンA（元ファイル）: "1on1: テスト 花子 × 管理者 太郎 - 2026/04/15..."
        // パターンB（リネーム済み）: "2604_テスト 花子_1on1.mp4"
        var fileMenuteeName = '';
        var nameMatchA = fileName.match(/1on1[：:]\s*(.+?)\s*[×x×]\s*(.+?)\s*[-－]/);
        if (nameMatchA) {
          fileMenuteeName = nameMatchA[1].trim();
        } else {
          // リネーム済みパターン: YYMM_メンティー名_1on1
          var nameMatchB = fileName.match(/^\d{4}_(.+?)_1on1/);
          if (nameMatchB) fileMenuteeName = nameMatchB[1].trim();
        }
        var isAlreadyRenamed = fileName.match(/^\d{4}_/) ? true : false;

        Logger.log('1on1マッチング開始: ' + fileName
          + ' | 作成(UTC)=' + new Date(fileCreated).toISOString()
          + ' | 抽出メンティー名="' + fileMenuteeName + '"'
          + ' | リネーム済み=' + isAlreadyRenamed);

        // ── 予約マッチング：時刻 + 名前で照合 ──
        var matched = bookings.find(function(b) {
          if (!b.scheduled_at) return false;
          var bTime = new Date(b.scheduled_at).getTime();
          var withinTime = Math.abs(fileCreated - bTime) < 1440 * 60 * 1000; // ±24時間
          if (!withinTime) return false;
          if (fileMenuteeName) {
            var bMentee = users.find(function(u){ return u.user_id === b.mentee_id; });
            if (bMentee) {
              var bName = bMentee.name.replace(/[\s　]/g,'');
              var fName = fileMenuteeName.replace(/[\s　]/g,'');
              if (bName.indexOf(fName) < 0 && fName.indexOf(bName) < 0) return false;
            }
          }
          return true;
        });

        // フォールバック：名前のみで最近の未処理予約を照合
        if (!matched && fileMenuteeName) {
          matched = bookings.find(function(b) {
            if (!b.scheduled_at || b.recording_url) return false;
            var bMentee = users.find(function(u){ return u.user_id === b.mentee_id; });
            if (!bMentee) return false;
            var bName = bMentee.name.replace(/[\s　]/g,'');
            var fName = fileMenuteeName.replace(/[\s　]/g,'');
            if (bName.indexOf(fName) < 0 && fName.indexOf(bName) < 0) return false;
            var bTime = new Date(b.scheduled_at).getTime();
            return (fileCreated - bTime) < 7 * 24 * 60 * 60 * 1000 && bTime <= fileCreated + 60 * 60 * 1000;
          });
          if (matched) Logger.log('名前マッチ（フォールバック）: ' + fileMenuteeName + ' → ' + matched.booking_id);
        }

        if (!matched) {
          Logger.log('1on1マッチング失敗: ' + fileName
            + ' | bookings件数=' + bookings.length
            + ' → bookingsに対応する予約なし');
          return;
        }

        // すでにrecording_urlが設定されているなら処理済み
        if (matched.recording_url) {
          Logger.log('処理済みスキップ: ' + matched.booking_id);
          return;
        }

        var mentee = users.find(function(u){ return u.user_id === matched.mentee_id; });
        var mentor = users.find(function(u){ return u.user_id === matched.mentor_id; });
        if (!mentee || !mentor) return;

        Logger.log('1on1処理開始: ' + fileName + ' → ' + mentee.name);

        // ① ファイル整理・移動・リネーム
        // ★ リネーム済みの場合はorganizeRecordingをスキップしてfile_urlを直接取得
        var orgResult;
        if (isAlreadyRenamed) {
          // 既にリネーム・移動済み → DriveAPIでファイルURLだけ取得
          var token = ScriptApp.getOAuthToken();
          var meta = driveApiGetFileMeta_(file.getId(), token);
          orgResult = {
            ok:       true,
            file_id:  file.getId(),
            file_url: meta ? (meta.webViewLink || '') : '',
            transcript_file_url: ''
          };
          Logger.log('リネーム済みファイル → organizeRecordingスキップ file_url=' + orgResult.file_url);
        } else {
          orgResult = organizeRecording({
            file_id:            file.getId(),
            mentee_name:        mentee.name,
            mentee_id:          mentee.user_id,
            personal_folder_id: mentee.personal_folder_id || '',  // ★ 明示的に渡す
            record_type:        '1on1',
            transcript:         ''
          });
          if (!orgResult.ok) {
            Logger.log('organizeRecording失敗: ' + orgResult.error);
            return;
          }
        }

        // ② 文字起こし取得
        // Meetが生成するファイル名パターン:
        //   動画:      "1on1: 花子 × 梅田 - 2026/04/07 15:17 JST～Recording"
        //   Geminiメモ: "1on1: 花子 × 梅田 - 2026/04/07 15:17 JST - Gemini によるメモ"
        // → 「 - 」で区切った最初の2要素（タイトル+日時）が一致する
        // → 複数MTGが同時刻でもタイトルで確実にペアリング可能
        var transcript = '';
        var geminiMemoId = '';
        try {
          var videoName = fileName; // ★ 上で宣言済みのfileNameを使用
          var matchKey  = '';

          // ★ リネーム済みファイルはタイトルマッチ不可 → 時刻マッチのみ
          if (!isAlreadyRenamed) {
            var recIdx  = videoName.indexOf('～');
            var dashIdx = videoName.indexOf(' - Gemini');
            var cutIdx  = recIdx >= 0 ? recIdx : (dashIdx >= 0 ? dashIdx : -1);
            matchKey = cutIdx > 0 ? videoName.substring(0, cutIdx).trim() : videoName.substring(0, 50).trim();
            Logger.log('Geminiメモ照合キー: ' + matchKey);
          } else {
            Logger.log('リネーム済みファイル → Geminiメモはタイトルマッチをスキップし時刻マッチのみ');
          }

          var rootFolderForMemo = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
          var allFilesForMemo   = rootFolderForMemo.getFiles();
          while (allFilesForMemo.hasNext()) {
            var mf = allFilesForMemo.next();
            if (mf.getMimeType() !== 'application/vnd.google-apps.document') continue;
            var memoName = mf.getName();
            // タイトルベース名が一致するか確認
            if (matchKey && memoName.indexOf(matchKey) >= 0) {
              geminiMemoId = mf.getId();
              Logger.log('Geminiメモ発見（タイトル一致）: ' + memoName);
              break;
            }
          }

          // タイトルマッチ失敗時のフォールバック: 作成時刻が前後10分以内（より厳格）
          if (!geminiMemoId) {
            Logger.log('タイトルマッチ失敗 → 時刻マッチ（前後10分）でフォールバック');
            var rootFolderForMemo2 = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
            var allFilesForMemo2   = rootFolderForMemo2.getFiles();
            var videoCreatedMs     = file.getDateCreated().getTime();
            var bestDiff = Infinity;
            while (allFilesForMemo2.hasNext()) {
              var mf2 = allFilesForMemo2.next();
              if (mf2.getMimeType() !== 'application/vnd.google-apps.document') continue;
              var diff = Math.abs(videoCreatedMs - mf2.getDateCreated().getTime());
              if (diff < 10 * 60 * 1000 && diff < bestDiff) {
                bestDiff     = diff;
                geminiMemoId = mf2.getId();
                Logger.log('Geminiメモ発見（時刻一致）: ' + mf2.getName() + ' 差分: ' + Math.round(diff/1000) + '秒');
              }
            }
          }
        } catch(memoErr) {
          Logger.log('Geminiメモ検索エラー: ' + memoErr.message);
        }

        if (geminiMemoId) {
          var memoResult = getFileContent({ file_id: geminiMemoId });
          transcript = memoResult.text || '';
          Logger.log('Geminiメモから文字起こし取得: ' + transcript.length + '文字');
          // ★ 取得後にGeminiメモを「元ファイル」フォルダへ移動
          moveToArchiveFolder_(geminiMemoId);
        } else {
          // 最終フォールバック: 動画をDriveでドキュメント変換
          var fcResult = getFileContent({ file_id: orgResult.file_id });
          transcript = fcResult.text || '';
          Logger.log('Drive変換から文字起こし取得: ' + transcript.length + '文字');
        }

        // ③ 文字起こしをテキストファイルとして整理フォルダに保存
        if (transcript) {
          organizeRecording({
            file_id:            orgResult.file_id,
            mentee_name:        mentee.name,
            mentee_id:          mentee.user_id,
            personal_folder_id: mentee.personal_folder_id || '',  // ★ 明示的に渡す
            record_type:        '1on1',
            transcript:         transcript
          });
        }

        // ④ AI要約生成
        var aiResult = generateMeetAiSummary_(transcript, mentee.name, mentor.name);

        // ⑤ bookings に recording_url を保存
        updateRowWhere('bookings', 'booking_id', matched.booking_id, {
          recording_url: orgResult.file_url,
          updated_at:    new Date().toISOString()
        });

        // ⑥ mentor_reports に保存
        var report_id = 'mr-' + Date.now() + '-' + Math.random().toString(36).substring(2,7);
        appendRow('mentor_reports', {
          report_id:               report_id,
          booking_id:              matched.booking_id,
          mentor_id:               matched.mentor_id,
          mentee_id:               matched.mentee_id,
          ai_summary:              aiResult.ai_summary              || '',
          ai_advice:               aiResult.ai_advice               || '',
          next_goal:               aiResult.next_goal               || '',
          next_month_project_goal: aiResult.next_month_project_goal || '',
          next_month_study_goal:   aiResult.next_month_study_goal   || '',
          mentor_edited:           'FALSE',
          is_published:            'FALSE',
          created_at:              new Date().toISOString(),
          published_at:            ''
        });

        // ⑦ メンターにメール通知
        sendMail(mentor.email,
          '【1on1レポート自動生成】' + mentee.name + ' さんとの録画が処理されました',
          '<h2>1on1録画レポートが自動生成されました</h2>'
          + '<p>' + mentor.name + ' さん</p>'
          + '<p>' + mentee.name + ' さんとの1on1録画を処理しました。'
          + '内容を確認・編集してメンティーに公開してください。</p>'
          + '<ul><li><a href="' + orgResult.file_url + '">🎬 録画ファイルを開く</a></li>'
          + (orgResult.transcript_file_url ? '<li><a href="' + orgResult.transcript_file_url + '">📄 文字起こしを開く</a></li>' : '')
          + '</ul>'
          + '<h3>AIサマリー</h3><p>' + (aiResult.ai_summary || '—') + '</p>'
        );

        processedMap[file.getId()] = Date.now();
        processed++;
        Logger.log('1on1処理完了: ' + mentee.name + ' / report_id: ' + report_id);

      } catch(fileErr) {
        Logger.log('1on1ファイル処理エラー: ' + file.getName() + ' / ' + fileErr.message);
      }
    });

    // 処理済みIDを保存
    try { props.setProperty('processed_file_ids', JSON.stringify(processedMap)); } catch(e) {}

    // ★ 動画なし・Geminiメモ単独処理
    // 動画が来ていないがGeminiメモだけある場合に文字起こし・AI要約・レポート作成を行う
    var geminiOnlyFiles = allFiles.filter(function(f) {
      if (processedMap[f.getId()]) return false; // 処理済みスキップ
      var mime = f.getMimeType();
      var name = f.getName();
      if (mime !== 'application/vnd.google-apps.document') return false;
      // 1on1のGeminiメモのみ対象（TELは除外）
      return name.indexOf('1on1:') >= 0 || name.indexOf('1on1：') >= 0;
    }).slice(0, 5);

    if (geminiOnlyFiles.length > 0) {
      Logger.log('Geminiメモ単独処理候補: ' + geminiOnlyFiles.length + '件');
      var users2   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
      invalidateCache_('bookings');
      var bookings2 = cachedSheetToObjects_('bookings')
        .filter(function(b){ return b.booking_id && b.status !== 'cancelled'; });

      geminiOnlyFiles.forEach(function(memoFile) {
        try {
          var memoName    = memoFile.getName();
          var memoCreated = memoFile.getDateCreated().getTime();

          // ファイル名からメンティー名を抽出
          // "1on1: テスト 花子 × 石井翔子 - 2026/04/20 14:00 JST - Gemini..."
          var nameMatch = memoName.match(/1on1[：:]\s*(.+?)\s*[×x×]\s*(.+?)\s*[-－]/);
          var fileMenuteeName = nameMatch ? nameMatch[1].trim() : '';
          Logger.log('Geminiメモ単独処理: ' + memoName + ' メンティー名="' + fileMenuteeName + '"');

          // 予約マッチング
          var matched2 = bookings2.find(function(b) {
            if (!b.scheduled_at) return false;
            var bTime = new Date(b.scheduled_at).getTime();
            if (Math.abs(memoCreated - bTime) > 1440 * 60 * 1000) return false;
            if (!fileMenuteeName) return true;
            var bMentee = users2.find(function(u){ return u.user_id === b.mentee_id; });
            if (!bMentee) return true;
            var bName = bMentee.name.replace(/[\s　]/g,'');
            var fName = fileMenuteeName.replace(/[\s　]/g,'');
            return bName.indexOf(fName) >= 0 || fName.indexOf(bName) >= 0;
          });

          // フォールバック：名前のみ
          if (!matched2 && fileMenuteeName) {
            matched2 = bookings2.find(function(b) {
              var bMentee = users2.find(function(u){ return u.user_id === b.mentee_id; });
              if (!bMentee) return false;
              var bName = bMentee.name.replace(/[\s　]/g,'');
              var fName = fileMenuteeName.replace(/[\s　]/g,'');
              return bName.indexOf(fName) >= 0 || fName.indexOf(bName) >= 0;
            });
          }

          if (!matched2) {
            Logger.log('Geminiメモ単独: 予約マッチなし → スキップ: ' + memoName);
            return;
          }

          var mentee2 = users2.find(function(u){ return u.user_id === matched2.mentee_id; });
          var mentor2 = users2.find(function(u){ return u.user_id === matched2.mentor_id; });
          if (!mentee2 || !mentor2) {
            Logger.log('Geminiメモ単独: ユーザー情報取得失敗');
            return;
          }

          // 文字起こし取得
          var memoResult2 = getFileContent({ file_id: memoFile.getId() });
          var transcript2 = memoResult2.text || '';
          if (!transcript2) {
            Logger.log('Geminiメモ単独: 文字起こし空 → スキップ');
            return;
          }
          Logger.log('Geminiメモ単独: 文字起こし取得 ' + transcript2.length + '文字');

          // 文字起こしをテキストファイルとして個人フォルダへ保存
          var mentee2FolderId = (mentee2.personal_folder_id || '').trim();
          if (!mentee2FolderId) {
            var token2 = ScriptApp.getOAuthToken();
            mentee2FolderId = driveApiGetOrCreateFolder_(CONFIG.INDIVIDUAL_FOLDER_ROOT_ID, mentee2.name, token2);
            if (mentee2FolderId) {
              updateRowWhere('users', 'user_id', mentee2.user_id, { personal_folder_id: mentee2FolderId });
              invalidateCache_('users');
            }
          }
          organizeRecording({
            file_id:            memoFile.getId(),
            mentee_name:        mentee2.name,
            mentee_id:          mentee2.user_id,
            personal_folder_id: mentee2FolderId,  // ★ 明示的に渡す
            record_type:        '1on1',
            transcript:         transcript2,
            memo_only:          true,
          });

          // AI要約生成
          var aiResult2 = generateMeetAiSummary_(transcript2, mentee2.name, mentor2.name);

          // mentor_reports に保存
          var existingReport2 = sheetToObjects(getSheet('mentor_reports'))
            .find(function(r){ return r.booking_id === matched2.booking_id; });

          if (!existingReport2) {
            var report_id2 = 'rpt-' + Date.now() + '-' + Math.random().toString(36).slice(2,7);
            appendRow('mentor_reports', {
              report_id:               report_id2,
              booking_id:              matched2.booking_id,
              mentor_id:               mentor2.user_id,
              mentee_id:               mentee2.user_id,
              ai_summary:              aiResult2.ai_summary              || '',
              ai_advice:               aiResult2.ai_advice               || '',
              next_goal:               aiResult2.next_goal               || '',
              next_month_project_goal: aiResult2.next_month_project_goal || '',
              next_month_study_goal:   aiResult2.next_month_study_goal   || '',
              mentor_edited:           '',
              is_published:            'FALSE',
              created_at:              new Date().toISOString(),
              published_at:            '',
            });
            Logger.log('Geminiメモ単独: mentor_reports 新規作成 ' + report_id2);
          } else {
            updateRowWhere('mentor_reports', 'report_id', existingReport2.report_id, {
              ai_summary:              aiResult2.ai_summary              || '',
              ai_advice:               aiResult2.ai_advice               || '',
              next_goal:               aiResult2.next_goal               || '',
              next_month_project_goal: aiResult2.next_month_project_goal || '',
              next_month_study_goal:   aiResult2.next_month_study_goal   || '',
              updated_at:              new Date().toISOString(),
            });
            Logger.log('Geminiメモ単独: mentor_reports 更新 ' + existingReport2.report_id);
          }
          invalidateCache_('mentor_reports');

          // Geminiメモを「元ファイル」フォルダへ移動
          moveToArchiveFolder_(memoFile.getId());

          processedMap[memoFile.getId()] = Date.now();
          processed++;
          Logger.log('Geminiメモ単独処理完了: ' + mentee2.name);

        } catch(memoErr) {
          Logger.log('Geminiメモ単独処理エラー: ' + memoFile.getName() + ' / ' + memoErr.message);
        }
      });

      try { props.setProperty('processed_file_ids', JSON.stringify(processedMap)); } catch(e) {}
    }

  } catch(err) {
    Logger.log('pollMeetRecordings エラー: ' + err.message);
  }
  return { processed: processed };
}

// ============================================================
// TEL ハドル録音の自動処理（F-12相当）
// ・RECORDINGS_FOLDER_ID 内に新しく追加された音声ファイルを検知
// ・usersのleader_assignmentsとマッチング → organizeRecording
// → 文字起こし → AI要約（4項目）→ call_reports保存
// ============================================================
function pollTelRecordings() {
  var processed = 0;
  try {
    // ★ 監視対象フォルダ内の全ファイルを取得（日時フィルタなし）
    var rootFolder = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
    var files    = rootFolder.getFiles();
    var allFiles = [];
    while (files.hasNext()) allFiles.push(files.next());

    // ★ 音声ファイル かつ ファイル名に「TEL:」を含むもののみ対象
    // Googleドキュメント（Geminiメモ）は除外
    // Chatスペース名「TEL: メンティー名」がそのままファイル名になる
    var audioMimes = ['audio/mp4','audio/mpeg','audio/ogg','audio/webm','audio/m4a'];
    var telFiles = allFiles.filter(function(f) {
      var mime = f.getMimeType();
      var name = f.getName();
      // Googleドキュメントは除外
      if (mime === 'application/vnd.google-apps.document') return false;
      var isAudio = audioMimes.indexOf(mime) >= 0
        || name.toLowerCase().endsWith('.m4a') || name.toLowerCase().endsWith('.mp3')
        || (name.toLowerCase().endsWith('.mp4') && mime.indexOf('audio') >= 0);
      var isTel   = name.indexOf('TEL:') >= 0 || name.indexOf('TEL：') >= 0;
      return isAudio && isTel;
    });

    // 処理済みファイルIDで重複チェック（7日以内のみ保持）
    var props2        = PropertiesService.getScriptProperties();
    var processedStr2 = props2.getProperty('processed_file_ids') || '{}';
    var processedMap2 = JSON.parse(processedStr2);
    var sevenDaysAgo2 = Date.now() - 7 * 24 * 60 * 60 * 1000;
    Object.keys(processedMap2).forEach(function(id) {
      if (processedMap2[id] < sevenDaysAgo2) delete processedMap2[id];
    });
    telFiles = telFiles.filter(function(f) { return !processedMap2[f.getId()]; });

    // 1回最大5件
    telFiles = telFiles.slice(0, 5);

    Logger.log('TEL録音候補（未処理・上限5件）: ' + telFiles.length + '件 / フォルダ内合計: ' + allFiles.length + '件');

    var users       = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    // weekly_callsから直近3日分を取得（TEL開始ボタンで記録されたもの）
    invalidateCache_('weekly_calls');
    var weeklyCalls = sheetToObjects(getSheet('weekly_calls'))
      .filter(function(c){ return c.call_id && c.call_date; });

    telFiles.forEach(function(file) {
      try {
        var fileCreated = file.getDateCreated().getTime();
        var fileName    = file.getName();
        var mentee = null;
        var leader = null;

        // ── メンティー特定：優先1 weekly_calls照合 ──
        // TEL開始ボタンを押した時刻（call_date）と動画作成時刻を照合（±60分）
        var bestCall = null;
        var bestDiff = Infinity;
        weeklyCalls.forEach(function(c) {
          var callMs = new Date(c.call_date).getTime();
          var diff   = Math.abs(fileCreated - callMs);
          if (diff < 60 * 60 * 1000 && diff < bestDiff) {
            bestDiff = diff;
            bestCall = c;
          }
        });

        if (bestCall) {
          mentee = users.find(function(u){ return u.user_id === bestCall.mentee_id; });
          leader = users.find(function(u){ return u.user_id === bestCall.leader_id; });
          Logger.log('weekly_callsマッチ: ' + (mentee ? mentee.name : '?') + ' diff=' + Math.round(bestDiff/1000/60) + '分');
        }

        // ── メンティー特定：優先2 ファイル名から「TEL: メンティー名」を抽出 ──
        if (!mentee) {
          var telNameMatch = fileName.match(/TEL[：:]\s*(.+?)(?:\s*[～\-_\.\(]|$)/);
          var telMenteeName = telNameMatch ? telNameMatch[1].trim() : '';
          if (telMenteeName) {
            mentee = users.find(function(u) {
              return u.name && u.name.replace(/\s/g,'') === telMenteeName.replace(/\s/g,'');
            });
            if (mentee) {
              leader = users.find(function(u){ return u.user_id === mentee.leader_id; });
              Logger.log('ファイル名マッチ: ' + mentee.name);
            }
          }
        }

        if (!leader || !mentee) {
          Logger.log('TELマッチング失敗: ' + fileName);
          return;
        }

        Logger.log('TEL処理開始: ' + fileName + ' → ' + mentee.name);

        // ① 動画ファイルを整理フォルダに移動・リネーム
        var orgResult = organizeRecording({
          file_id:            file.getId(),
          mentee_name:        mentee.name,
          mentee_id:          mentee.user_id,
          personal_folder_id: mentee.personal_folder_id || '',  // ★ 明示的に渡す
          record_type:        'TEL',
          transcript:         ''
        });
        if (!orgResult.ok) {
          Logger.log('organizeRecording失敗: ' + orgResult.error);
          return;
        }

        // ② Geminiメモをファイル名の時刻で照合して文字起こし取得
        // Geminiメモ名: "2026/04/08 10:58 JST に開始した会議 - Gemini によるメモ"
        var transcript  = '';
        var geminiMemoId = '';
        try {
          var rootFolderForMemo = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
          var allFiles = rootFolderForMemo.getFiles();
          var bestMemoDiff = Infinity;

          while (allFiles.hasNext()) {
            var mf = allFiles.next();
            if (mf.getMimeType() !== 'application/vnd.google-apps.document') continue;
            var memoName = mf.getName();

            // Geminiメモのファイル名から開始時刻を抽出
            // 例: "2026/04/08 10:58 JST に開始した会議"
            var timeMatch = memoName.match(/(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})\s+JST/);
            if (timeMatch) {
              var memoTimeStr = timeMatch[1].replace(/\//g, '-') + ':00+09:00';
              var memoTimeMs  = new Date(memoTimeStr).getTime();
              var diff = Math.abs(fileCreated - memoTimeMs);
              // 動画作成時刻と±30分以内で最も近いGeminiメモを選択
              if (diff < 30 * 60 * 1000 && diff < bestMemoDiff) {
                bestMemoDiff = diff;
                geminiMemoId = mf.getId();
                Logger.log('Geminiメモ候補: ' + memoName + ' diff=' + Math.round(diff/1000/60) + '分');
              }
            }
          }
        } catch(memoErr) {
          Logger.log('Geminiメモ検索エラー: ' + memoErr.message);
        }

        if (geminiMemoId) {
          var memoResult = getFileContent({ file_id: geminiMemoId });
          transcript = memoResult.text || '';
          Logger.log('Geminiメモから文字起こし取得: ' + transcript.length + '文字');
          // ★ 取得後にGeminiメモを「元ファイル」フォルダへ移動
          moveToArchiveFolder_(geminiMemoId);
        } else {
          Logger.log('Geminiメモが見つからないためAI要約をスキップ');
        }

        // ③ 文字起こしをテキストファイルとして保存
        if (transcript) {
          organizeRecording({
            file_id:            orgResult.file_id,
            mentee_name:        mentee.name,
            mentee_id:          mentee.user_id,
            personal_folder_id: mentee.personal_folder_id || '',  // ★ 明示的に渡す
            record_type:        'TEL',
            transcript:         transcript
          });
        }

        // ④ AI要約（Geminiメモのテキストを使用）
        var aiResult = generateTelAiSummary_(transcript, leader.name, mentee.name);

        // ⑤ Meet実施回数をカウント
        var existingReps = sheetToObjects(getSheet('call_reports'))
          .filter(function(r){ return r.mentee_id === mentee.user_id && (r.meet_url || r.recording_url); });
        var meetCount = existingReps.length + 1;

        // ⑥ call_reportsに保存（weekly_callsのcall_idを紐付け）
        var report_id = 'cr-' + Date.now() + '-' + Math.random().toString(36).substring(2,7);
        appendRow('call_reports', {
          report_id:           report_id,
          call_id:             bestCall ? bestCall.call_id : '',
          leader_id:           leader.user_id,
          mentee_id:           mentee.user_id,
          transcript:          transcript.substring(0, 5000),
          ai_summary:          aiResult.ai_summary        || '',
          ai_mentee_status:    aiResult.ai_mentee_status  || 'green',
          next_action:         aiResult.memo              || '',
          talk_content:        aiResult.talk_content      || '',
          concerns:            aiResult.concerns          || '',
          good_points:         aiResult.good_points       || '',
          memo:                aiResult.memo              || '',
          meet_url:            orgResult.file_url         || '',
          meet_count:          meetCount,
          is_confirmed:        'FALSE',
          recording_url:       orgResult.file_url         || '',
          recording_file_id:   orgResult.file_id          || '',
          transcript_file_url: orgResult.transcript_file_url || '',
          created_at:          new Date().toISOString()
        });

        // ⑦ リーダーにメール通知
        sendMail(leader.email,
          '【週次TELレポート自動生成】' + mentee.name + ' さんとの録音が処理されました',
          '<h2>週次TELレポートが自動生成されました</h2>'
          + '<p>' + leader.name + ' さん</p>'
          + '<p>' + mentee.name + ' さんとのTEL録音を処理しました。内容を確認・編集して確定してください。</p>'
          + '<ul>'
          + '<li><a href="' + orgResult.file_url + '">🎙 録音ファイルを開く</a></li>'
          + (orgResult.transcript_file_url ? '<li><a href="' + orgResult.transcript_file_url + '">📄 文字起こしを開く</a></li>' : '')
          + '</ul>'
          + (aiResult.talk_content ? '<h3>📋 話した内容</h3><p>' + aiResult.talk_content + '</p>' : '')
          + (aiResult.concerns     ? '<h3>⚠️ 悩み・課題</h3><p>' + aiResult.concerns     + '</p>' : '')
          + (aiResult.good_points  ? '<h3>✅ 良かった点</h3><p>' + aiResult.good_points  + '</p>' : '')
          + '<hr><p><small>サマリー: ' + (aiResult.ai_summary || '—') + '</small></p>'
        );

        processedMap2[file.getId()] = Date.now();
        processed++;
        Logger.log('TEL処理完了: ' + mentee.name + ' / report_id: ' + report_id);

      } catch(fileErr) {
        Logger.log('TELファイル処理エラー: ' + file.getName() + ' / ' + fileErr.message);
      }
    });

    // 処理済みIDを保存
    try { props2.setProperty('processed_file_ids', JSON.stringify(processedMap2)); } catch(e) {}

  } catch(err) {
    Logger.log('pollTelRecordings エラー: ' + err.message);
  }
  return { processed: processed };
}

// ============================================================
// AI プロンプト管理
// ScriptProperties に PROMPT_1ON1 / PROMPT_TEL として保存
// プレースホルダー: {{mentor_name}} {{mentee_name}} {{leader_name}} {{transcript}}
// ============================================================

var DEFAULT_PROMPT_1ON1 = '以下はメンターとメンティーの1on1ミーティングの文字起こしです。\n\n'
  + 'メンター: {{mentor_name}}\n'
  + 'メンティー: {{mentee_name}}\n\n'
  + '文字起こし:\n{{transcript}}\n\n'
  + '以下のJSON形式のみで返答してください（前後の説明文不要、JSONのみ出力）。\n'
  + '各フィールドは簡潔に。\n\n'
  + '{"ai_summary":"1on1全体のサマリー150字以内","ai_advice":"メンターからのアドバイス150字以内","next_goal":"次回までの目標・アクション150字以内","next_month_project_goal":"来月の業務目標100字以内（不明なら空文字）","next_month_study_goal":"来月の学習目標100字以内（不明なら空文字）"}';

var DEFAULT_PROMPT_TEL = '以下はリーダーとメンティーの週次電話（ハドル）の文字起こしです。\n\n'
  + 'リーダー: {{leader_name}}\n'
  + 'メンティー: {{mentee_name}}\n\n'
  + '文字起こし:\n{{transcript}}\n\n'
  + '以下のJSON形式のみで返答してください（他の文章不要）。\n'
  + '※ 出力文中で人物を指す際は「メンティー」ではなく実際の名前（{{mentee_name}}、{{leader_name}}）を使ってください。\n\n'
  + '{"talk_content":"話した内容(200字以内)","concerns":"悩み・課題(なければ空文字)","good_points":"良かった点・成長(なければ空文字)","memo":"メモ・フォローアクション(なければ空文字)","ai_summary":"全体サマリー(100字以内)","ai_mentee_status":"green/yellow/red"}';

// プロンプトをScriptPropertiesから取得（なければデフォルト）
function getPrompt_(key, defaultVal) {
  var stored = PropertiesService.getScriptProperties().getProperty(key);
  return (stored && stored.length > 10) ? stored : defaultVal;
}

// GET api/admin/ai-prompts
function handleGetAiPrompts(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  return jsonResponse({
    ok: true,
    prompt_1on1: getPrompt_('PROMPT_1ON1', DEFAULT_PROMPT_1ON1),
    prompt_tel:  getPrompt_('PROMPT_TEL',  DEFAULT_PROMPT_TEL),
    default_1on1: DEFAULT_PROMPT_1ON1,
    default_tel:  DEFAULT_PROMPT_TEL,
  });
}

// POST api/admin/ai-prompts { prompt_1on1, prompt_tel }
function handleSaveAiPrompts(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body  = parseBody_(e);
  var props = PropertiesService.getScriptProperties();
  if (body.prompt_1on1 !== undefined) props.setProperty('PROMPT_1ON1', String(body.prompt_1on1));
  if (body.prompt_tel  !== undefined) props.setProperty('PROMPT_TEL',  String(body.prompt_tel));
  Logger.log('AIプロンプト更新: 1on1=' + (body.prompt_1on1 ? body.prompt_1on1.length+'文字' : '未変更')
           + ' / TEL=' + (body.prompt_tel ? body.prompt_tel.length+'文字' : '未変更'));
  return jsonResponse({ ok: true });
}

// ============================================================
// AI要約生成ヘルパー（1on1 Meet用）
// ============================================================
function generateMeetAiSummary_(transcript, menteeName, mentorName) {
  if (!transcript) return { ai_summary:'', ai_advice:'', next_goal:'', next_month_project_goal:'', next_month_study_goal:'' };
  try {
    var tmpl   = getPrompt_('PROMPT_1ON1', DEFAULT_PROMPT_1ON1);
    var prompt = tmpl
      .replace(/\{\{mentor_name\}\}/g,  mentorName  || '')
      .replace(/\{\{mentee_name\}\}/g,  menteeName  || '')
      .replace(/\{\{transcript\}\}/g,   transcript.substring(0, 8000));
    var result = generateTextGemini(prompt);
    var text   = result.text || '';
    Logger.log('generateMeetAiSummary_ Gemini応答長: ' + text.length + '文字');
    Logger.log('generateMeetAiSummary_ Gemini応答全文: ' + text);  // ★ 全文ログ（デバッグ用）

    // ★ コードブロック除去
    text = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    // ★ JSONパース試行1: そのまま
    var parsed = null;
    var match = text.match(/\{[\s\S]*\}/);
    if (match) {
      var jsonStr = match[0].replace(/"([^"]+)"\s*:/g, function(m, key) {
        return '"' + key.replace(/\s+/g, '_') + '":';
      });
      try {
        parsed = JSON.parse(jsonStr);
        Logger.log('generateMeetAiSummary_ JSON parse成功（試行1）');
      } catch(e1) {
        Logger.log('generateMeetAiSummary_ 試行1失敗: ' + e1.message);
        // ★ JSONパース試行2: 値内の改行をスペースに置換してから再試行
        try {
          var sanitized = jsonStr
            .replace(/:\s*"([\s\S]*?)(?="[\s,\}])/g, function(m, val) {
              return ': "' + val.replace(/\n/g, ' ').replace(/\r/g, '').replace(/\t/g, ' ');
            });
          parsed = JSON.parse(sanitized + '"');
          Logger.log('generateMeetAiSummary_ JSON parse成功（試行2 sanitize）');
        } catch(e2) {
          Logger.log('generateMeetAiSummary_ 試行2失敗: ' + e2.message);
        }
      }
    }

    if (parsed) {
      return {
        ai_summary:              String(parsed.ai_summary              || ''),
        ai_advice:               String(parsed.ai_advice               || ''),
        next_goal:               String(parsed.next_goal               || ''),
        next_month_project_goal: String(parsed.next_month_project_goal || ''),
        next_month_study_goal:   String(parsed.next_month_study_goal   || ''),
      };
    }

    // ★ JSONパース完全失敗 → 各フィールドを正規表現で個別抽出
    Logger.log('generateMeetAiSummary_ JSON全失敗 → 正規表現抽出');
    function extractField(fieldName) {
      // "fieldName": "値" の形式で抽出（複数行対応）
      var re = new RegExp('"' + fieldName + '"\\s*:\\s*"([\\s\\S]*?)(?:(?<!\\\\)",|(?<!\\\\)"\\s*[\\}])', '');
      var m  = text.match(re);
      if (m) return m[1].replace(/\\"/g, '"').replace(/\\n/g, '\n').trim();
      // シンプルな1行版フォールバック
      var re2 = new RegExp('"' + fieldName + '"\\s*:\\s*"([^"]*)"');
      var m2  = text.match(re2);
      return m2 ? m2[1].trim() : '';
    }
    var r = {
      ai_summary:              extractField('ai_summary'),
      ai_advice:               extractField('ai_advice'),
      next_goal:               extractField('next_goal'),
      next_month_project_goal: extractField('next_month_project_goal'),
      next_month_study_goal:   extractField('next_month_study_goal'),
    };
    Logger.log('generateMeetAiSummary_ 抽出結果 summary=' + r.ai_summary.substring(0,50));
    // 何も取れなかった場合のみ空を返す
    return r;

  } catch(e) {
    Logger.log('generateMeetAiSummary_ error: ' + e.message);
    return { ai_summary:'', ai_advice:'', next_goal:'', next_month_project_goal:'', next_month_study_goal:'' };
  }
}

// ============================================================
// AI要約生成ヘルパー（TEL ハドル用・4項目）
// ============================================================
function generateTelAiSummary_(transcript, leaderName, menteeName) {
  if (!transcript) return { talk_content:'', concerns:'', good_points:'', memo:'', ai_summary:'', ai_mentee_status:'green' };
  try {
    var tmpl   = getPrompt_('PROMPT_TEL', DEFAULT_PROMPT_TEL);
    var prompt = tmpl
      .replace(/\{\{leader_name\}\}/g,  leaderName  || '')
      .replace(/\{\{mentee_name\}\}/g,  menteeName  || '')
      .replace(/\{\{transcript\}\}/g,   transcript.substring(0, 8000));
    var result = generateTextGemini(prompt);
    var text   = result.text || '';
    Logger.log('generateTelAiSummary_ Gemini応答: ' + text.substring(0, 200));

    // ★ コードブロック（```json ... ```）を除去
    text = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    var match = text.match(/\{[\s\S]*\}/); // ★ 貪欲マッチ（? を削除）
    if (match) {
      // ★ JSONキーのスペースをアンダースコアに正規化
      var jsonStr = match[0].replace(/"([^"]+)"\s*:/g, function(m, key) {
        return '"' + key.replace(/\s+/g, '_') + '":';
      });
      try {
        var parsed = JSON.parse(jsonStr);
        return {
          talk_content:     parsed.talk_content     || '',
          concerns:         parsed.concerns         || '',
          good_points:      parsed.good_points      || '',
          memo:             parsed.memo             || '',
          ai_summary:       parsed.ai_summary       || '',
          ai_mentee_status: parsed.ai_mentee_status || 'green',
        };
      } catch(parseErr) {
        Logger.log('generateTelAiSummary_ JSON parse error: ' + parseErr.message);
      }
    }
    return { talk_content: text.slice(0,200), concerns:'', good_points:'', memo:'', ai_summary:'', ai_mentee_status:'green' };
  } catch(e) {
    Logger.log('generateTelAiSummary_ error: ' + e.message);
    return { talk_content:'', concerns:'', good_points:'', memo:'', ai_summary:'', ai_mentee_status:'green' };
  }
}


// ============================================================
// ウォームアップ（コールドスタート対策）
// setupWarmupTrigger() をGASエディタから1回手動実行する
// ============================================================
// ============================================================
// ★ メンター全員の default_1on1_duration を60分に一括設定
// 未設定または30分のメンターのみ更新（60分以外に明示設定済みは維持）
// GASエディタから1回手動実行する
// ============================================================
function setAllMentorDurationTo60() {
  var sheet = getSheet('users');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];
  var roleIdx    = headers.indexOf('role');
  var durIdx     = headers.indexOf('default_1on1_duration');
  var nameIdx    = headers.indexOf('name');

  if (durIdx < 0) {
    Logger.log('default_1on1_duration 列が存在しません。列を追加してください。');
    return;
  }

  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    var role = String(data[i][roleIdx] || '');
    if (role !== 'mentor') continue;
    var dur  = String(data[i][durIdx] || '').trim();
    // 未設定・空・30分のメンターのみ60に更新
    if (dur === '' || dur === '30' || dur === '0') {
      data[i][durIdx] = '60';
      sheet.getRange(i + 1, durIdx + 1).setValue('60');
      Logger.log('更新: ' + data[i][nameIdx] + ' → 60分');
      updated++;
    } else {
      Logger.log('スキップ: ' + data[i][nameIdx] + ' (現在: ' + dur + '分)');
    }
  }
  invalidateCache_('users');
  Logger.log('完了: ' + updated + '名のメンターを60分に更新しました');
}

function warmup() {
  // GASインスタンスを起動状態に保つだけ（処理なし）
  // Logger.log は意図的に省略（実行ログを汚さないため）
}

function setupWarmupTrigger() {
  // 既存のwarmupトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'warmup') {
      ScriptApp.deleteTrigger(t);
      Logger.log('既存warmupトリガー削除');
    }
  });
  // 5分毎に実行
  ScriptApp.newTrigger('warmup')
    .timeBased().everyMinutes(5).create();
  Logger.log('✅ warmupトリガー設定完了（5分毎）');
}


function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // ★ 録音・録画ポーリング（10分おき・24時間365日）
  ScriptApp.newTrigger('pollRecordingsMain')
    .timeBased().everyMinutes(10).create();

  // 日次リマインド
  ScriptApp.newTrigger('dailyReminder')
    .timeBased().everyDays(1).atHour(9).create();

  // 録画・TELレポート未対応リマインド（毎日10時）
  ScriptApp.newTrigger('recordingUploadReminder')
    .timeBased().everyDays(1).atHour(10).create();

  // 週次TELリマインド（毎週金曜9時）
  ScriptApp.newTrigger('weeklyTelReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();

  // TELレポート未提出アラート（毎週月曜9時）
  ScriptApp.newTrigger('weeklyTelReportAlert')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();

  // 月次集計（毎月1日0時）
  ScriptApp.newTrigger('monthlyAggregation')
    .timeBased().onMonthDay(1).atHour(0).create();

  Logger.log('Triggers set up successfully. polling: every 10 min.');
}

// ポーリングのみ再設定（他のトリガーを消さずにポーリングだけ追加）
function setupPollingTriggerOnly() {
  // 既存のポーリングトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pollRecordingsMain') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 新規追加
  ScriptApp.newTrigger('pollRecordingsMain')
    .timeBased().everyMinutes(10).create();
  Logger.log('ポーリングトリガー設定完了（10分おき）');
}

// ============================================================
// テストユーザーセットアップ（GASエディタから手動実行）
// ※ 実行後は削除またはコメントアウト推奨
// ============================================================
function setupTestUsers() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('users');
  if (!sheet) { Logger.log('usersシートが見つかりません'); return; }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('headers: ' + JSON.stringify(headers));

  // 既存メールアドレスを取得
  var data = sheet.getDataRange().getValues();
  var emailCol = headers.indexOf('email');
  var existingEmails = data.slice(1).map(function(r){ return String(r[emailCol]||'').toLowerCase(); });

  var now = new Date().toISOString();

  var testUsers = [
    {
      user_id:       'mentor-test-001',
      email:         'test.mentor@socialshift.work',
      name:          '田中 メンター',
      role:          'mentor',
      has_leader_role: 'TRUE',
      password_hash: 'sha256:31a8536b11f3ebd553982ebd491fcadd068d2c8d28a81bfb7b39f26534a671f0',
      mentor_id:     '',
      leader_id:     '',
      phone_number:  '090-0000-0002',
      workplace:     'テストオフィス',
      work_status:   '正社員',
      hourly_wage:   '',
      status:        'active',
      created_at:    now,
      updated_at:    now,
      birthday:      '1988-03-15',
      chat_url:      '',
    },
    {
      user_id:       'mentee-test-001',
      email:         'test.mentee@socialshift.work',
      name:          'テスト メンティー',
      role:          'mentee',
      has_leader_role: 'FALSE',
      password_hash: 'sha256:cc2549c78b52e24fbd5dcb072c7a2066e8063f5908fc288e31580285d23fca61',
      mentor_id:     'mentor-test-001',
      leader_id:     'mentor-test-001',
      phone_number:  '090-0000-0003',
      workplace:     'テストオフィス',
      work_status:   'アルバイト',
      hourly_wage:   '1200',
      status:        'active',
      created_at:    now,
      updated_at:    now,
      birthday:      '2000-05-10',
      chat_url:      '',
    },
  ];

  var added = 0;
  testUsers.forEach(function(u) {
    if (existingEmails.indexOf(u.email.toLowerCase()) >= 0) {
      Logger.log('スキップ（既存）: ' + u.email);
      return;
    }
    var row = headers.map(function(h) { return u[h] !== undefined ? u[h] : ''; });
    sheet.appendRow(row);
    added++;
    Logger.log('追加: ' + u.name + ' (' + u.email + ')');
  });

  Logger.log('完了: ' + added + '件追加しました');
  Logger.log('メンター  : test.mentor@socialshift.work / Mentor12!');
  Logger.log('メンティー: test.mentee@socialshift.work / Mentee12!');
}


// ============================================================
// ログインデバッグ（GASエディタから手動実行→実行ログで確認）
// ============================================================
function debugLogin() {
  var email    = 'test.mentor@socialshift.work';
  var password = 'Mentor12!';

  var users = cachedSheetToObjects_('users');
  Logger.log('全ユーザー数: ' + users.length);

  users.forEach(function(u) {
    Logger.log('---');
    Logger.log('user_id: '       + u.user_id);
    Logger.log('email: '         + JSON.stringify(u.email));
    Logger.log('email(trim): '   + String(u.email||'').trim().toLowerCase());
    Logger.log('role: '          + u.role);
    Logger.log('status: '        + u.status);
    Logger.log('password_hash: ' + u.password_hash);
  });

  var inputHash = sha256Hash(password);
  Logger.log('=== ログイン試行 ===');
  Logger.log('email入力: '      + email);
  Logger.log('password入力: '   + password);
  Logger.log('計算ハッシュ: '   + inputHash);

  var user = users.find(function(u) {
    return String(u.email||'').trim().toLowerCase() === email.trim().toLowerCase();
  });

  if (!user) {
    Logger.log('結果: ユーザーが見つからない → スプシにメールアドレスが登録されているか確認');
  } else {
    Logger.log('ユーザー発見: ' + user.name);
    Logger.log('ハッシュ一致: ' + (user.password_hash === inputHash));
    Logger.log('スプシのhash: ' + user.password_hash);
    Logger.log('計算したhash: ' + inputHash);
  }
}


// ============================================================
// メンターからの1on1予約作成
// ============================================================
function handleMentorCreateBooking(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var mentor_id    = auth.payload.user_id;
  var mentee_id    = (body.mentee_id || '').trim();
  var scheduled_at = (body.scheduled_at || '').trim();
  var meet_link    = (body.meet_link || '').trim();
  var note         = (body.note || '').trim();

  if (!mentee_id || !scheduled_at) return errorResponse('MISSING_FIELDS', 400);

  var users  = cachedSheetToObjects_('users');
  var mentor = users.find(function(u) { return u.user_id === mentor_id; });
  var mentee = users.find(function(u) { return u.user_id === mentee_id; });

  // ★ duration_minutes: フロントから明示送信された値を優先
  // 未送信の場合はメンターの default_1on1_duration を使用
  // メンター未設定・30分以下の場合は60分（システムデフォルト）
  var duration_minutes = parseInt(body.duration_minutes) > 0
    ? parseInt(body.duration_minutes)
    : (mentor && parseInt(mentor.default_1on1_duration) > 30
        ? parseInt(mentor.default_1on1_duration)
        : 60);
  if (!mentor) return errorResponse('MENTOR_NOT_FOUND', 404);
  if (!mentee) return errorResponse('MENTEE_NOT_FOUND', 404);

  // 同じ日時に既存予約がないかチェック
  var bookings = cachedSheetToObjects_('bookings');
  var dup = bookings.find(function(b) {
    return b.status !== 'cancelled' && b.scheduled_at === scheduled_at &&
           (b.mentor_id === mentor_id || b.mentee_id === mentee_id);
  });
  if (dup) return errorResponse('DUPLICATE_BOOKING', 409);

  var booking_id = 'bk-' + Date.now() + '-' + Math.random().toString(36).substring(2, 9);
  var now = new Date().toISOString();

  appendRow('bookings', {
    booking_id:       booking_id,
    mentee_id:        mentee_id,
    mentor_id:        mentor_id,
    scheduled_at:     scheduled_at,
    duration_minutes: duration_minutes,
    status:           'scheduled',
    meet_link:        meet_link,
    recording_url:    '',
    created_at:       now,
    updated_at:       now
  });

  var dateStr = toJST_(scheduled_at); // JST変換

  var noteHtml = note ? '<li>メモ: ' + note + '</li>' : '';
  var meetHtml = meet_link ? '<li>ミーティングURL: <a href="' + meet_link + '">' + meet_link + '</a></li>' : '';

  // カレンダー登録
  var calResult2 = addCalendarEvent_({
    title:               '1on1: ' + mentee.name + ' × ' + mentor.name,
    startIso:            scheduled_at,
    durationMinutes:     duration_minutes,
    meetLink:            meet_link,
    mentorEmail:         mentor.email,
    mentorCalendarEmail: mentor.calendar_email || '',
    menteeEmail:         mentee.email,
    note:                note,
    bookingId:           booking_id,
  });
  if (calResult2.ok && calResult2.eventId) {
    var updates2 = {
      calendar_event_id: calResult2.eventId,
      calendar_url:      calResult2.htmlLink || '',  // ★ カレンダーイベント直リンク
      updated_at: new Date().toISOString()
    };
    if (calResult2.meetLink) updates2.meet_link = calResult2.meetLink;
    updateRowWhere('bookings', 'booking_id', booking_id, updates2);
  } else {
    Logger.log('カレンダー登録失敗（予約は完了）: ' + (calResult2.error || ''));
  }

  var calInfo2 = calResult2.ok ? '<li>📅 カレンダーに登録済み（招待メール送信済み）</li>' : '<li>⚠️ カレンダー登録に失敗しました</li>';

  sendMail(mentor.email,
    '【1on1予約確定】' + mentee.name + ' さんとの1on1を登録しました',
    '<h2>1on1予約確定</h2>' +
    '<p>' + mentor.name + ' さん</p>' +
    '<p>' + mentee.name + ' さんとの1on1を登録しました。</p>' +
    '<ul><li>日時: ' + dateStr + '</li><li>時間: ' + duration_minutes + '分</li>' + meetHtml + noteHtml + calInfo2 + '</ul>'
  );
  sendMail(mentee.email,
    '【1on1予約確定】' + mentor.name + ' さんとの1on1が登録されました',
    '<h2>1on1予約確定</h2>' +
    '<p>' + mentee.name + ' さん</p>' +
    '<p>' + mentor.name + ' さんとの1on1が登録されました。</p>' +
    '<ul><li>日時: ' + dateStr + '</li><li>時間: ' + duration_minutes + '分</li>' + meetHtml + noteHtml + calInfo2 + '</ul>'
  );

  return jsonResponse({ ok: true, booking_id: booking_id, calendar_ok: calResult2.ok });
}

// ============================================================
// 1on1予約変更（日時・Meet URL変更）+ 双方メール通知
// ============================================================
function handleUpdateBooking(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var booking_id   = (body.booking_id || '').trim();
  var scheduled_at = (body.scheduled_at || '').trim();
  var meet_link    = body.meet_link !== undefined ? (body.meet_link || '') : null;
  var note         = (body.note || '').trim();

  if (!booking_id) return errorResponse('MISSING_BOOKING_ID', 400);
  if (!scheduled_at) return errorResponse('MISSING_SCHEDULED_AT', 400);

  var bookings = cachedSheetToObjects_('bookings');
  var booking  = bookings.find(function(b) { return b.booking_id === booking_id; });
  if (!booking) return errorResponse('BOOKING_NOT_FOUND', 404);

  // 権限チェック：mentor または admin のみ変更可
  var caller_id = auth.payload.user_id;
  var role      = auth.payload.role;
  if (role !== 'admin' && booking.mentor_id !== caller_id) {
    return errorResponse('FORBIDDEN', 403);
  }

  var updates = { scheduled_at: scheduled_at, updated_at: new Date().toISOString() };
  if (meet_link !== null) updates.meet_link = meet_link;

  updateRowWhere('bookings', 'booking_id', booking_id, updates);

  var users  = cachedSheetToObjects_('users');
  var mentor = users.find(function(u) { return u.user_id === booking.mentor_id; });
  var mentee = users.find(function(u) { return u.user_id === booking.mentee_id; });

  var dateStr = toJST_(scheduled_at); // JST変換

  var oldDateStr = toJST_(booking.scheduled_at); // JST変換

  var newMeetLink = meet_link !== null ? meet_link : (booking.meet_link || '');
  var meetHtml    = newMeetLink ? '<li>ミーティングURL: <a href="' + newMeetLink + '">' + newMeetLink + '</a></li>' : '';
  var noteHtml    = note ? '<li>変更メモ: ' + note + '</li>' : '';

  var body_html = '<h2>1on1日程変更のお知らせ</h2>' +
    '<ul>' +
    '<li>変更前: ' + oldDateStr + '</li>' +
    '<li>変更後: ' + dateStr + '</li>' +
    meetHtml + noteHtml +
    '</ul>';

  // カレンダーイベント更新（本体）
  var calEventId = booking.calendar_event_id || '';
  var calUpResult;
  if (calEventId) {
    calUpResult = updateCalendarEvent_(calEventId, {
      startIso:        scheduled_at,
      durationMinutes: parseInt(booking.duration_minutes) || 60,
      meetLink:        newMeetLink,
      note:            note,
    });
  } else {
    // イベントIDがない場合は新規作成
    calUpResult = addCalendarEvent_({
      title:               '1on1: ' + (mentee ? mentee.name : '') + ' × ' + (mentor ? mentor.name : ''),
      startIso:            scheduled_at,
      durationMinutes:     parseInt(booking.duration_minutes) || 60,
      meetLink:            newMeetLink,
      mentorEmail:         mentor ? mentor.email : '',
      mentorCalendarEmail: mentor ? (mentor.calendar_email || '') : '',
      menteeEmail:         mentee ? mentee.email : '',
      note:                note,
      bookingId:           booking_id,
    });
    if (calUpResult.ok && calUpResult.eventId) {
      updateRowWhere('bookings', 'booking_id', booking_id, { calendar_event_id: calUpResult.eventId });
    }
  }

  // ★ 確認用カレンダーイベント更新（本体）
  var subEventId = booking.sub_calendar_event_id || '';
  if (subEventId) {
    updateSubCalendarEvent_(subEventId, {
      startIso:        scheduled_at,
      durationMinutes: parseInt(booking.duration_minutes) || 60,
    });
  } else if (mentor && mentor.calendar_email) {
    // sub_calendar_event_id がなければ新規作成
    var subNew = addSubCalendarEvent_({
      title:               '1on1: ' + (mentee ? mentee.name : '') + ' × ' + (mentor ? mentor.name : ''),
      startIso:            scheduled_at,
      durationMinutes:     parseInt(booking.duration_minutes) || 60,
      mentorEmail:         mentor.email,
      mentorCalendarEmail: mentor.calendar_email,
      bookingId:           booking_id,
    });
    if (subNew.ok && subNew.eventId) {
      updateRowWhere('bookings', 'booking_id', booking_id, { sub_calendar_event_id: subNew.eventId });
    }
  }
  var calInfo3 = (calUpResult && calUpResult.ok) ? '<li>📅 カレンダーも更新済み</li>' : '<li>⚠️ カレンダーの更新に失敗しました（手動で更新してください）</li>';

  if (mentor) sendMail(mentor.email, '【1on1日程変更】' + (mentee ? mentee.name : '') + ' さんとの1on1が変更されました', body_html + calInfo3 + '</ul>');
  if (mentee) sendMail(mentee.email, '【1on1日程変更】' + (mentor ? mentor.name : '') + ' さんとの1on1が変更されました', body_html + calInfo3 + '</ul>');

  return jsonResponse({ ok: true, calendar_ok: calUpResult ? calUpResult.ok : false });
}

// ============================================================
// メンティー: 自分のchat_urlとリーダーの電話番号を取得
// ============================================================
function handleMenteeChatUrl(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var user_id = auth.payload.user_id;

  var users  = cachedSheetToObjects_('users');
  var me     = users.find(function(u) { return u.user_id === user_id; });
  if (!me) return errorResponse('USER_NOT_FOUND', 404);

  var leader = me.leader_id ? users.find(function(u) { return u.user_id === me.leader_id; }) : null;

  return jsonResponse({
    ok:           true,
    chat_url:     me.chat_url     || '',
    leader_name:  leader ? leader.name        : '',
    leader_phone: leader ? (leader.phone_number || '') : '',
    leader_chat:  leader ? (leader.chat_url    || '') : '',
  });
}


// ============================================================
// チャットURL手動登録API
// Chat appの自動作成は非対応のため、管理者が手動でURLを登録する方式
// 手順: Google Chatでスペースを作成 → URLをコピー → 本APIで登録
// POST api/admin/create-chat-space { user_id, chat_url }
// ============================================================
function handleCreateChatSpace(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var user_id  = body.user_id  || '';
  var chat_url = body.chat_url || '';
  if (!user_id)  return errorResponse('MISSING_USER_ID', 400);
  if (!chat_url) return errorResponse('MISSING_CHAT_URL', 400);

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var target = users.find(function(u){ return u.user_id === user_id; });
  if (!target) return errorResponse('USER_NOT_FOUND', 404);

  updateRowWhere('users', 'user_id', user_id, {
    chat_url:   chat_url.trim(),
    updated_at: new Date().toISOString(),
  });
  return jsonResponse({ ok: true, chat_url: chat_url.trim() });
}



// ============================================================
// リーダー重複チェック共通関数
// 1人のメンティーに複数のリーダーは登録不可
// ============================================================
function checkDuplicateLeader_(mentee_id, exclude_mentee_id_for_update) {
  // usersシートのleader_idカラムをチェック
  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var target = users.find(function(u){ return u.user_id === mentee_id; });
  // すでに別のleader_idが設定されているかは呼び出し元で判断するのでここでは不要

  // leader_assignmentsシートもチェック（二重管理防止）
  var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a){ return a.assignment_id; });
  var existing = assignments.filter(function(a){
    return a.mentee_id === mentee_id && a.mentee_id !== exclude_mentee_id_for_update;
  });
  return existing; // 既存の割り当て一覧を返す（空なら重複なし）
}

/**
 * メンティーのleader_idを更新する際に重複をチェックして更新する
 * @param {string} mentee_id
 * @param {string} new_leader_id - '' の場合は解除
 * @param {string} current_leader_id - 既存のleader_id（更新時）
 * @returns {{ ok: boolean, error: string|null }}
 */
function setLeaderForMentee_(mentee_id, new_leader_id, current_leader_id) {
  if (!mentee_id) return { ok: false, error: 'MISSING_MENTEE_ID' };

  // leader_idを解除する場合は無条件でOK
  if (!new_leader_id) {
    return { ok: true, error: null };
  }

  // 同じリーダーへの変更はOK
  if (new_leader_id === current_leader_id) {
    return { ok: true, error: null };
  }

  // leader_assignmentsシートに同じメンティーの割り当てがないかチェック
  var assignments = cachedSheetToObjects_('leader_assignments').filter(function(a){ return a.assignment_id; });
  var dup = assignments.find(function(a){
    return a.mentee_id === mentee_id && a.leader_id !== current_leader_id;
  });
  if (dup) {
    var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    var dupLeader = users.find(function(u){ return u.user_id === dup.leader_id; });
    return {
      ok: false,
      error: 'DUPLICATE_LEADER: このメンティーにはすでに「' + (dupLeader ? dupLeader.name : dup.leader_id) + '」がリーダーとして割り当てられています。先に既存の割り当てを削除してください。'
    };
  }

  return { ok: true, error: null };
}


// ============================================================
// Mentor: 自分のスケジュール設定取得
// ============================================================
function handleGetSchedule(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var mentor_id = auth.payload.user_id;
  var schedules = cachedSheetToObjects_('mentor_schedules')
    .filter(function(s){ return s.schedule_id && s.mentor_id === mentor_id; })
    .map(function(s) {
      // ★ スプレッドシートが時刻型に変換した値を正規化して返す
      return {
        schedule_id: s.schedule_id,
        mentor_id:   s.mentor_id,
        day_of_week: s.day_of_week,
        start_time:  normalizeTimeStr_(s.start_time) || '10:00',
        end_time:    normalizeTimeStr_(s.end_time)   || '19:00',
        is_active:   s.is_active,
        created_at:  s.created_at,
      };
    });
  // default_1on1_duration を users シートから取得
  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentorUser = users.find(function(u){ return u.user_id === mentor_id; });
  var defaultDuration = (mentorUser && mentorUser.default_1on1_duration) ? String(mentorUser.default_1on1_duration) : '60';
  return jsonResponse({ ok: true, schedules: schedules, default_duration: defaultDuration });
}

// ============================================================
// Mentor: メモ取得（admin_memosシートを共用、mentor権限で自分の担当のみ）
// ============================================================
function handleGetMentorMemos(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var user_id = body.user_id || '';
  if (!user_id) return errorResponse('MISSING_USER_ID', 400);

  // 担当メンティーかチェック
  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee  = users.find(function(u){ return u.user_id === user_id; });
  if (!mentee) return errorResponse('USER_NOT_FOUND', 404);
  if (mentee.mentor_id !== auth.payload.user_id && auth.payload.role !== 'admin') {
    return errorResponse('FORBIDDEN', 403);
  }

  var sheet = getSheet('admin_memos');
  if (!sheet) return jsonResponse({ ok: true, memos: [] });
  var all = sheetToObjects(sheet).filter(function(m){ return m.memo_id; });
  var memos = all
    .filter(function(m){ return m.target_id === user_id; })
    .sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); })
    .slice(0,50)
    .map(function(m){ return { memo_id:m.memo_id, content:m.content, created_at:m.created_at }; });
  return jsonResponse({ ok: true, memos: memos });
}

// ============================================================
// Mentor: メモ保存（admin_memosシートを共用）
// ============================================================
function handleSaveMentorMemo(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var target_id = body.target_id || '';
  var content   = body.content   || '';
  if (!target_id || !content) return errorResponse('MISSING_FIELDS', 400);

  // 担当メンティーかチェック
  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === target_id; });
  if (mentee && mentee.mentor_id !== auth.payload.user_id && auth.payload.role !== 'admin') {
    return errorResponse('FORBIDDEN', 403);
  }

  var memo_id = 'memo-' + Date.now() + '-' + Math.random().toString(36).substring(2,7);
  var now     = new Date().toISOString();

  var memoSheet = getSheet('admin_memos');
  if (!memoSheet) {
    memoSheet = getSpreadsheet_().insertSheet('admin_memos');
    memoSheet.appendRow(['memo_id','admin_id','target_id','content','created_at','updated_at']);
  }
  try {
    appendRow('admin_memos', {
      memo_id:    memo_id,
      admin_id:   auth.payload.user_id,
      target_id:  target_id,
      content:    content,
      created_at: now,
      updated_at: now
    });
  } catch(err) {
    return errorResponse('SHEET_ERROR: ' + err.message, 500);
  }
  return jsonResponse({ ok: true, memo_id: memo_id });
}

// ============================================================
// Mentor: 担当メンティーのレポート・録画一覧
// ============================================================
function handleMentorMenteeReports(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var mentee_id = body.mentee_id || '';
  if (!mentee_id) return errorResponse('MISSING_MENTEE_ID', 400);

  // bookingsとcall_reportsを合わせて返す
  var bookings = cachedSheetToObjects_('bookings')
    .filter(function(b){ return b.booking_id && b.mentee_id === mentee_id && b.status === 'completed'; })
    .sort(function(a,b){ return (b.scheduled_at||'').localeCompare(a.scheduled_at||''); });

  var callReports = sheetToObjects(getSheet('call_reports'))
    .filter(function(r){ return r.report_id && r.mentee_id === mentee_id; })
    .sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); });

  return jsonResponse({ ok: true, bookings: bookings, call_reports: callReports });
}


// ============================================================
// Mentor: TEL AIサマリー生成
// POST api/mentor/tel-ai-summary { meet_url, mentee_id }
// ============================================================
function handleTelAiSummary(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  if (auth.payload.role === 'mentee') return errorResponse('FORBIDDEN', 403);
  var body = parseBody_(e);

  var meet_url  = body.meet_url  || '';
  var mentee_id = body.mentee_id || '';

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === mentee_id; });
  var menteeName = mentee ? mentee.name : 'メンティー';

  var prompt = 'あなたは1on1管理システムのAIアシスタントです。\n'
    + '以下のGoogle Meet URLのハドル録音を基に、週次TELレポートを作成してください。\n'
    + 'Meet URL: ' + meet_url + '\n'
    + '対象メンティー: ' + menteeName + '\n\n'
    + '以下のJSON形式のみで回答してください（他の文章は不要）:\n'
    + '{\n'
    + '  \"talk_content\": \"話した内容の要約（2〜3文）\",\n'
    + '  \"concerns\": \"悩みや課題があれば記載（なければ空文字）\",\n'
    + '  \"good_points\": \"良かった点・成長があれば記載（なければ空文字）\",\n'
    + '  \"memo\": \"その他特記事項（なければ空文字）\"\n'
    + '}';

  try {
    var result = generateTextGemini(prompt);
    if (!result.success) {
      return jsonResponse({ ok: false, error: 'AI生成失敗: ' + (result.error || '') });
    }
    var text  = result.text || '';
    var match = text.match(/\{[\s\S]*?\}/);
    if (!match) return jsonResponse({ ok: false, error: 'AI応答のパースに失敗しました' });
    var summary = JSON.parse(match[0]);
    return jsonResponse({ ok: true, summary: summary, meet_url: meet_url });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}



// ============================================================
// n8n F-12: 電話録音処理完了 → call_reportsに保存
// secret認証。n8nのcall_reports INSERTの代替として使う場合はこちら。
// n8nが直接Sheetsにappendする場合は不要（現状はSheetsノードで直接書いている）
// ただしステータス更新・メール通知など追加処理が必要な場合に活用
// ============================================================
function handleN8nSaveCallReport(body) {
  try {
    var mentee_id    = (body.mentee_id    || '').trim();
    var leader_id    = (body.leader_id    || '').trim();
    var meet_url     = (body.meet_url     || body.recording_url || '').trim();
    var talk_content = body.talk_content  || '';
    var concerns     = body.concerns      || '';
    var good_points  = body.good_points   || '';
    var memo         = body.memo          || '';
    var ai_summary   = body.ai_summary    || '';
    var ai_mentee_status = body.ai_mentee_status || '';
    var transcript   = (body.transcript   || '').substring(0, 5000);
    var recording_url     = body.recording_url     || meet_url;
    var recording_file_id = body.recording_file_id || '';

    if (!mentee_id) return { ok: false, error: 'MISSING_MENTEE_ID' };

    // Meet実施回数を更新
    var existing = sheetToObjects(getSheet('call_reports'))
      .filter(function(r){ return r.report_id && r.mentee_id === mentee_id && (r.meet_url || r.recording_url); });
    var meet_count = existing.length + (meet_url ? 1 : 0);

    var report_id = 'cr-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
    var now = new Date().toISOString();

    appendRow('call_reports', {
      report_id:         report_id,
      call_id:           body.call_id || '',
      leader_id:         leader_id,
      mentee_id:         mentee_id,
      transcript:        transcript,
      ai_summary:        ai_summary,
      ai_mentee_status:  ai_mentee_status,
      next_action:       body.next_action || memo,
      talk_content:      talk_content,
      concerns:          concerns,
      good_points:       good_points,
      memo:              memo,
      meet_url:          meet_url,
      meet_count:        meet_count,
      is_confirmed:      'FALSE',  // リーダーが確認・編集して確定
      recording_url:     recording_url,
      recording_file_id: recording_file_id,
      created_at:        now,
    });

    // メンティーのステータスを更新（ai_mentee_status が red/yellow の場合）
    if (ai_mentee_status === 'red' || ai_mentee_status === 'yellow') {
      updateRowWhere('users', 'user_id', mentee_id, {
        status:     ai_mentee_status === 'red' ? 'red' : 'yellow',
        updated_at: now,
      });
    }

    return { ok: true, report_id: report_id, meet_count: meet_count };
  } catch(err) {
    Logger.log('handleN8nSaveCallReport error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ============================================================
// n8n F-11: 1on1録画処理完了 → bookingsのrecording_urlを更新
// ============================================================
function handleN8nSaveMeetRecording(body) {
  try {
    var booking_id    = (body.booking_id    || '').trim();
    var recording_url = (body.recording_url || '').trim();
    var transcript    = (body.transcript    || '').substring(0, 5000);

    if (!booking_id) return { ok: false, error: 'MISSING_BOOKING_ID' };

    updateRowWhere('bookings', 'booking_id', booking_id, {
      recording_url: recording_url,
      updated_at:    new Date().toISOString(),
    });

    return { ok: true, booking_id: booking_id };
  } catch(err) {
    Logger.log('handleN8nSaveMeetRecording error: ' + err.message);
    return { ok: false, error: err.message };
  }
}


// ============================================================
// Drive連携: 録音ファイルをメンティー別フォルダに整理 + 文字起こし保存
//
// ★ 共有ドライブ対応版（Drive API使用）
// DriveAppは共有ドライブに対応していないため、Drive REST APIを直接呼ぶ
//
// 格納先: CONFIG.RECORDINGS_FOLDER_ID（共有ドライブ上）
//   └── {mentee_name}/
//         ├── yymmdd_メンティー名_1on1.mp4
//         └── yymmdd_メンティー名_1on1_文字起こし.txt
// ============================================================
function organizeRecording(body) {
  try {
    var fileId     = (body.file_id     || '').trim();
    var menteeName = (body.mentee_name || 'unknown').trim();
    var recordType = (body.record_type || body.type || 'TEL').trim();
    var transcript = body.transcript   || '';
    var memoOnly   = body.memo_only    || false; // ★ 動画なし・Geminiメモ単独フラグ

    if (!fileId) return { ok: false, error: 'MISSING_FILE_ID' };

    var token = ScriptApp.getOAuthToken();

    // ── 日付プレフィックス（yymmdd形式）──
    var nowJST = getNowJST_();
    var yy     = String(nowJST.getUTCFullYear()).slice(-2);
    var mm     = ('0' + (nowJST.getUTCMonth() + 1)).slice(-2);
    var dd     = ('0' + nowJST.getUTCDate()).slice(-2);
    var prefix = yy + mm + dd + '_' + menteeName + '_' + recordType;

    // ── メンティー個人フォルダを取得（personal_folder_id 優先） ──
    // body.personal_folder_id が渡された場合はそれを使用、なければ名前検索
    var menteeFolderId = (body.personal_folder_id || '').trim();
    if (!menteeFolderId) {
      menteeFolderId = driveApiGetOrCreateFolder_(
        CONFIG.INDIVIDUAL_FOLDER_ROOT_ID, menteeName, token
      );
      Logger.log('organizeRecording: personal_folder_id未設定のためフォルダ名検索 → ' + menteeFolderId);
    } else {
      Logger.log('organizeRecording: personal_folder_id使用 → ' + menteeFolderId);
    }
    if (!menteeFolderId) return { ok: false, error: 'MENTEE_FOLDER_CREATE_FAILED' };

    // ★ 取得したフォルダIDを users シートに自動保存（personal_folder_id が空のユーザーのみ更新）
    try {
      var users = cachedSheetToObjects_('users');
      var menteeUser = users.find(function(u) {
        return (u.name === menteeName || u.user_id === (body.mentee_id || ''))
          && u.user_id;
      });
      if (menteeUser && !menteeUser.personal_folder_id) {
        updateRowWhere('users', 'user_id', menteeUser.user_id, {
          personal_folder_id: menteeFolderId
        });
        invalidateCache_('users');
        Logger.log('organizeRecording: personal_folder_id を users に自動保存 → user_id=' + menteeUser.user_id + ' folder=' + menteeFolderId);
      }
    } catch(saveErr) {
      Logger.log('personal_folder_id 自動保存失敗（無視）: ' + saveErr.message);
    }

    // ── ファイルのメタデータ取得（拡張子用）──
    var fileMeta = driveApiGetFileMeta_(fileId, token);
    if (!fileMeta) return { ok: false, error: 'FILE_NOT_FOUND: ' + fileId };

    var origName = fileMeta.name || '';
    var ext      = '';
    var dotIdx   = origName.lastIndexOf('.');
    if (dotIdx >= 0) ext = origName.slice(dotIdx);
    var newFileName = prefix + ext;

    // ★ memo_only の場合はファイル移動・リネームをスキップ（Geminiメモは呼び出し元で移動済み）
    if (!memoOnly) {
      // ── ファイルをリネームして共有ドライブフォルダに移動 ──
      var currentParents = (fileMeta.parents || []).join(',');
      var moveUrl = 'https://www.googleapis.com/drive/v3/files/' + fileId
        + '?addParents=' + menteeFolderId
        + (currentParents ? '&removeParents=' + currentParents : '')
        + '&supportsAllDrives=true&fields=id,name,webViewLink,parents';
      var moveRes = UrlFetchApp.fetch(moveUrl, {
        method: 'PATCH',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ name: newFileName }),
        muteHttpExceptions: true
      });
      var movedFile = JSON.parse(moveRes.getContentText());
      if (!movedFile.id) {
        Logger.log('ファイル移動失敗: ' + moveRes.getContentText());
        return { ok: false, error: 'FILE_MOVE_FAILED: ' + (movedFile.error && movedFile.error.message || '') };
      }
      Logger.log('録音ファイル移動・リネーム完了: ' + newFileName);
    } else {
      Logger.log('organizeRecording: memo_only モード → ファイル移動スキップ');
    }

    // ── 文字起こしテキストファイルを共有ドライブに作成 ──
    var transcriptFileId  = '';
    var transcriptFileUrl = '';
    if (transcript) {
      var transcriptFileName = prefix + '_文字起こし.txt';
      // 既存ファイルを検索
      var searchUrl = 'https://www.googleapis.com/drive/v3/files'
        + '?q=' + encodeURIComponent(
            "name='" + transcriptFileName.replace(/'/g, "\\'") + "'"
            + " and '" + menteeFolderId + "' in parents"
            + " and trashed=false"
          )
        + '&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,webViewLink)';
      var searchRes  = JSON.parse(UrlFetchApp.fetch(searchUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }).getContentText());
      var existingFiles = searchRes.files || [];

      if (existingFiles.length > 0) {
        // 上書き更新
        var updateUrl = 'https://www.googleapis.com/upload/drive/v3/files/' + existingFiles[0].id
          + '?uploadType=media&supportsAllDrives=true';
        UrlFetchApp.fetch(updateUrl, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'text/plain' },
          payload: transcript,
          muteHttpExceptions: true
        });
        transcriptFileId  = existingFiles[0].id;
        transcriptFileUrl = existingFiles[0].webViewLink;
        Logger.log('文字起こし更新: ' + transcriptFileName);
      } else {
        // 新規作成（multipart upload）
        var boundary   = 'boundary_1on1_transcript';
        var metaJson   = JSON.stringify({ name: transcriptFileName, parents: [menteeFolderId] });
        var multipart  = '--' + boundary + '\r\n'
          + 'Content-Type: application/json; charset=UTF-8\r\n\r\n'
          + metaJson + '\r\n'
          + '--' + boundary + '\r\n'
          + 'Content-Type: text/plain; charset=UTF-8\r\n\r\n'
          + transcript + '\r\n'
          + '--' + boundary + '--';
        var createRes = JSON.parse(UrlFetchApp.fetch(
          'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink',
          {
            method: 'POST',
            headers: {
              'Authorization': 'Bearer ' + token,
              'Content-Type': 'multipart/related; boundary=' + boundary
            },
            payload: multipart,
            muteHttpExceptions: true
          }
        ).getContentText());
        transcriptFileId  = createRes.id  || '';
        transcriptFileUrl = createRes.webViewLink || '';
        Logger.log('文字起こし作成: ' + transcriptFileName);
      }
    }

    return {
      ok:                   true,
      file_id:              movedFile.id,
      file_url:             movedFile.webViewLink || '',
      file_name:            newFileName,
      folder_id:            menteeFolderId,
      folder_url:           'https://drive.google.com/drive/folders/' + menteeFolderId,
      transcript_file_id:   transcriptFileId,
      transcript_file_url:  transcriptFileUrl,
    };

  } catch(err) {
    Logger.log('organizeRecording error: ' + err.message + ' | ' + err.stack);
    return { ok: false, error: err.message };
  }
}

// ── Drive API ヘルパー: フォルダ取得または作成（共有ドライブ対応）──
function driveApiGetOrCreateFolder_(parentId, folderName, token) {
  try {
    // 検索
    var q = "name='" + folderName.replace(/'/g, "\\'") + "'"
      + " and mimeType='application/vnd.google-apps.folder'"
      + " and '" + parentId + "' in parents"
      + " and trashed=false";
    var searchRes = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files'
        + '?q=' + encodeURIComponent(q)
        + '&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id)',
      {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    ).getContentText());
    if (searchRes.files && searchRes.files.length > 0) {
      return searchRes.files[0].id;
    }
    // 新規作成
    var createRes = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          name: folderName,
          mimeType: 'application/vnd.google-apps.folder',
          parents: [parentId]
        }),
        muteHttpExceptions: true
      }
    ).getContentText());
    Logger.log('フォルダ作成: ' + folderName + ' → ' + createRes.id);
    return createRes.id || null;
  } catch(e) {
    Logger.log('driveApiGetOrCreateFolder_ error: ' + e.message);
    return null;
  }
}

// ── Drive API ヘルパー: ファイルメタデータ取得（共有ドライブ対応）──
function driveApiGetFileMeta_(fileId, token) {
  try {
    var res = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId
        + '?supportsAllDrives=true&fields=id,name,mimeType,parents,webViewLink',
      {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    ).getContentText());
    return res.id ? res : null;
  } catch(e) {
    Logger.log('driveApiGetFileMeta_ error: ' + e.message);
    return null;
  }
}

// ============================================================
// Drive連携: 音声・動画ファイルを文字起こし
// n8n F-11/F-12 から呼ばれる（secret認証）
// input:  { file_id }
// output: { ok, content, text, method }
// ============================================================
function getFileContent(body) {
  try {
    var fileId = body.file_id || '';
    if (!fileId) return { ok: false, error: 'MISSING_FILE_ID' };

    var file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch(e) {
      return { ok: false, error: 'FILE_NOT_FOUND: ' + e.message };
    }

    var mimeType = file.getMimeType();
    var fileName = file.getName().toLowerCase();
    Logger.log('getFileContent: ' + fileName + ' (' + mimeType + ')');

    // ── 方法1: Googleドキュメント（文字起こし済みテキスト）の場合そのまま返す ──
    if (mimeType === 'application/vnd.google-apps.document') {
      var doc = DocumentApp.openById(fileId);
      var text = doc.getBody().getText();
      return { ok: true, content: text, text: text, method: 'google_doc' };
    }

    // ── 方法2: 音声・動画ファイルをGoogleドキュメントに変換して文字起こし ──
    // Google Drive APIのimportData経由でOCR/音声認識
    var isAudio = mimeType.indexOf('audio') >= 0 || mimeType.indexOf('video') >= 0
      || fileName.endsWith('.m4a') || fileName.endsWith('.mp3')
      || fileName.endsWith('.mp4') || fileName.endsWith('.webm');

    if (isAudio) {
      // 音声ファイルをGoogleドキュメントとしてインポート（自動文字起こし）
      try {
        var token     = ScriptApp.getOAuthToken();
        var fileBytes = file.getBlob();
        var metadata  = {
          name:     file.getName(),
          mimeType: 'application/vnd.google-apps.document',
          parents:  []
        };
        var formData  = {
          metadata: Utilities.newBlob(
            JSON.stringify(metadata), 'application/json'
          ),
          file: fileBytes
        };
        var res = UrlFetchApp.fetch(
          'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
          {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + token },
            payload: formData,
            muteHttpExceptions: true
          }
        );
        var result = JSON.parse(res.getContentText());
        if (result.id) {
          // 変換されたGoogleドキュメントを開く
          Utilities.sleep(3000); // 変換待機
          try {
            var doc2 = DocumentApp.openById(result.id);
            var text2 = doc2.getBody().getText();
            // 一時ファイルを削除
            try { DriveApp.getFileById(result.id).setTrashed(true); } catch(e) {}
            Logger.log('文字起こし成功: ' + text2.length + '文字');
            return { ok: true, content: text2, text: text2, method: 'drive_import' };
          } catch(docErr) {
            Logger.log('ドキュメント読み込みエラー: ' + docErr.message);
          }
        } else {
          Logger.log('Drive変換失敗: ' + JSON.stringify(result));
        }
      } catch(importErr) {
        Logger.log('音声インポートエラー: ' + importErr.message);
      }
    }

    // ── 方法3: テキストファイルの場合そのまま返す ──
    if (mimeType === 'text/plain' || fileName.endsWith('.txt') || fileName.endsWith('.vtt')) {
      var text3 = file.getBlob().getDataAsString('UTF-8');
      return { ok: true, content: text3, text: text3, method: 'text_file' };
    }

    // ★ Google Docs形式（.txtという名前でもGoogle Docsとして保存される場合がある）
    if (mimeType === 'application/vnd.google-apps.document') {
      try {
        var token3 = ScriptApp.getOAuthToken();
        var exportUrl = 'https://www.googleapis.com/drive/v3/files/' + fileId
          + '/export?mimeType=text%2Fplain';
        var exportRes = UrlFetchApp.fetch(exportUrl, {
          headers: { 'Authorization': 'Bearer ' + token3 },
          muteHttpExceptions: true
        });
        if (exportRes.getResponseCode() === 200) {
          var docText = exportRes.getContentText('UTF-8');
          Logger.log('Google Docs export成功: ' + docText.length + '文字');
          return { ok: true, content: docText, text: docText, method: 'gdocs_export' };
        }
        Logger.log('Google Docs export失敗: ' + exportRes.getResponseCode());
      } catch(exportErr) {
        Logger.log('Google Docs export エラー: ' + exportErr.message);
      }
    }

    // 文字起こし不可の場合は空を返す（n8n側でスキップ処理）
    return { ok: true, content: '', text: '', method: 'unsupported', mime_type: mimeType };

  } catch(err) {
    Logger.log('getFileContent error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ============================================================
// TELレポート確定（is_confirmed を TRUE に更新）
// POST api/mentor/call-reports/confirm { report_id }
// ============================================================
function handleConfirmCallReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  if (auth.payload.role === 'mentee') return errorResponse('FORBIDDEN', 403);
  var body = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  // call_reports の全件取得して対象を探す
  var all = sheetToObjects(getSheet('call_reports')).filter(function(r){ return r.report_id; });
  var target = all.find(function(r){ return r.report_id === report_id; });
  if (!target) return errorResponse('REPORT_NOT_FOUND', 404);

  // updateRowWhere で is_confirmed を TRUE に更新（列の存在チェック不要）
  try {
    updateRowWhere('call_reports', 'report_id', report_id, {
      is_confirmed: 'TRUE',
      updated_at:   new Date().toISOString()
    });
    invalidateCache_('call_reports');
    return jsonResponse({ ok: true });
  } catch(err) {
    Logger.log('handleConfirmCallReport error: ' + err.message);
    return errorResponse('UPDATE_FAILED: ' + err.message, 500);
  }
}


// ============================================================
// call_reports シートに不足カラムを自動追加
// GASエディタから一度だけ手動実行
// ============================================================
function setupCallReportsColumns() {
  var sheet = getSheet('call_reports');
  if (!sheet) {
    Logger.log('call_reports シートが見つかりません');
    return;
  }
  var required = [
    'report_id','call_id','leader_id','mentee_id',
    'transcript','ai_summary','ai_mentee_status','next_action',
    'talk_content','concerns','good_points','memo',
    'meet_url','meet_count','is_confirmed',
    'recording_url','recording_file_id','created_at','updated_at'
  ];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var added = [];
  required.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      headers.push(col);
      added.push(col + '（列' + nextCol + '）');
    }
  });
  if (added.length === 0) {
    Logger.log('call_reports: 全カラム揃っています');
  } else {
    Logger.log('追加したカラム: ' + added.join(', '));
  }
}


// ============================================================
// Admin: 事前レポート一覧（全件・新しい順・ページング対応）
// GET api/admin/pre-reports?page=1&per=20&mentee_id=xxx
// ============================================================
function handleAdminPreReports(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var page    = parseInt(body.page || e.parameter.page || 1);
  var per     = Math.min(parseInt(body.per  || e.parameter.per  || 20), 50);
  var menteeId= body.mentee_id || e.parameter.mentee_id || '';

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {}; users.forEach(function(u){ userMap[u.user_id] = u; });

  var all = sheetToObjects(getSheet('pre_reports')).filter(function(r){ return r.report_id; });
  if (menteeId) all = all.filter(function(r){ return r.user_id === menteeId; });
  all.sort(function(a,b){ return (b.submitted_at||'').localeCompare(a.submitted_at||''); });

  var total  = all.length;
  var paged  = all.slice((page-1)*per, page*per).map(function(r){
    var u = userMap[r.user_id] || {};
    var m = userMap[(userMap[r.user_id]||{}).mentor_id] || {};
    return Object.assign({}, r, {
      mentee_name: u.name || r.user_id,
      mentor_name: m.name || ''
    });
  });
  return jsonResponse({ ok:true, total:total, page:page, per:per, reports:paged });
}

// ============================================================
// Admin: 1on1レポート一覧（mentor_reports・全件）
// GET api/admin/mentor-reports?page=1&per=20&mentee_id=xxx&mentor_id=xxx
// ============================================================
function handleAdminMentorReports(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var page     = parseInt(body.page     || e.parameter.page     || 1);
  var per      = Math.min(parseInt(body.per || e.parameter.per  || 20), 50);
  var menteeId = body.mentee_id  || e.parameter.mentee_id  || '';
  var mentorId = body.mentor_id  || e.parameter.mentor_id  || '';

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {}; users.forEach(function(u){ userMap[u.user_id] = u; });

  var all = sheetToObjects(getSheet('mentor_reports')).filter(function(r){ return r.report_id; });
  if (menteeId) all = all.filter(function(r){ return r.mentee_id === menteeId; });
  if (mentorId) all = all.filter(function(r){ return r.mentor_id === mentorId; });
  all.sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); });

  var total = all.length;
  var paged = all.slice((page-1)*per, page*per).map(function(r){
    return Object.assign({}, r, {
      mentee_name: (userMap[r.mentee_id]||{}).name || r.mentee_id,
      mentor_name: (userMap[r.mentor_id]||{}).name || r.mentor_id
    });
  });
  return jsonResponse({ ok:true, total:total, page:page, per:per, reports:paged });
}

// ============================================================
// Admin: TELレポート一覧（call_reports・全件）
// GET api/admin/call-reports?page=1&per=20&mentee_id=xxx&leader_id=xxx
// ============================================================
function handleAdminCallReports(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var page     = parseInt(body.page     || e.parameter.page     || 1);
  var per      = Math.min(parseInt(body.per || e.parameter.per  || 20), 50);
  var menteeId = body.mentee_id  || e.parameter.mentee_id  || '';
  var leaderId = body.leader_id  || e.parameter.leader_id  || '';
  var confirmed= body.is_confirmed || e.parameter.is_confirmed || '';

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {}; users.forEach(function(u){ userMap[u.user_id] = u; });

  var all = sheetToObjects(getSheet('call_reports')).filter(function(r){ return r.report_id; });
  if (menteeId)  all = all.filter(function(r){ return r.mentee_id  === menteeId; });
  if (leaderId)  all = all.filter(function(r){ return r.leader_id  === leaderId; });
  if (confirmed === 'TRUE')  all = all.filter(function(r){ return String(r.is_confirmed||'').toUpperCase() === 'TRUE'; });
  if (confirmed === 'FALSE') all = all.filter(function(r){ return r.is_confirmed !== 'TRUE'; });
  all.sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); });

  var total = all.length;
  var paged = all.slice((page-1)*per, page*per).map(function(r){
    return Object.assign({}, r, {
      mentee_name: (userMap[r.mentee_id]||{}).name || r.mentee_id,
      leader_name: (userMap[r.leader_id]||{}).name || r.leader_id
    });
  });
  return jsonResponse({ ok:true, total:total, page:page, per:per, reports:paged });
}

// ============================================================
// Mentor: 自分が担当するメンティーの mentor_reports 一覧
// GET api/mentor/mentor-reports?page=1&mentee_id=xxx
// ============================================================
function handleMentorReportsList(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var mentor_id = auth.payload.user_id;
  var page      = parseInt(body.page || 1);
  var per       = 20;
  var menteeId  = body.mentee_id || '';

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {}; users.forEach(function(u){ userMap[u.user_id] = u; });

  var all = sheetToObjects(getSheet('mentor_reports')).filter(function(r){
    return r.report_id && r.mentor_id === mentor_id;
  });
  if (menteeId) all = all.filter(function(r){ return r.mentee_id === menteeId; });
  all.sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); });

  var total = all.length;
  var paged = all.slice((page-1)*per, page*per).map(function(r){
    return Object.assign({}, r, {
      mentee_name: (userMap[r.mentee_id]||{}).name || r.mentee_id
    });
  });
  return jsonResponse({ ok:true, total:total, page:page, per:per, reports:paged });
}

// ============================================================
// Mentor: 1on1レポートのAI要約を再生成
// POST api/mentor/mentor-reports/regenerate-ai { report_id }
// ① transcript_file_url → transcript → フォルダ検索 の順で文字起こしを取得
// ② generateMeetAiSummary_ を実行
// ③ mentor_reports の該当行を更新して結果を返す
// ============================================================
function handleRegenerateAi(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  // ── レポート取得 ──
  var reports = sheetToObjects(getSheet('mentor_reports')).filter(function(r){ return r.report_id; });
  var report  = reports.find(function(r){ return r.report_id === report_id; });
  if (!report) return errorResponse('REPORT_NOT_FOUND', 404);

  // ── ユーザー情報取得（personal_folder_id は最新値が必要なためキャッシュなし）──
  var users   = sheetToObjects(getSheet('users')).filter(function(u){ return u.user_id; });
  var mentee  = users.find(function(u){ return u.user_id === report.mentee_id; });
  // ★ mentor_id が空または見つからない場合はログイン中のメンターを使用
  var mentor  = users.find(function(u){ return u.user_id === report.mentor_id; });
  if (!mentor) mentor = users.find(function(u){ return u.user_id === auth.payload.user_id; });
  if (!mentee) return errorResponse('MENTEE_NOT_FOUND', 404);
  if (!mentor) return errorResponse('MENTOR_NOT_FOUND', 404);
  Logger.log('handleRegenerateAi: mentor=' + mentor.name + ' mentee=' + mentee.name + ' report_mentor_id=' + report.mentor_id);

  // ── 予約日時を bookings から取得（日付絞り込みの基準） ──
  var scheduledAt = '';
  if (report.booking_id) {
    var bookings = cachedSheetToObjects_('bookings');
    var booking  = bookings.find(function(b){ return b.booking_id === report.booking_id; });
    if (booking && booking.scheduled_at) scheduledAt = booking.scheduled_at;
  }
  // bookingがない場合は report.created_at を使う
  if (!scheduledAt) scheduledAt = report.created_at || '';

  var transcript = '';
  var source     = '';

  // ── 優先1: transcript_file_url からファイルIDを取得して読み込み ──
  // このURLはレポート作成時に紐づけられた特定ファイルなので最も安全
  if (!transcript && report.transcript_file_url) {
    try {
      var fileIdMatch = report.transcript_file_url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!fileIdMatch) fileIdMatch = report.transcript_file_url.match(/id=([a-zA-Z0-9_-]+)/);
      if (fileIdMatch) {
        var txtResult = getFileContent({ file_id: fileIdMatch[1] });
        if (txtResult.ok && txtResult.text) {
          transcript = txtResult.text;
          source = 'transcript_file_url';
          Logger.log('handleRegenerateAi: 文字起こしファイルから取得 ' + transcript.length + '文字');
        }
      }
    } catch(e1) { Logger.log('transcript_file_url読み込み失敗: ' + e1.message); }
  }

  // ── 優先2: transcriptカラム（スプシ内保存済みテキスト）──
  if (!transcript && report.transcript) {
    transcript = report.transcript;
    source = 'transcript_column';
    Logger.log('handleRegenerateAi: transcriptカラムから取得 ' + transcript.length + '文字');
  }

  // ── 優先3: メンティーフォルダ内の 1on1文字起こし.txt を日付で厳密検索 ──
  if (!transcript) {
    try {
      var token = ScriptApp.getOAuthToken();
      // ★ users.personal_folder_id を優先使用。なければフォルダ名検索
      var menteeFolderId = (mentee.personal_folder_id || '').trim();
      if (!menteeFolderId) {
        menteeFolderId = driveApiGetOrCreateFolder_(CONFIG.INDIVIDUAL_FOLDER_ROOT_ID, mentee.name, token);
        Logger.log('handleRegenerateAi: personal_folder_id未設定のためフォルダ名検索 → ' + menteeFolderId);
        // ★ 見つかったフォルダIDを users シートに自動保存
        if (menteeFolderId) {
          try {
            updateRowWhere('users', 'user_id', mentee.user_id, { personal_folder_id: menteeFolderId });
            invalidateCache_('users');
            Logger.log('handleRegenerateAi: personal_folder_id を自動保存 → ' + menteeFolderId);
          } catch(se) { Logger.log('personal_folder_id 保存失敗: ' + se.message); }
        }
      } else {
        Logger.log('handleRegenerateAi: personal_folder_id使用 → ' + menteeFolderId);
      }
      if (menteeFolderId) {
        // ★ scheduled_at（予約日時）から年月日を取得（JST基準）
        var yy = '', mm = '', dd = '';
        if (scheduledAt) {
          var d   = new Date(scheduledAt);
          var jst = new Date(d.getTime() + 9 * 60 * 60 * 1000); // UTC→JST
          yy = String(jst.getUTCFullYear()).slice(-2);
          mm = ('0' + (jst.getUTCMonth() + 1)).slice(-2);
          dd = ('0' + jst.getUTCDate()).slice(-2);
        }

        // ★ 1on1専用キーワード（TELは絶対に参照しない）
        // mimeType フィルタなし（.txt でもGoogle Docs形式で保存される場合があるため）
        var q = "'" + menteeFolderId + "' in parents"
          + " and name contains '_1on1_文字起こし'"
          + " and trashed=false";
        // ★ yymmdd（年月日）で絞り込む。ヒットしない場合は±1日でも試みる
        if (yy && mm && dd) q += " and name contains '" + yy + mm + dd + "'";
        else if (yy && mm)  q += " and name contains '" + yy + mm + "'";

        var fetchFiles_ = function(query) {
          var res = JSON.parse(UrlFetchApp.fetch(
            'https://www.googleapis.com/drive/v3/files'
              + '?q=' + encodeURIComponent(query)
              + '&supportsAllDrives=true&includeItemsFromAllDrives=true'
              + '&orderBy=createdTime+desc&pageSize=5'
              + '&fields=files(id,name,createdTime)',
            { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
          ).getContentText());
          return res.files || [];
        };

        var txtFiles = fetchFiles_(q);
        Logger.log('handleRegenerateAi: フォルダ検索結果 ' + txtFiles.length + '件 (mentee=' + mentee.name + ', date=' + yy+mm+dd + ')');

        // ★ ヒットしなかった場合、±1日で再検索
        if (txtFiles.length === 0 && yy && mm && dd && scheduledAt) {
          Logger.log('handleRegenerateAi: 日付一致なし → ±1日で再検索');
          var baseMs = new Date(scheduledAt).getTime();
          var jstBase = new Date(baseMs + 9*60*60*1000);
          [-1, 1].forEach(function(diff) {
            if (txtFiles.length > 0) return;
            var altDate = new Date(jstBase.getTime() + diff * 24*60*60*1000);
            var ayy = String(altDate.getUTCFullYear()).slice(-2);
            var amm = ('0' + (altDate.getUTCMonth()+1)).slice(-2);
            var add = ('0' + altDate.getUTCDate()).slice(-2);
            var altQ = "'" + menteeFolderId + "' in parents"
              + " and name contains '_1on1_文字起こし'"
              + " and name contains '" + ayy + amm + add + "'"
              + " and trashed=false";
            txtFiles = fetchFiles_(altQ);
            Logger.log('handleRegenerateAi: ±' + diff + '日検索(' + ayy+amm+add + ')結果: ' + txtFiles.length + '件');
          });
        }

        // yymm のみのフォールバック
        if (txtFiles.length === 0 && yy && mm) {
          var qFallback = "'" + menteeFolderId + "' in parents"
            + " and name contains '_1on1_文字起こし'"
            + " and name contains '" + yy + mm + "'"
            + " and trashed=false";
          txtFiles = fetchFiles_(qFallback);
          Logger.log('handleRegenerateAi: yymm フォールバック結果: ' + txtFiles.length + '件');
        }

        if (txtFiles.length > 0) {
          // ★ 予約日時に最も近いファイルを選択（同日内での精度向上）
          var bestFile = txtFiles[0];
          if (scheduledAt && txtFiles.length > 1) {
            var reportMs = new Date(scheduledAt).getTime();
            var bestDiff = Infinity;
            txtFiles.forEach(function(f) {
              var fMs = new Date(f.createdTime).getTime();
              var diff = Math.abs(reportMs - fMs);
              if (diff < bestDiff) { bestDiff = diff; bestFile = f; }
            });
          }
          Logger.log('handleRegenerateAi: 使用ファイル=' + bestFile.name + ' (mentee=' + mentee.name + ')');
          var txtResult2 = getFileContent({ file_id: bestFile.id });
          if (txtResult2.ok && txtResult2.text) {
            transcript = txtResult2.text;
            source = 'folder_search:' + bestFile.name;
            Logger.log('handleRegenerateAi: フォルダから取得 ' + transcript.length + '文字');
          }
        }
      }
    } catch(e3) { Logger.log('フォルダ検索失敗: ' + e3.message); }
  }

  if (!transcript) {
    return jsonResponse({
      ok: false,
      no_transcript: true,
      message: '文字起こしが見つかりませんでした（対象: ' + mentee.name + ', 日時: ' + scheduledAt + '）。Google Meet の文字起こし機能が有効になっているか確認してください。'
    });
  }

  Logger.log('handleRegenerateAi: AI生成開始 source=' + source + ' mentee=' + mentee.name);

  // ── AI要約生成（1on1専用プロンプト） ──
  var aiResult = generateMeetAiSummary_(transcript, mentee.name, mentor.name);
  Logger.log('handleRegenerateAi: AI生成完了 summary=' + (aiResult.ai_summary || '[空]').slice(0, 100)
    + ' advice=' + (aiResult.ai_advice || '[空]').slice(0, 50));

  // ── mentor_reports 更新 ──
  updateRowWhere('mentor_reports', 'report_id', report_id, {
    ai_summary:              aiResult.ai_summary              || '',
    ai_advice:               aiResult.ai_advice               || '',
    next_goal:               aiResult.next_goal               || '',
    next_month_project_goal: aiResult.next_month_project_goal || '',
    next_month_study_goal:   aiResult.next_month_study_goal   || '',
    updated_at:              new Date().toISOString(),
  });
  invalidateCache_('mentor_reports');

  return jsonResponse({
    ok:      true,
    source:  source,
    transcript_length: transcript.length,
    ai_summary:              aiResult.ai_summary              || '',
    ai_advice:               aiResult.ai_advice               || '',
    next_goal:               aiResult.next_goal               || '',
    next_month_project_goal: aiResult.next_month_project_goal || '',
    next_month_study_goal:   aiResult.next_month_study_goal   || '',
  });
}

// ============================================================
// Mentor: 1on1レポート保存（編集）
// POST api/mentor/mentor-reports/save
// ============================================================
function handleSaveMentorReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  var updates = {
    ai_summary:    body.ai_summary    !== undefined ? String(body.ai_summary  || '') : undefined,
    ai_advice:     body.ai_advice     !== undefined ? String(body.ai_advice   || '') : undefined,
    next_goal:     body.next_goal     !== undefined ? String(body.next_goal   || '') : undefined,
    mentor_edited: 'TRUE',
    updated_at:    new Date().toISOString()
  };
  // undefinedのキーを除去
  Object.keys(updates).forEach(function(k){ if (updates[k] === undefined) delete updates[k]; });
  if (body.next_month_project_goal !== undefined) updates.next_month_project_goal = String(body.next_month_project_goal || '');
  if (body.next_month_study_goal   !== undefined) updates.next_month_study_goal   = String(body.next_month_study_goal   || '');

  Logger.log('handleSaveMentorReport: report_id=' + report_id + ' keys=' + Object.keys(updates).join(','));
  try {
    updateRowWhere('mentor_reports', 'report_id', report_id, updates);
  } catch(err) {
    Logger.log('handleSaveMentorReport エラー: ' + err.message);
    return errorResponse('SAVE_FAILED: ' + err.message, 500);
  }
  invalidateCache_('mentor_reports');
  return jsonResponse({ ok: true });
}

// ============================================================
// Mentor: 1on1レポート公開（メンティーに開示）
// POST api/mentor/mentor-reports/publish
// ============================================================
function handlePublishMentorReport(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  var now = new Date().toISOString();

  // ★ 公開前に編集内容も同時保存（保存を押し忘れても反映される）
  var saveUpdates = { mentor_edited: 'TRUE', updated_at: now };
  if (body.ai_summary    !== undefined) saveUpdates.ai_summary    = body.ai_summary    || '';
  if (body.ai_advice     !== undefined) saveUpdates.ai_advice     = body.ai_advice     || '';
  if (body.next_goal     !== undefined) saveUpdates.next_goal     = body.next_goal     || '';
  if (body.next_month_project_goal !== undefined) saveUpdates.next_month_project_goal = body.next_month_project_goal || '';
  if (body.next_month_study_goal   !== undefined) saveUpdates.next_month_study_goal   = body.next_month_study_goal   || '';
  if (Object.keys(saveUpdates).length > 2) {
    updateRowWhere('mentor_reports', 'report_id', report_id, saveUpdates);
  }

  // PDF 生成（publish 前にレポートデータを取得）
  var allReports = sheetToObjects(getSheet('mentor_reports'));
  var report     = allReports.find(function(r){ return r.report_id === report_id; });
  var pdfUrl     = '';
  if (report) {
    var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    var mentee  = users.find(function(u){ return u.user_id === report.mentee_id; });
    var mentor  = users.find(function(u){ return u.user_id === report.mentor_id; });
    var menteeName = mentee ? mentee.name : (report.mentee_id || 'unknown');
    var mentorName = mentor ? mentor.name : '';
    pdfUrl = generateMentorReportPdf_(report, menteeName, mentorName);
  }

  updateRowWhere('mentor_reports', 'report_id', report_id, {
    is_published: 'TRUE',
    published_at: now,
    updated_at:   now,
    pdf_url:      pdfUrl
  });
  invalidateCache_('mentor_reports');

  // メンティーにメール通知（最新データで再取得）
  var all    = sheetToObjects(getSheet('mentor_reports'));
  report     = all.find(function(r){ return r.report_id === report_id; });
  if (report) {
    var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
    var mentee = users.find(function(u){ return u.user_id === report.mentee_id; });
    var mentor = users.find(function(u){ return u.user_id === report.mentor_id; });
    if (mentee) {
      sendMail(mentee.email,
        '【1on1レポート公開】' + (mentor ? mentor.name : 'メンター') + ' さんからのレポートが届きました',
        '<h2>1on1レポートが公開されました</h2>'
        + '<p>' + mentee.name + ' さん</p>'
        + '<p>メンターからあなたの1on1レポートが公開されました。</p><hr>'
        + '<h3>AIサマリー</h3><p>' + (report.ai_summary || '—') + '</p>'
        + '<h3>アドバイス</h3><p>' + (report.ai_advice || '—') + '</p>'
        + '<h3>次回の目標</h3><p>' + (report.next_goal || '—') + '</p>'
        + (report.next_month_project_goal ? '<h3>次月プロジェクト目標</h3><p>' + report.next_month_project_goal + '</p>' : '')
        + (report.next_month_study_goal   ? '<h3>次月学習目標</h3><p>'         + report.next_month_study_goal   + '</p>' : '')
      );
    }
  }
  return jsonResponse({ ok: true });
}

// ============================================================
// Mentor: 1on1レポート下書きに戻す（公開取り消し）
// POST api/mentor/mentor-reports/unpublish
// ============================================================
function handleUnpublishMentorReport(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  // レポート存在チェック（mentorロール認証済みのため所有者チェックは不要）
  var reports = cachedSheetToObjects_('mentor_reports');
  var report  = reports.find(function(r){ return r.report_id === report_id; });
  if (!report) return errorResponse('REPORT_NOT_FOUND', 404);

  updateRowWhere('mentor_reports', 'report_id', report_id, {
    is_published: 'FALSE',
    published_at: '',
    updated_at:   new Date().toISOString()
  });
  invalidateCache_('mentor_reports');
  Logger.log('handleUnpublishMentorReport: ' + report_id + ' → 下書きに戻す by ' + auth.payload.user_id);
  return jsonResponse({ ok: true });
}

// ============================================================
// Leader: TEL週次レポート AI再生成
// POST api/leader/call-reports/regenerate-ai { report_id }
// ============================================================
function handleRegenerateTelAi(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var report_id = (body.report_id || '').trim();
  if (!report_id) return errorResponse('MISSING_REPORT_ID', 400);

  // ── レポート取得（call_reports） ──
  var reports = sheetToObjects(getSheet('call_reports')).filter(function(r){ return r.report_id; });
  var report  = reports.find(function(r){ return r.report_id === report_id; });
  if (!report) return errorResponse('REPORT_NOT_FOUND', 404);

  // ── ユーザー情報取得（personal_folder_id は最新値が必要なためキャッシュなし）──
  var users  = sheetToObjects(getSheet('users')).filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === report.mentee_id; });
  var leader = users.find(function(u){ return u.user_id === report.leader_id; });
  if (!mentee) return errorResponse('MENTEE_NOT_FOUND', 404);
  if (!leader) return errorResponse('LEADER_NOT_FOUND', 404);

  // ★ 通話日時（created_at）を基準に日付絞り込み
  var calledAt = report.created_at || '';

  var transcript = '';
  var source     = '';

  // ── 優先1: transcript_file_url から直接取得 ──
  if (!transcript && report.transcript_file_url) {
    try {
      var fileIdMatch = report.transcript_file_url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!fileIdMatch) fileIdMatch = report.transcript_file_url.match(/id=([a-zA-Z0-9_-]+)/);
      if (fileIdMatch) {
        var txtResult = getFileContent({ file_id: fileIdMatch[1] });
        if (txtResult.ok && txtResult.text) {
          transcript = txtResult.text;
          source = 'transcript_file_url';
          Logger.log('handleRegenerateTelAi: ファイルから取得 ' + transcript.length + '文字');
        }
      }
    } catch(e1) { Logger.log('transcript_file_url読み込み失敗: ' + e1.message); }
  }

  // ── 優先2: transcriptカラム ──
  if (!transcript && report.transcript) {
    transcript = report.transcript;
    source = 'transcript_column';
    Logger.log('handleRegenerateTelAi: transcriptカラムから取得 ' + transcript.length + '文字');
  }

  // ── 優先3: メンティーフォルダ内の TEL文字起こし.txt を日付で厳密検索 ──
  if (!transcript) {
    try {
      var token = ScriptApp.getOAuthToken();
      // ★ users.personal_folder_id を優先使用。なければフォルダ名検索
      var menteeFolderId = (mentee.personal_folder_id || '').trim();
      if (!menteeFolderId) {
        menteeFolderId = driveApiGetOrCreateFolder_(CONFIG.INDIVIDUAL_FOLDER_ROOT_ID, mentee.name, token);
        Logger.log('handleRegenerateTelAi: personal_folder_id未設定のためフォルダ名検索 → ' + menteeFolderId);
        // ★ 見つかったフォルダIDを users シートに自動保存
        if (menteeFolderId) {
          try {
            updateRowWhere('users', 'user_id', mentee.user_id, { personal_folder_id: menteeFolderId });
            invalidateCache_('users');
            Logger.log('handleRegenerateTelAi: personal_folder_id を自動保存 → ' + menteeFolderId);
          } catch(se) { Logger.log('personal_folder_id 保存失敗: ' + se.message); }
        }
      } else {
        Logger.log('handleRegenerateTelAi: personal_folder_id使用 → ' + menteeFolderId);
      }
      if (menteeFolderId) {
        var yy = '', mm = '', dd = '';
        if (calledAt) {
          var dt = new Date(calledAt);
          yy = String(dt.getFullYear()).slice(-2);
          mm = ('0' + (dt.getMonth() + 1)).slice(-2);
          dd = ('0' + dt.getDate()).slice(-2);
        }

        // ★ TEL専用キーワード（1on1は絶対に参照しない）
        // mimeType フィルタなし（.txt でもGoogle Docs形式で保存される場合があるため）
        var q = "'" + menteeFolderId + "' in parents"
          + " and name contains '_TEL_文字起こし'"
          + " and trashed=false";
        // ★ yymmdd（年月日）で絞り込む
        if (yy && mm && dd) q += " and name contains '" + yy + mm + dd + "'";
        else if (yy && mm)  q += " and name contains '" + yy + mm + "'";

        var searchRes = JSON.parse(UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files'
            + '?q=' + encodeURIComponent(q)
            + '&supportsAllDrives=true&includeItemsFromAllDrives=true'
            + '&orderBy=createdTime+desc&pageSize=5'
            + '&fields=files(id,name,createdTime)',
          { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
        ).getContentText());

        var txtFiles = (searchRes.files || []);
        Logger.log('handleRegenerateTelAi: フォルダ検索結果 ' + txtFiles.length + '件 (mentee=' + mentee.name + ', date=' + yy+mm+dd + ')');

        if (txtFiles.length > 0) {
          // ★ 通話日時に最も近いファイルを選択
          var bestFile = txtFiles[0];
          if (calledAt && txtFiles.length > 1) {
            var callMs = new Date(calledAt).getTime();
            var bestDiff = Infinity;
            txtFiles.forEach(function(f) {
              var fMs = new Date(f.createdTime).getTime();
              var diff = Math.abs(callMs - fMs);
              if (diff < bestDiff) { bestDiff = diff; bestFile = f; }
            });
          }
          Logger.log('handleRegenerateTelAi: 使用ファイル=' + bestFile.name + ' (mentee=' + mentee.name + ')');
          var txtResult2 = getFileContent({ file_id: bestFile.id });
          if (txtResult2.ok && txtResult2.text) {
            transcript = txtResult2.text;
            source = 'folder_search:' + bestFile.name;
          }
        }
      }
    } catch(e3) { Logger.log('フォルダ検索失敗: ' + e3.message); }
  }

  if (!transcript) {
    return jsonResponse({
      ok: false,
      no_transcript: true,
      message: '文字起こしが見つかりませんでした（対象: ' + mentee.name + ', 日時: ' + calledAt + '）。'
    });
  }

  Logger.log('handleRegenerateTelAi: AI生成開始 source=' + source + ' mentee=' + mentee.name);

  // ★ TEL専用プロンプトで生成（1on1プロンプトは使わない）
  var aiResult = generateTelAiSummary_(transcript, leader.name, mentee.name);
  Logger.log('handleRegenerateTelAi: AI生成完了');

  updateRowWhere('call_reports', 'report_id', report_id, {
    ai_summary:       aiResult.ai_summary       || '',
    talk_content:     aiResult.talk_content      || '',
    concerns:         aiResult.concerns          || '',
    good_points:      aiResult.good_points       || '',
    memo:             aiResult.memo              || '',
    ai_mentee_status: aiResult.ai_mentee_status  || 'green',
    updated_at:       new Date().toISOString(),
  });
  invalidateCache_('call_reports');

  return jsonResponse({
    ok:               true,
    source:           source,
    transcript_length: transcript.length,
    ai_summary:       aiResult.ai_summary       || '',
    talk_content:     aiResult.talk_content      || '',
    concerns:         aiResult.concerns          || '',
    good_points:      aiResult.good_points       || '',
    memo:             aiResult.memo              || '',
    ai_mentee_status: aiResult.ai_mentee_status  || 'green',
  });
}

// ============================================================
// Mentor: 担当メンティーの事前レポート一覧
// GET api/mentor/pre-reports?mentee_id=xxx
// ============================================================
function handleMentorPreReports(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var mentor_id = auth.payload.user_id;
  var body = parseBody_(e);
  var menteeId = body.mentee_id || '';

  // 担当メンティーのIDリストを取得
  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var myMenteeIds = users
    .filter(function(u){ return u.mentor_id === mentor_id; })
    .map(function(u){ return u.user_id; });

  var all = sheetToObjects(getSheet('pre_reports')).filter(function(r){ return r.report_id; });
  // 担当メンティーのみ
  all = all.filter(function(r){ return myMenteeIds.indexOf(r.user_id) >= 0; });
  if (menteeId) all = all.filter(function(r){ return r.user_id === menteeId; });
  all.sort(function(a,b){ return (b.submitted_at||'').localeCompare(a.submitted_at||''); });

  var userMap = {}; users.forEach(function(u){ userMap[u.user_id] = u; });
  var result = all.slice(0, 50).map(function(r){
    return Object.assign({}, r, { mentee_name: (userMap[r.user_id]||{}).name || r.user_id });
  });
  return jsonResponse({ ok:true, reports:result });
}


// ============================================================
// admin_memos シートに不足カラムを自動追加（一度だけ手動実行）
// ============================================================
function setupAdminMemosColumns() {
  var sheet = getSheet('admin_memos');
  if (!sheet) { Logger.log('admin_memos シートが見つかりません'); return; }
  var required = ['memo_id','admin_id','target_id','content',
    'is_notice','is_active','display_from','display_until','created_at','updated_at'];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var added = [];
  required.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      headers.push(col);
      added.push(col);
    }
  });
  Logger.log(added.length === 0 ? 'admin_memos: 全カラム揃っています' : '追加: ' + added.join(', '));
}


// ============================================================
// ★ appsscript.json に追加が必要なスコープ（GASエディタで設定）
// ============================================================
// {
//   "oauthScopes": [
//     "https://www.googleapis.com/auth/spreadsheets",
//     "https://www.googleapis.com/auth/gmail.send",
//     "https://www.googleapis.com/auth/script.external_request",
//     "https://www.googleapis.com/auth/chat.spaces.create",
//     "https://www.googleapis.com/auth/chat.memberships.create",
//     "https://www.googleapis.com/auth/chat.messages.create"
//   ]
// }
// ============================================================

// ============================================================
// Google カレンダー登録ヘルパー
// CalendarApp で双方のカレンダーにイベントを作成する
// ============================================================

/**
 * 1on1予約をGoogleカレンダーに登録する
 * @param {object} params - { title, startIso, durationMinutes, meetLink, mentorEmail, menteeEmail, note, bookingId }
 * @returns {{ ok: boolean, eventId: string|null, error: string|null }}
 */
function addCalendarEvent_(params) {
  try {
    var startDt    = new Date(params.startIso);
    var endDt      = new Date(startDt.getTime() + (params.durationMinutes || 60) * 60 * 1000);
    var title      = params.title      || '1on1 ミーティング';
    var meetLink   = params.meetLink   || '';
    var note       = params.note       || '';
    var bookingId  = params.bookingId  || '';

    var desc = '1on1管理システムから自動登録された予約です。';
    if (meetLink)  desc += '\n\nGoogle Meet: ' + meetLink;
    if (note)      desc += '\n\nメモ: ' + note;
    if (bookingId) desc += '\n\n予約ID: ' + bookingId;

    // ★ ゲストリスト: mentor.email(@socialshift.work) + mentee.email のみ
    // ※ mentor.calendar_email(@agent-network.com) は外部ドメインのため招待に含めない
    //   （外部ドメイン混在でMeetの待機室が発動するため）
    var guestList = [
      params.mentorEmail,   // @socialshift.work（ログインメアド）
      params.menteeEmail,   // メンティー
    ].filter(function(e){ return e && e.trim(); })
     .filter(function(e, i, arr){ return arr.indexOf(e) === i; }); // 重複除去

    var options = {
      description:            desc,
      guests:                 guestList.join(','),
      sendInvites:            true,
      guestCanInviteOthers:   true,   // ★ ゲストが待機中参加者を許可できる
      guestCanModify:         false,  // イベント編集は不可
      guestCanSeeOtherGuests: true,   // ゲストリストを閲覧可能
    };
    if (meetLink) {
      options.location = meetLink;
    }

    // CalendarApp はスクリプト実行者（デプロイオーナー）のカレンダーに登録し
    // guests でメンター・メンティーを招待する
    var calendar = CalendarApp.getDefaultCalendar();
    var event    = calendar.createEvent(title, startDt, endDt, options);

    // Google Meet URLを取得（GASのCalendarEventにはgetHangoutLink()がある）
    var autoMeetLink = '';
    try {
      // GASのCalendarEventオブジェクトのgetHangoutLink()を試みる
      if (typeof event.getHangoutLink === 'function') {
        autoMeetLink = event.getHangoutLink() || '';
      }
    } catch(meetErr) {
      Logger.log('MeetLink取得失敗（無視）: ' + meetErr.message);
    }

    return { ok: true, eventId: event.getId(), htmlLink: event.getHtmlLink() || '', meetLink: autoMeetLink, error: null };
  } catch (err) {
    Logger.log('カレンダー登録エラー: ' + err.message);
    return { ok: false, eventId: null, error: err.message };
  }
}

/**
 * カレンダーイベントを更新する（予約変更時）
 * @param {string} eventId - 既存イベントのID
 * @param {object} params  - { startIso, durationMinutes, meetLink, note }
 */
function updateCalendarEvent_(eventId, params) {
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var events   = calendar.getEventById(eventId);
    if (!events) {
      // IDが見つからない場合は新規作成にフォールバック
      Logger.log('イベントID ' + eventId + ' が見つからないため新規作成');
      return addCalendarEvent_(params);
    }

    var startDt = new Date(params.startIso);
    var endDt   = new Date(startDt.getTime() + (params.durationMinutes || 60) * 60 * 1000);
    events.setTime(startDt, endDt);

    var desc = '1on1管理システムから自動登録された予約です。（日程変更済み）';
    if (params.meetLink) desc += '\n\nGoogle Meet: ' + params.meetLink;
    if (params.note)     desc += '\n\n変更メモ: ' + params.note;
    events.setDescription(desc);
    if (params.meetLink) events.setLocation(params.meetLink);

    return { ok: true, eventId: eventId, error: null };
  } catch (err) {
    Logger.log('カレンダー更新エラー: ' + err.message);
    return { ok: false, eventId: null, error: err.message };
  }
}

/**
 * カレンダーイベントを削除する（予約キャンセル時）
 * @param {string} eventId
 */
/**
 * カレンダーイベントを削除する
 * 削除できない場合は【キャンセル1on1】マークを付与する
 */
function deleteCalendarEvent_(eventId) {
  if (!eventId) return;
  var cancelledAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var event    = calendar.getEventById(eventId);
    if (!event) {
      // CalendarAppで取得できない場合はREST APIで削除を試みる
      _deleteCalendarEventViaApi_(eventId, 'primary');
      return;
    }
    try {
      event.deleteEvent();
      Logger.log('[Calendar] メインイベント削除成功: ' + eventId);
    } catch (deleteErr) {
      Logger.log('[Calendar] メインイベント削除失敗、キャンセルマーク付与: ' + deleteErr.message);
      _markEventAsCancelled_(event, cancelledAt);
    }
  } catch (err) {
    Logger.log('[Calendar] メインイベント取得エラー eventId=' + eventId + ': ' + err.message);
  }
}

/**
 * REST API経由でカレンダーイベントを削除する（CalendarAppで取得できない場合のフォールバック）
 */
function _deleteCalendarEventViaApi_(eventId, calendarId) {
  try {
    var url = 'https://www.googleapis.com/calendar/v3/calendars/'
              + encodeURIComponent(calendarId)
              + '/events/' + encodeURIComponent(eventId);
    var response = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 204 || code === 200) {
      Logger.log('[CalendarAPI] 削除成功: ' + eventId + ' (calendar=' + calendarId + ')');
    } else {
      Logger.log('[CalendarAPI] 削除失敗 HTTP' + code + ' eventId=' + eventId);
    }
  } catch (e) {
    Logger.log('[CalendarAPI] 削除例外: ' + e.message);
  }
}

/**
 * カレンダーイベントにキャンセル済みマークを付与する
 * タイトルに【キャンセル1on1】、説明にキャンセル日時、色をグレーに変更
 */
function _markEventAsCancelled_(event, cancelledAt) {
  try {
    var originalTitle = event.getTitle();
    if (originalTitle.indexOf('【キャンセル1on1】') === -1) {
      event.setTitle('【キャンセル1on1】' + originalTitle);
    }
    var originalDesc = event.getDescription() || '';
    var originalDesc = event.getDescription() || '';
    event.setDescription('⚠️ この予約はキャンセルされました\nキャンセル日時: ' + cancelledAt + '\n\n' + originalDesc);
    event.setColor(CalendarApp.EventColor.GRAPHITE);
    Logger.log('[Calendar] キャンセルマーク付与: ' + event.getId());
  } catch (e) {
    Logger.log('[Calendar] キャンセルマーク付与失敗: ' + e.message);
  }
}

// ============================================================
// ★ 確認用カレンダーイベント作成（@agent-network.com 向け）
// ダブルブッキング防止用。Meetリンクなし、注意書きあり。
// ============================================================
function addSubCalendarEvent_(params) {
  // mentorCalendarEmail が未設定の場合はスキップ
  if (!params.mentorCalendarEmail) return { ok: false, eventId: null };
  try {
    var startDt = new Date(params.startIso);
    var endDt   = new Date(startDt.getTime() + (params.durationMinutes || 60) * 60 * 1000);
    var title   = '【確認用】' + params.title;
    var desc    = '※ この予定は ' + params.mentorEmail + ' に登録された1on1の確認用です。\n'
      + 'この予定を変更・削除したい場合は、' + params.mentorEmail + ' のカレンダーから操作してください。\n\n'
      + '予約ID: ' + (params.bookingId || '');
    // GASオーナーのカレンダーに作成し、mentorCalendarEmail のみゲスト招待
    var calendar = CalendarApp.getDefaultCalendar();
    var event    = calendar.createEvent(title, startDt, endDt, {
      description:            desc,
      guests:                 params.mentorCalendarEmail,
      sendInvites:            true,
      guestCanModify:         false,
      guestCanInviteOthers:   false,
      guestCanSeeOtherGuests: false,
    });
    Logger.log('確認用カレンダー作成: ' + params.mentorCalendarEmail + ' eventId=' + event.getId());
    return { ok: true, eventId: event.getId() };
  } catch(err) {
    Logger.log('確認用カレンダー作成エラー: ' + err.message);
    return { ok: false, eventId: null };
  }
}

// ★ Meet Recordings 内の「元ファイル」フォルダにファイルを移動
function moveToArchiveFolder_(fileId) {
  try {
    var meetFolder   = DriveApp.getFolderById(CONFIG.MEET_RECORDINGS_FOLDER_ID);
    var archiveName  = '元ファイル';
    var archiveIter  = meetFolder.getFoldersByName(archiveName);
    var archiveFolder = archiveIter.hasNext()
      ? archiveIter.next()
      : meetFolder.createFolder(archiveName);
    var file = DriveApp.getFileById(fileId);
    file.moveTo(archiveFolder);
    Logger.log('元ファイルへ移動: ' + file.getName());
  } catch(err) {
    Logger.log('元ファイル移動失敗（無視）: ' + err.message);
  }
}

function deleteSubCalendarEvent_(eventId) {
  if (!eventId) return;
  var cancelledAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var event    = calendar.getEventById(eventId);
    if (!event) {
      // CalendarAppで取得できない場合はREST APIで削除を試みる
      _deleteCalendarEventViaApi_(eventId, 'primary');
      return;
    }
    try {
      event.deleteEvent();
      Logger.log('[Calendar] 確認用イベント削除成功: ' + eventId);
    } catch (deleteErr) {
      Logger.log('[Calendar] 確認用イベント削除失敗、キャンセルマーク付与: ' + deleteErr.message);
      _markEventAsCancelled_(event, cancelledAt);
    }
  } catch(err) {
    Logger.log('[Calendar] 確認用イベント取得エラー eventId=' + eventId + ': ' + err.message);
  }
}

function updateSubCalendarEvent_(eventId, params) {
  if (!eventId) return { ok: false };
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var event    = calendar.getEventById(eventId);
    if (!event) { Logger.log('確認用イベントID未発見: ' + eventId); return { ok: false }; }
    var startDt = new Date(params.startIso);
    var endDt   = new Date(startDt.getTime() + (params.durationMinutes || 60) * 60 * 1000);
    event.setTime(startDt, endDt);
    Logger.log('確認用カレンダー更新: ' + eventId);
    return { ok: true };
  } catch(err) {
    Logger.log('確認用カレンダー更新エラー: ' + err.message);
    return { ok: false };
  }
}

// bookings シートに sub_calendar_event_id 列を追加（一回だけ手動実行）
function addSubCalendarEventIdColumn() {
  var sheet = getSheet('bookings');
  if (!sheet) { Logger.log('bookingsシートが見つかりません'); return; }
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('sub_calendar_event_id') >= 0) {
    Logger.log('sub_calendar_event_id列はすでに存在します');
    return;
  }
  sheet.getRange(1, sheet.getLastColumn()+1).setValue('sub_calendar_event_id');
  Logger.log('sub_calendar_event_id列を追加しました');
}

// mentor_reports シートに必要な列を追加（一回だけ手動実行）
function addMentorReportColumns() {
  var sheet = getSheet('mentor_reports');
  if (!sheet) { Logger.log('mentor_reportsシートが見つかりません'); return; }
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var toAdd = ['next_month_project_goal','next_month_study_goal','pdf_url','transcript','transcript_file_url','updated_at','target_month'];
  var added = [];
  toAdd.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      sheet.getRange(1, nextCol, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
      headers.push(col);
      added.push(col);
    }
  });
  Logger.log('追加完了: ' + (added.length > 0 ? added.join(', ') : 'なし（全列既存）'));
  Logger.log('現在の列数: ' + sheet.getLastColumn());
}


// カレンダーイベントID列をbookingsシートに追加（一回だけ手動実行）
function addCalendarEventIdColumn() {
  var sheet = getSheet('bookings');
  if (!sheet) { Logger.log('bookingsシートが見つかりません'); return; }
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('calendar_event_id') >= 0) {
    Logger.log('calendar_event_id列はすでに存在します');
    return;
  }
  var lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol+1).setValue('calendar_event_id');
  Logger.log('calendar_event_id列を追加しました（列' + (lastCol+1) + '）');
}

// ============================================================
// [MEET ARCHIVE] 追加定数 — 既存定数には手を加えない
// ============================================================
var MEET_ARCHIVE_MASTER_EMAIL    = 'test.admin@socialshift.work';
var MEET_ARCHIVE_SHARED_DRIVE_ID = '1xduD0ziGDXz64z6Iva29k65Lvgl48jPl'; // 共有ドライブID（既存CONFIG.SHARED_DRIVE_IDと同値）
var MEET_ARCHIVE_DEST_FOLDER_ID  = '1sCxZ7TFkGbAzfHZHF8k0LfqGmCqnjygs'; // 転送先フォルダID（既存CONFIG.RECORDINGS_FOLDER_IDと同値）
var MEET_ARCHIVE_CHAT_WEBHOOK    = '<<Google Chat Webhook URL>>';          // ★要設定: Google Chat Incoming Webhook URL
var MEET_ARCHIVE_PROCESSED_KEY   = 'meetArchive_processedFiles';           // 既存キー 'processed_file_ids' とは別キー

// ============================================================
// [MEET ARCHIVE] 定例会議設定（Script C 用）
// ============================================================
var MEET_ARCHIVE_MEETINGS = [
  { title: '全社定例',   day: 'MONDAY',    time: '10:00', durationMin: 60 },
  { title: '開発チーム', day: 'WEDNESDAY', time: '14:00', durationMin: 30 },
  { title: 'マーケ定例', day: 'FRIDAY',    time: '15:00', durationMin: 45 },
];
var MEET_ARCHIVE_GUESTS = ['team@socialshift.work']; // ★要設定: 招待するグループ or 個人メール

// ============================================================
// [MEET ARCHIVE] Script A: 録画を Shared Drive へ転送
// トリガー: 時間ベース・15分毎
// 実行アカウント: test.admin@socialshift.work にバインドした GAS プロジェクト
// ============================================================
function meetArchive_transferRecordings() {
  var prop      = PropertiesService.getScriptProperties();
  var processed = {};
  try {
    processed = JSON.parse(prop.getProperty(MEET_ARCHIVE_PROCESSED_KEY) || '{}');
  } catch (e) {
    Logger.log('[MEET ARCHIVE] processed プロパティのパースに失敗。空オブジェクトで続行。');
  }

  // マスターアカウントの "Meet Recordings" フォルダを取得
  var folders = DriveApp.getFoldersByName('Meet Recordings');
  if (!folders.hasNext()) {
    Logger.log('[MEET ARCHIVE] "Meet Recordings" フォルダが見つかりません。スキップ。');
    return;
  }
  var srcFolder = folders.next();

  // 転送先フォルダ（既存の RECORDINGS_FOLDER_ID とは別に設定可能）
  var destFolder;
  try {
    destFolder = DriveApp.getFolderById(MEET_ARCHIVE_DEST_FOLDER_ID);
  } catch (e) {
    Logger.log('[MEET ARCHIVE] 転送先フォルダの取得に失敗: ' + e.message);
    return;
  }

  var files = srcFolder.getFiles();
  var transferCount = 0;

  while (files.hasNext()) {
    var file = files.next();

    // 動画ファイル（録画）のみ対象
    if (file.getMimeType() !== 'video/mp4') continue;

    var fileId = file.getId();
    if (processed[fileId]) {
      Logger.log('[MEET ARCHIVE] 処理済みスキップ: ' + file.getName());
      continue;
    }

    // 共有ドライブへコピー → 元ファイルをゴミ箱へ（コピー成功時のみ）
    var copiedFile = null;
    try {
      copiedFile = file.makeCopy(file.getName(), destFolder);
      Logger.log('[MEET ARCHIVE] コピー完了: ' + file.getName() + ' → ' + copiedFile.getId());
    } catch (copyErr) {
      Logger.log('[MEET ARCHIVE] コピー失敗（スキップ）: ' + file.getName() + ' / ' + copyErr.message);
      continue; // コピー失敗時は元ファイルを削除しない
    }

    // コピー成功後のみ元ファイルをゴミ箱へ
    try {
      file.setTrashed(true);
      Logger.log('[MEET ARCHIVE] 元ファイルをゴミ箱へ: ' + file.getName());
    } catch (trashErr) {
      Logger.log('[MEET ARCHIVE] ゴミ箱移動失敗（コピー済み）: ' + file.getName() + ' / ' + trashErr.message);
    }

    processed[fileId] = new Date().toISOString();
    transferCount++;
    Logger.log('[MEET ARCHIVE] 転送完了: ' + file.getName());
  }

  // 処理済みIDを保存（30日以上前のエントリは削除して肥大化を防止）
  var cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  Object.keys(processed).forEach(function(id) {
    if (processed[id] < cutoff) delete processed[id];
  });
  prop.setProperty(MEET_ARCHIVE_PROCESSED_KEY, JSON.stringify(processed));

  Logger.log('[MEET ARCHIVE] 転送処理完了。転送件数: ' + transferCount);
}

// ============================================================
// [MEET ARCHIVE] Script B: 逸脱録画を検知して Chat へ通知
// トリガー: 時間ベース・毎日 9:00
// 実行アカウント: Super Admin 権限を持つアカウントにバインドした GAS プロジェクト
// 注意: Drive API でユーザーのマイドライブを横断検索するため
//       Domain-wide Delegation が有効な Super Admin アカウントで実行すること
// ============================================================
function meetArchive_detectStrayRecordings() {
  var token = ScriptApp.getOAuthToken();

  // 昨日以降に作成された Meet 録画ファイルを検索（マイドライブのみ・共有ドライブは除外）
  var yesterday = new Date(Date.now() - 86400000).toISOString();
  var query = encodeURIComponent(
    "mimeType='video/mp4' and createdTime > '" + yesterday + "' and name contains 'Meet'"
  );
  var url =
    'https://www.googleapis.com/drive/v3/files' +
    '?q=' + query +
    '&fields=files(id,name,owners,webViewLink)' +
    '&supportsAllDrives=false' +
    '&includeItemsFromAllDrives=false';

  var res;
  try {
    res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('[MEET ARCHIVE] Drive API フェッチ失敗: ' + e.message);
    return;
  }

  if (res.getResponseCode() !== 200) {
    Logger.log('[MEET ARCHIVE] Drive API エラー (' + res.getResponseCode() + '): ' + res.getContentText());
    return;
  }

  var files = [];
  try {
    files = JSON.parse(res.getContentText()).files || [];
  } catch (parseErr) {
    Logger.log('[MEET ARCHIVE] レスポンスのパース失敗: ' + parseErr.message);
    return;
  }

  // マスターアカウント以外が所有する録画を抽出
  var strays = files.filter(function(f) {
    return f.owners && f.owners[0] && f.owners[0].emailAddress !== MEET_ARCHIVE_MASTER_EMAIL;
  });

  if (strays.length === 0) {
    Logger.log('[MEET ARCHIVE] 逸脱録画なし。Chat への通知はスキップ。');
    return;
  }

  // Chat メッセージ組み立て
  var lines = strays.map(function(f) {
    return '• *' + f.name + '*\n  所有者: ' + f.owners[0].emailAddress + '\n  ' + (f.webViewLink || '（URLなし）');
  });
  var message =
    '⚠️ *逸脱録画を検知しました（' + strays.length + ' 件）*\n\n' +
    lines.join('\n\n') +
    '\n\n※ 上記ファイルを共有ドライブへ手動で移動してください。';

  // Webhook URL が未設定の場合はログのみ
  if (!MEET_ARCHIVE_CHAT_WEBHOOK || MEET_ARCHIVE_CHAT_WEBHOOK.indexOf('<<') === 0) {
    Logger.log('[MEET ARCHIVE] Webhook URL 未設定。ログのみ出力:');
    Logger.log(message);
    return;
  }

  try {
    UrlFetchApp.fetch(MEET_ARCHIVE_CHAT_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true,
    });
    Logger.log('[MEET ARCHIVE] Chat アラート送信完了: ' + strays.length + ' 件');
  } catch (chatErr) {
    Logger.log('[MEET ARCHIVE] Chat 通知送信失敗: ' + chatErr.message);
  }
}

// ============================================================
// [MEET ARCHIVE] Script C: 定例会議をマスターカレンダーに登録
// トリガー: 手動実行（または必要に応じて週次トリガー）
// 実行アカウント: test.admin@socialshift.work にバインドした GAS プロジェクト
// ============================================================
function meetArchive_createWeeklyMeetings() {
  var cal = CalendarApp.getDefaultCalendar();

  MEET_ARCHIVE_MEETINGS.forEach(function(m) {
    var start = meetArchive_getNextWeekday(m.day, m.time);
    var end   = new Date(start.getTime() + m.durationMin * 60000);

    try {
      cal.createEvent(m.title, start, end, {
        guests:      MEET_ARCHIVE_GUESTS.join(','),
        sendInvites: true,
        description:
          '【録画について】このミーティングはマスターアカウントが主催しているため、' +
          '録画は自動的に共有ドライブ（Meet Recordings Archive）に保存されます。',
      });
      Logger.log('[MEET ARCHIVE] 予定作成: ' + m.title + ' / ' + start);
    } catch (calErr) {
      Logger.log('[MEET ARCHIVE] 予定作成失敗: ' + m.title + ' / ' + calErr.message);
    }
  });
}

// ============================================================
// [MEET ARCHIVE] 補助: 次の指定曜日・時刻の Date を返す
// ============================================================
function meetArchive_getNextWeekday(dayName, timeStr) {
  var dayMap = {
    SUNDAY: 0, MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3,
    THURSDAY: 4, FRIDAY: 5, SATURDAY: 6,
  };
  var target = dayMap[dayName];
  if (target === undefined) {
    Logger.log('[MEET ARCHIVE] 不正な曜日名: ' + dayName);
    return new Date();
  }
  var now  = new Date();
  var diff = (target - now.getDay() + 7) % 7 || 7;
  var d    = new Date(now);
  d.setDate(d.getDate() + diff);
  var parts = timeStr.split(':').map(Number);
  d.setHours(parts[0], parts[1], 0, 0);
  return d;
}

// ============================================================
// [MEET ARCHIVE] トリガー設定補助（GASエディタから手動実行）
// 既存トリガーは削除・変更しない。新規トリガーのみ追加する。
// ============================================================
function meetArchive_setupTriggers() {
  // 既存トリガーの関数名を取得して誤削除を防止
  var existingHandlers = ScriptApp.getProjectTriggers().map(function(t) {
    return t.getHandlerFunction();
  });
  Logger.log('[MEET ARCHIVE] 既存トリガー: ' + existingHandlers.join(', '));

  // meetArchive_ 系の重複トリガーがあれば削除（冪等性）
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction().indexOf('meetArchive_') === 0) {
      ScriptApp.deleteTrigger(t);
      Logger.log('[MEET ARCHIVE] 既存 meetArchive_ トリガー削除: ' + t.getHandlerFunction());
    }
  });

  // Script A: 15分毎
  ScriptApp.newTrigger('meetArchive_transferRecordings')
    .timeBased().everyMinutes(15).create();
  Logger.log('[MEET ARCHIVE] トリガー追加: meetArchive_transferRecordings (15分毎)');

  // Script B: 毎日 9:00
  ScriptApp.newTrigger('meetArchive_detectStrayRecordings')
    .timeBased().everyDays(1).atHour(9).create();
  Logger.log('[MEET ARCHIVE] トリガー追加: meetArchive_detectStrayRecordings (毎日9時)');

  Logger.log('[MEET ARCHIVE] トリガー設定完了');
}

// ============================================================
// [CHAT WEBHOOK] スプシ users に chat_space_id / chat_webhook_url 列を追加
// GASエディタから一度だけ手動実行
// ============================================================
function meetArchive_setupChatColumns() {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('users');
  if (!sheet) { Logger.log('users シートが見つかりません'); return; }

  ['chat_space_id', 'chat_webhook_url'].forEach(function(col) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf(col) < 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      Logger.log('[CHAT WEBHOOK] 列追加: ' + col);
    } else {
      Logger.log('[CHAT WEBHOOK] 既存列: ' + col);
    }
  });
  Logger.log('[CHAT WEBHOOK] meetArchive_setupChatColumns 完了');
}

// ============================================================
// [CHAT WEBHOOK] chat_url から space_id を抽出するユーティリティ
// https://chat.google.com/room/SPACE_ID/... → SPACE_ID
// https://chat.google.com/u/0/room/SPACE_ID/... → SPACE_ID
// ============================================================
function meetArchive_extractSpaceId(chatUrl) {
  if (!chatUrl) return '';
  // /room/SPACE_ID or /dm/SPACE_ID にマッチ
  var m = String(chatUrl).match(/\/(?:room|dm)\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : '';
}

// ============================================================
// [CHAT WEBHOOK] chat_url 登録時に chat_space_id を自動抽出して保存
// handleCreateChatSpace の処理後に呼ばれる（既存関数を上書きせず拡張）
// POST api/admin/chat-space/register { user_id, chat_url }
// ============================================================
// ルート追加は routeRequest の末尾に記載
function handleRegisterChatSpace(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var user_id  = (body.user_id  || '').trim();
  var chat_url = (body.chat_url || '').trim();
  if (!user_id)  return errorResponse('MISSING_USER_ID', 400);
  if (!chat_url) return errorResponse('MISSING_CHAT_URL', 400);
  if (!chat_url.includes('chat.google.com'))
    return errorResponse('INVALID_CHAT_URL', 400);

  // space_id を自動抽出
  var space_id = meetArchive_extractSpaceId(chat_url);
  if (!space_id) {
    Logger.log('[CHAT WEBHOOK] space_id 抽出失敗: ' + chat_url);
  }

  // users シートに chat_url + chat_space_id を保存
  updateRowWhere('users', 'user_id', user_id, {
    chat_url:      chat_url,
    chat_space_id: space_id,
    updated_at:    new Date().toISOString(),
  });
  invalidateCache_('users');

  Logger.log('[CHAT WEBHOOK] ChatSpace登録: user=' + user_id + ' space=' + space_id);
  return jsonResponse({ ok: true, chat_url: chat_url, chat_space_id: space_id });
}

// ============================================================
// [CHAT WEBHOOK] Webhook URL を保存する
// POST api/admin/chat-webhook/save { user_id, webhook_url }
// ============================================================
function handleSaveChatWebhook(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var user_id     = (body.user_id     || '').trim();
  var webhook_url = (body.webhook_url || '').trim();
  if (!user_id)     return errorResponse('MISSING_USER_ID', 400);
  if (!webhook_url) return errorResponse('MISSING_WEBHOOK_URL', 400);
  if (!webhook_url.includes('chat.googleapis.com'))
    return errorResponse('INVALID_WEBHOOK_URL', 400);

  updateRowWhere('users', 'user_id', user_id, {
    chat_webhook_url: webhook_url,
    updated_at:       new Date().toISOString(),
  });
  invalidateCache_('users');

  Logger.log('[CHAT WEBHOOK] Webhook保存: user=' + user_id);
  return jsonResponse({ ok: true });
}

// ============================================================
// [CHAT WEBHOOK] ユーザーの Webhook URL を取得する
// GET api/admin/chat-webhook?user_id=xxx
// ============================================================
function handleGetChatWebhook(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var user_id = body.user_id || (e.parameter && e.parameter.user_id) || '';
  if (!user_id) return errorResponse('MISSING_USER_ID', 400);

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var target = users.find(function(u){ return u.user_id === user_id; });
  if (!target) return errorResponse('USER_NOT_FOUND', 404);

  return jsonResponse({
    ok:               true,
    chat_url:         target.chat_url         || '',
    chat_space_id:    target.chat_space_id    || '',
    chat_webhook_url: target.chat_webhook_url || '',
    has_webhook:      !!(target.chat_webhook_url),
  });
}

// ============================================================
// [CHAT WEBHOOK] 指定ユーザーの個別スペースへ通知を送信する
// 他の関数から呼び出す共通ユーティリティ
// @param {string} userId     - users シートの user_id
// @param {string} message    - 送信するテキスト（Markdown可）
// @returns {{ ok: boolean, error?: string }}
// ============================================================
function meetArchive_sendToMenteeSpace(userId, message) {
  if (!userId || !message) return { ok: false, error: 'MISSING_PARAMS' };

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var target = users.find(function(u){ return u.user_id === userId; });
  if (!target) return { ok: false, error: 'USER_NOT_FOUND' };

  var webhookUrl = target.chat_webhook_url || '';
  if (!webhookUrl) {
    Logger.log('[CHAT WEBHOOK] Webhook未設定: user=' + userId);
    return { ok: false, error: 'WEBHOOK_NOT_SET' };
  }

  try {
    var res = UrlFetchApp.fetch(webhookUrl, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify({ text: message }),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('[CHAT WEBHOOK] 送信失敗 HTTP ' + code + ': ' + res.getContentText());
      return { ok: false, error: 'HTTP_' + code };
    }
    Logger.log('[CHAT WEBHOOK] 送信完了: user=' + userId);
    return { ok: true };
  } catch(err) {
    Logger.log('[CHAT WEBHOOK] 送信エラー: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ============================================================
// [CHAT WEBHOOK] 全メンバーの Webhook 設定状況を確認する（診断用）
// GASエディタから手動実行してログで確認
// ============================================================
function meetArchive_checkWebhookStatus() {
  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id && u.role === 'mentee'; });
  Logger.log('[CHAT WEBHOOK] メンティー数: ' + users.length);
  users.forEach(function(u) {
    var status = u.chat_webhook_url ? '✅ 設定済み' : '❌ 未設定';
    Logger.log('[CHAT WEBHOOK] ' + status + ' | ' + u.name + ' | space=' + (u.chat_space_id || '-'));
  });
}

// ============================================================
// Google Drive 個人フォルダ ヘルパー（共有ドライブ対応版）
// ============================================================
function getOrCreatePersonalFolders_(menteeName) {
  var token      = ScriptApp.getOAuthToken();
  var personalId = driveApiGetOrCreateFolder_(CONFIG.PERSONAL_FOLDER_ROOT_ID, menteeName, token);
  var recordsId  = driveApiGetOrCreateFolder_(personalId,  '記録管理',      token);
  var mediaId    = driveApiGetOrCreateFolder_(personalId,  '動画＆録音管理', token);
  // DriveApp.getFolderById互換オブジェクトを返す（createFileのみ必要）
  return {
    personal: { getId: function(){ return personalId; } },
    records:  {
      getId: function(){ return recordsId; },
      createFile: function(blob) {
        return driveApiUploadBlob_(blob, recordsId, token);
      }
    },
    media: { getId: function(){ return mediaId; } },
  };
}

// Phase 5 で名称統一: ensurePersonalFolders_ は getOrCreatePersonalFolders_ のエイリアス
var ensurePersonalFolders_ = getOrCreatePersonalFolders_;

function findOrCreateFolder_(parent, name) {
  // 後方互換用（DriveAppフォルダオブジェクトが渡された場合）
  if (parent && typeof parent.getFoldersByName === 'function') {
    var it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
  }
  // Drive API版（IDが渡された場合）
  var token = ScriptApp.getOAuthToken();
  var parentId = (typeof parent === 'string') ? parent : parent.getId();
  return driveApiGetOrCreateFolder_(parentId, name, token);
}

// ── Drive API: BlobをPDFとして共有ドライブフォルダにアップロード ──
function driveApiUploadBlob_(blob, parentFolderId, token) {
  try {
    var boundary  = 'boundary_1on1_pdf';
    var metaJson  = JSON.stringify({ name: blob.getName(), parents: [parentFolderId] });
    var blobBytes = blob.getBytes();
    // multipart upload
    var bodyParts = '--' + boundary + '\r\n'
      + 'Content-Type: application/json; charset=UTF-8\r\n\r\n'
      + metaJson + '\r\n'
      + '--' + boundary + '\r\n'
      + 'Content-Type: ' + blob.getContentType() + '\r\n'
      + 'Content-Transfer-Encoding: base64\r\n\r\n'
      + Utilities.base64Encode(blobBytes) + '\r\n'
      + '--' + boundary + '--';
    var res = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files'
        + '?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        payload: bodyParts,
        muteHttpExceptions: true
      }
    ).getContentText());
    Logger.log('driveApiUploadBlob_: ' + blob.getName() + ' → ' + res.id);
    // DriveApp.File互換オブジェクトを返す
    return {
      getId:    function(){ return res.id || ''; },
      getUrl:   function(){ return res.webViewLink || ''; },
      getName:  function(){ return blob.getName(); }
    };
  } catch(e) {
    Logger.log('driveApiUploadBlob_ error: ' + e.message);
    throw e;
  }
}

// ============================================================
// Phase 4: 1on1レポート PDF 生成
// Google Doc を作成して PDF 変換し、個人フォルダの記録管理に保存
// @param {object} report - mentor_reports 行データ
// @param {string} menteeName - メンティー名
// @param {string} mentorName - メンター名
// @return {string} PDF ファイルの URL
// ============================================================
function generateMentorReportPdf_(report, menteeName, mentorName) {
  try {
    // ファイル名: YYMM_月次_メンティー名
    var dt  = report.created_at ? new Date(report.created_at) : new Date();
    var yy  = String(dt.getFullYear()).slice(-2);
    var mm  = ('0' + (dt.getMonth() + 1)).slice(-2);
    var title = yy + mm + '_月次_' + menteeName;

    // Google Doc を作成してコンテンツを書き込む
    var doc  = DocumentApp.create(title);
    var body = doc.getBody();

    body.getParagraphs()[0].setText(title);
    body.getParagraphs()[0].setHeading(DocumentApp.ParagraphHeading.HEADING1);

    var meta = 'メンティー: ' + menteeName + '　／　メンター: ' + (mentorName || '—')
      + (report.target_month ? '　／　実施月: ' + report.target_month : '');
    body.appendParagraph(meta).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 10 });
    body.appendHorizontalRule();

    var sections = [
      { label: 'AIサマリー',           val: report.ai_summary               },
      { label: 'メンターからのアドバイス', val: report.ai_advice                 },
      { label: '次回の目標',            val: report.next_goal                 },
      { label: '次月プロジェクト目標',   val: report.next_month_project_goal   },
      { label: '次月学習目標',          val: report.next_month_study_goal     },
    ];
    sections.forEach(function(s) {
      if (!s.val) return;
      body.appendParagraph('【' + s.label + '】').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(String(s.val));
    });

    doc.saveAndClose();

    // PDF 変換
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(title + '.pdf');

    // 個人フォルダ/記録管理 に保存
    var folders = getOrCreatePersonalFolders_(menteeName);
    var pdfFile = folders.records.createFile(pdfBlob);

    // 元の Google Doc を削除（PDF のみ残す）
    docFile.setTrashed(true);

    return pdfFile.getUrl();
  } catch (err) {
    Logger.log('generateMentorReportPdf_ error: ' + err.message);
    return '';
  }
}

// ============================================================
// Admin: ユーザー更新（新エンドポイント）
// POST api/admin/users/update
// ============================================================
function handleAdminUpdateUser(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  if (!body.user_id) return errorResponse('MISSING_USER_ID', 400);

  var updates = { updated_at: new Date().toISOString() };
  var allowed = ['name','role','has_leader_role','mentor_id','leader_id','phone_number','workplace',
                 'work_status','hourly_wage','status','birthday','chat_url',
                 'goal_work_6m','goal_skill_6m','current_project','default_1on1_duration'];
  allowed.forEach(function(f) {
    if (body[f] !== undefined) updates[f] = String(body[f]);
  });
  if (body.has_leader_role !== undefined) {
    updates.has_leader_role = body.has_leader_role ? 'TRUE' : 'FALSE';
  }

  updateRowWhere('users', 'user_id', body.user_id, updates);
  invalidateCache_('users');
  return jsonResponse({ ok: true });
}

// ============================================================
// Admin: ユーザー一括エクスポート
// GET api/admin/users/export
// ============================================================
function handleAdminUsersExport(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var body = parseBody_(e);
  var userIds = body.user_ids || null;  // nullの場合は全員
  var fields  = body.fields  || ['name','email','role','workplace','work_status','phone_number','hire_date','status','mentor_id'];

  // password_hash は除外
  fields = fields.filter(function(f){ return f !== 'password_hash'; });

  var users = cachedSheetToObjects_('users').filter(function(u) {
    return u.user_id && String(u.status).toLowerCase() !== 'deleted';
  });
  if (userIds && userIds.length > 0) {
    users = users.filter(function(u){ return userIds.indexOf(u.user_id) >= 0; });
  }

  // スプレッドシートに新シートを作成
  var ss = getSpreadsheet_();
  var now = new Date();
  var stamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  var sheetName = 'エクスポート_' + stamp;
  var newSheet = ss.insertSheet(sheetName);

  // ヘッダー行
  newSheet.getRange(1, 1, 1, fields.length).setValues([fields]);
  newSheet.getRange(1, 1, 1, fields.length).setFontWeight('bold');

  // データ行
  if (users.length > 0) {
    var rows = users.map(function(u) {
      return fields.map(function(f){ return u[f] || ''; });
    });
    newSheet.getRange(2, 1, rows.length, fields.length).setValues(rows);
  }

  var sheetUrl = ss.getUrl() + '#gid=' + newSheet.getSheetId();
  return jsonResponse({ ok: true, sheet_url: sheetUrl, count: users.length });
}

// ============================================================
// Admin: ユーザー一括インポート
// ============================================================
// Admin: ユーザー一括インポート（upsert: 新規追加 + 既存更新）
// POST api/admin/users/bulk-import
// ・CSVのemailで既存ユーザーを検索
//   → 存在しない: 新規追加（passwordは必須）
//   → 存在する:   CSVのデータで上書き更新（passwordは省略可）
// ・既存ユーザーでCSVに含まれないユーザーは変更なし（削除されない）
// ============================================================
// ============================================================
// Admin: パスワードリセット
// POST api/admin/user/reset-password
// body: { user_id, new_password_hash? }
//   new_password_hash が省略された場合は生年月日（YYYYMMDD）ハッシュにリセット
// ============================================================
function handleAdminResetPassword(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body    = parseBody_(e);
  var user_id = (body.user_id || '').trim();
  if (!user_id) return errorResponse('MISSING_USER_ID', 400);

  var users  = sheetToObjects(getSheet('users'));
  var target = users.find(function(u){ return u.user_id === user_id; });
  if (!target) return errorResponse('USER_NOT_FOUND', 404);

  var newHash = '';
  if (body.new_password_hash) {
    // 管理者が任意のパスワードハッシュを指定
    newHash = normalizePasswordHash(body.new_password_hash);
  } else {
    // 生年月日（YYYYMMDD）のSHA-256にリセット
    var birthday = (target.birthday || '').replace(/-/g, '').trim();
    if (!birthday || birthday.length < 8) {
      return errorResponse('NO_BIRTHDAY: 生年月日が未設定のためリセットできません', 400);
    }
    newHash = sha256Hash(birthday);
  }

  updateRowWhere('users', 'user_id', user_id, {
    password_hash: newHash,
    updated_at:    new Date().toISOString()
  });
  invalidateCache_('users');
  Logger.log('handleAdminResetPassword: user_id=' + user_id + ' by admin=' + auth.payload.user_id);
  return jsonResponse({ ok: true, user_name: target.name });
}

// ============================================================
// Admin: ユーザー一括更新
// POST api/admin/users/bulk-update
// body: { updates: [{ user_id, mentor_id?, leader_id?, ... }] }
// ============================================================
function handleAdminBulkUpdate(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body    = parseBody_(e);
  var updates = body.updates;
  if (!Array.isArray(updates) || updates.length === 0)
    return errorResponse('MISSING_UPDATES', 400);

  var now     = new Date().toISOString();
  var updated = 0;
  var errors  = [];
  var ALLOWED = ['mentor_id','leader_id','role','employment_type',
                 'workplace','hourly_wage','status','has_leader_role','phone','calendar_email','personal_folder_id'];

  // leader_assignments 同期に必要な現在のusers情報を取得
  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u; });

  updates.forEach(function(u) {
    if (!u.user_id) { errors.push('user_id missing'); return; }
    var fields = { updated_at: now };
    ALLOWED.forEach(function(k) {
      if (u[k] !== undefined) fields[k] = String(u[k]);
    });
    if (Object.keys(fields).length <= 1) return;
    try {
      updateRowWhere('users', 'user_id', u.user_id, fields);
      updated++;

      // ★ leader_id が変更された場合は leader_assignments シートも同期
      if (u.leader_id !== undefined) {
        var current   = userMap[u.user_id] || {};
        var oldLeader = String(current.leader_id || '');
        var newLeader = String(u.leader_id || '');
        var role      = u.role !== undefined ? String(u.role) : String(current.role || '');

        if (role === 'mentee' && newLeader !== oldLeader) {
          var laSheet = getSheet('leader_assignments');
          if (laSheet) {
            var laData    = laSheet.getDataRange().getValues();
            var laHeaders = laData[0];
            var mCol = laHeaders.indexOf('mentee_id');
            var lCol = laHeaders.indexOf('leader_id');
            // 既存のこのメンティーの割り当てを削除（後ろから）
            for (var ri = laData.length - 1; ri >= 1; ri--) {
              if (String(laData[ri][mCol]) === String(u.user_id)) {
                laSheet.deleteRow(ri + 1);
              }
            }
            // 新しいリーダーがある場合は追加
            if (newLeader) {
              var la_id = 'la-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
              appendRow('leader_assignments', {
                assignment_id: la_id,
                leader_id:     newLeader,
                mentee_id:     String(u.user_id),
                created_at:    now,
                updated_at:    now,
              });
            }
            Logger.log('handleAdminBulkUpdate: leader_assignments 同期 mentee=' + u.user_id + ' ' + oldLeader + '→' + newLeader);
          }
        }
      }
    } catch(err) {
      errors.push(u.user_id + ': ' + err.message);
    }
  });

  invalidateCache_('users');
  invalidateCache_('leader_assignments');
  Logger.log('handleAdminBulkUpdate: ' + updated + '件更新 errors=' + errors.length);
  return jsonResponse({ ok: true, updated: updated, errors: errors });
}

// ============================================================
// ★ セットアップ用：usersシートのleader_idを leader_assignments に一括同期
// GASエディタから手動で1回実行すると、users.leader_id に設定済みの
// リーダー紐づけが leader_assignments シートに反映される
// ============================================================
function syncLeaderAssignmentsFromUsers() {
  Logger.log('=== syncLeaderAssignmentsFromUsers 開始 ===');
  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentees = users.filter(function(u){ return u.role === 'mentee' && u.leader_id; });

  // 既存の leader_assignments を全取得
  var laSheet = getSheet('leader_assignments');
  if (!laSheet) { Logger.log('leader_assignments シートが存在しません'); return; }

  var laData    = laSheet.getDataRange().getValues();
  var laHeaders = laData[0];
  var mCol      = laHeaders.indexOf('mentee_id');

  // 既に登録済みのメンティーIDセット
  var existing = {};
  for (var i = 1; i < laData.length; i++) {
    if (laData[i][mCol]) existing[String(laData[i][mCol])] = true;
  }

  var added = 0;
  var now   = new Date().toISOString();
  mentees.forEach(function(u) {
    if (existing[u.user_id]) {
      Logger.log('スキップ（既存）: ' + u.name + ' → leader_id=' + u.leader_id);
      return;
    }
    var la_id = 'la-sync-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
    appendRow('leader_assignments', {
      assignment_id: la_id,
      leader_id:     u.leader_id,
      mentee_id:     u.user_id,
      created_at:    now,
      updated_at:    now,
    });
    Logger.log('追加: ' + u.name + ' → leader=' + u.leader_id);
    added++;
  });

  invalidateCache_('leader_assignments');
  Logger.log('=== syncLeaderAssignmentsFromUsers 完了: ' + added + '件追加 ===');
}

function handleAdminBulkImport(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var rows = body.rows;
  if (!rows || !Array.isArray(rows) || rows.length === 0) {
    return errorResponse('MISSING_ROWS', 400);
  }

  var results      = [];
  var now          = new Date().toISOString();
  var addedCount   = 0;
  var updatedCount = 0;

  // ロック取得（一括処理中の競合防止）
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    // シートを一度だけ読み込んでメモリ上で操作
    var _r       = getUsersRaw_();
    var sheet    = _r.sheet;
    var headers  = _r.headers;

    rows.forEach(function(row) {
      if (!row.name || !row.email || !row.role) {
        results.push({ email: row.email || '?', ok: false, action: 'skip', error: 'name/email/role は必須です' });
        return;
      }
      if (['mentee','mentor','admin'].indexOf(String(row.role)) < 0) {
        results.push({ email: row.email, ok: false, action: 'skip', error: '無効なロール: ' + row.role });
        return;
      }

      var emailLower = String(row.email).toLowerCase().trim();

      // メールで既存ユーザーを検索（deleted 以外）
      var existing = _r.rows.find(function(u) {
        return String(u.email || '').toLowerCase() === emailLower
          && String(u.status || '').toLowerCase() !== 'deleted'
          && u.user_id;
      });

      if (existing) {
        // ── 既存ユーザー: 更新（upsert の update 部分）──
        // CSVに値がある列だけ上書き。空欄の場合は既存値を維持
        var updates = {
          name:        row.name        ? String(row.name).trim()        : String(existing.name        || ''),
          role:        row.role        ? String(row.role)               : String(existing.role        || ''),
          phone_number:row.phone_number? String(row.phone_number).trim(): String(existing.phone_number|| ''),
          workplace:   row.workplace   ? String(row.workplace).trim()   : String(existing.workplace   || ''),
          work_status: row.work_status ? String(row.work_status)        : String(existing.work_status || ''),
          hourly_wage: row.hourly_wage ? Number(row.hourly_wage) || ''  : existing.hourly_wage,
          birthday:    row.birthday    ? String(row.birthday)           : String(existing.birthday    || ''),
          hire_date:   row.hire_date   ? String(row.hire_date)          : String(existing.hire_date   || ''),
          mentor_id:   row.mentor_id   !== undefined ? String(row.mentor_id || '') : String(existing.mentor_id || ''),
          updated_at:  now,
        };
        // パスワードが指定されている場合のみ更新
        if (row.password) {
          updates.password_hash = sha256Hash(String(row.password));
        }
        // 1行まとめて書き込み
        var currentRow = headers.map(function(h) { return existing[h] !== undefined ? existing[h] : ''; });
        headers.forEach(function(h, i) { if (updates.hasOwnProperty(h)) currentRow[i] = updates[h]; });
        sheet.getRange(existing.__rowIndex, 1, 1, headers.length).setValues([currentRow]);
        updatedCount++;
        results.push({ email: row.email, ok: true, action: 'updated', user_id: existing.user_id });

      } else {
        // ── 新規ユーザー: 追加（upsert の insert 部分）──
        if (!row.password) {
          results.push({ email: row.email, ok: false, action: 'skip', error: '新規ユーザーにはパスワードが必須です' });
          return;
        }
        var user_id = 'u-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
        appendRow('users', {
          user_id:              user_id,
          email:                emailLower,
          name:                 String(row.name).trim(),
          role:                 String(row.role),
          has_leader_role:      'FALSE',
          password_hash:        sha256Hash(String(row.password)),
          mentor_id:            row.mentor_id    || '',
          leader_id:            '',
          phone_number:         row.phone_number || '',
          workplace:            row.workplace    || '',
          work_status:          row.work_status  || '',
          hourly_wage:          row.hourly_wage  ? Number(row.hourly_wage) || '' : '',
          status:               'active',
          created_at:           now,
          updated_at:           now,
          birthday:             row.birthday  || '',
          hire_date:            row.hire_date || '',
          chat_url:             '',
          default_1on1_duration:'60',
        });
        addedCount++;
        results.push({ email: row.email, ok: true, action: 'added', user_id: user_id });
      }
    });

  } finally {
    lock.releaseLock();
    invalidateCache_('users');
  }

  return jsonResponse({
    ok:            true,
    results:       results,
    added_count:   addedCount,
    updated_count: updatedCount,
    success_count: addedCount + updatedCount,
    total:         rows.length,
  });
}

// ============================================================
// Admin: TEL用Meet URL設定
// POST api/admin/set-tel-meet-url { mentee_id, tel_meet_url }
// usersシートの tel_meet_url 列に保存し、Chatスペースにピン止め
// ============================================================
function handleSetTelMeetUrl(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var mentee_id    = (body.mentee_id    || '').trim();
  var tel_meet_url = (body.tel_meet_url || '').trim();
  var chat_url     = (body.chat_url     || '').trim();
  var webhook_url  = (body.webhook_url  || '').trim();

  if (!mentee_id) return errorResponse('MISSING_MENTEE_ID', 400);

  // tel_meet_urlが指定されている場合はURL形式チェック
  if (tel_meet_url && tel_meet_url.indexOf('meet.google.com') < 0) {
    return errorResponse('INVALID_URL: Google Meet URLを入力してください', 400);
  }
  // chat_urlもwebhook_urlも tel_meet_urlもない場合はエラー
  if (!tel_meet_url && !chat_url && !webhook_url) {
    return errorResponse('MISSING_PARAMS', 400);
  }

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === mentee_id; });
  if (!mentee) return errorResponse('USER_NOT_FOUND', 404);

  // 更新フィールドを構築（指定されたものだけ更新）
  var updates = { updated_at: new Date().toISOString() };
  if (tel_meet_url) updates.tel_meet_url     = tel_meet_url;
  if (chat_url)     updates.chat_url         = chat_url;
  if (webhook_url)  updates.chat_webhook_url = webhook_url;

  // chat_urlからspace_idを抽出して保存
  if (chat_url) {
    var m = chat_url.match(/\/room\/([^\/\?]+)/);
    if (!m) m = chat_url.match(/\/spaces\/([^\/\?]+)/);
    if (!m) m = chat_url.match(/\/chat\/([^\/\?]+)/);
    if (m) updates.chat_space_id = m[1];
  }

  updateRowWhere('users', 'user_id', mentee_id, updates);
  invalidateCache_('users');

  // Webhookでピン止め通知（webhook_urlが確定している場合）
  var effectiveWebhook = webhook_url || mentee.chat_webhook_url || '';
  var pinResult = { ok: false, reason: 'no_webhook' };
  if (effectiveWebhook && tel_meet_url) {
    try {
      var res = UrlFetchApp.fetch(effectiveWebhook, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: '📌 *TEL用Meet URL*\n' + tel_meet_url + '\n\n毎週のTELはこのURLを使ってください。'
        }),
        muteHttpExceptions: true
      });
      pinResult = { ok: res.getResponseCode() === 200 };
    } catch(webhookErr) {
      pinResult = { ok: false, reason: webhookErr.message };
    }
  }

  Logger.log('TEL情報設定: ' + mentee.name + ' chat=' + !!chat_url + ' webhook=' + !!webhook_url + ' meet=' + !!tel_meet_url);
  return jsonResponse({
    ok:          true,
    mentee_id:   mentee_id,
    mentee_name: mentee.name,
    tel_meet_url: tel_meet_url,
    pin_result:  pinResult
  });
}

// ============================================================
// TEL スペースストック管理
// ============================================================
// シート: tel_space_stock
// 列: stock_id, chat_url, chat_space_id, webhook_url,
//     meet_url, status(available/assigned), assigned_to, assigned_at, created_at

function handleGetTelStock(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var sheet = getSheet('tel_space_stock');
  if (!sheet) return jsonResponse({ ok: true, stocks: [], available_count: 0 });

  var stocks = sheetToObjects(sheet).filter(function(s){ return s.stock_id; });
  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u; });

  var result = stocks.map(function(s) {
    return {
      stock_id:     s.stock_id,
      chat_url:     s.chat_url    || '',
      webhook_url:  s.webhook_url || '',
      meet_url:     s.meet_url    || '',
      status:       s.status      || 'available',
      assigned_to:  s.assigned_to || '',
      assigned_name: s.assigned_to ? ((userMap[s.assigned_to] || {}).name || s.assigned_to) : '',
      assigned_at:  s.assigned_at || '',
      created_at:   s.created_at  || '',
    };
  });
  var available = result.filter(function(s){ return s.status === 'available'; }).length;
  return jsonResponse({ ok: true, stocks: result, available_count: available });
}

function handleAddTelStock(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var items = body.items || [];
  if (!Array.isArray(items) || items.length === 0) return errorResponse('MISSING_ITEMS', 400);

  var ss    = getSpreadsheet_();
  var sheet = getSheet('tel_space_stock');
  if (!sheet) {
    sheet = ss.insertSheet('tel_space_stock');
    sheet.appendRow(['stock_id','chat_url','chat_space_id','webhook_url','meet_url','status','assigned_to','assigned_at','created_at']);
  }

  var added = 0;
  var now   = new Date().toISOString();
  items.forEach(function(item) {
    var chat_url    = (item.chat_url    || '').trim();
    var webhook_url = (item.webhook_url || '').trim();
    if (!chat_url || !webhook_url) return;

    var space_id = '';
    var m = chat_url.match(/\/room\/([^\/\?]+)/);
    if (!m) m = chat_url.match(/\/spaces\/([^\/\?]+)/);
    if (m) space_id = m[1];

    var stock_id = 'ts-' + Date.now() + '-' + Math.random().toString(36).substring(2, 7);
    appendRow('tel_space_stock', {
      stock_id:      stock_id,
      chat_url:      chat_url,
      chat_space_id: space_id,
      webhook_url:   webhook_url,
      meet_url:      item.meet_url || '',
      status:        'available',
      assigned_to:   '',
      assigned_at:   '',
      created_at:    now,
    });
    added++;
  });

  return jsonResponse({ ok: true, added: added });
}

function handleAssignTelStock(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var mentee_id = (body.mentee_id || '').trim();
  var stock_id  = (body.stock_id  || '').trim();
  var meet_url  = (body.meet_url  || '').trim();
  if (!mentee_id) return errorResponse('MISSING_MENTEE_ID', 400);

  var users  = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var mentee = users.find(function(u){ return u.user_id === mentee_id; });
  if (!mentee) return errorResponse('USER_NOT_FOUND', 404);

  var sheet = getSheet('tel_space_stock');
  if (!sheet) return errorResponse('STOCK_SHEET_NOT_FOUND', 404);
  var stocks = sheetToObjects(sheet).filter(function(s){ return s.stock_id; });

  var target;
  if (stock_id) {
    target = stocks.find(function(s){ return s.stock_id === stock_id; });
    if (!target)                         return errorResponse('STOCK_NOT_FOUND', 404);
    if (target.status === 'assigned')    return errorResponse('ALREADY_ASSIGNED', 400);
  } else {
    target = stocks.find(function(s){ return s.status === 'available'; });
    if (!target)                         return errorResponse('NO_STOCK_AVAILABLE', 400);
  }

  var now          = new Date().toISOString();
  var finalMeetUrl = meet_url || target.meet_url || '';

  updateRowWhere('tel_space_stock', 'stock_id', target.stock_id, {
    status:      'assigned',
    assigned_to: mentee_id,
    assigned_at: now,
    meet_url:    finalMeetUrl,
  });

  updateRowWhere('users', 'user_id', mentee_id, {
    chat_url:         target.chat_url      || '',
    chat_space_id:    target.chat_space_id || '',
    chat_webhook_url: target.webhook_url   || '',
    tel_meet_url:     finalMeetUrl,
    updated_at:       now,
  });
  invalidateCache_('users');

  // Chat APIでスペース名を「TEL: メンティー名」に変更
  var renameResult = { ok: false, reason: 'no_space_id' };
  if (target.chat_space_id) {
    renameResult = renameChatSpace_(target.chat_space_id, 'TEL: ' + mentee.name);
  }

  // ── Webhookでメッセージ送信 → Chat APIで自動ピン止め（ストック割り当て = B パターン）──
  var pinResult = { ok: false, reason: 'no_webhook' };
  if (target.webhook_url && finalMeetUrl) {
    try {
      var msg = '📌 *TEL用Meet URL（固定）*\n' + finalMeetUrl
        + '\n\n毎週のTELはこのURLを使ってください。';
      var msgRes = UrlFetchApp.fetch(target.webhook_url, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ text: msg }),
        muteHttpExceptions: true
      });
      var msgCode = msgRes.getResponseCode();
      if (msgCode === 200) {
        // Webhookレスポンスからmessage.nameを取得してピン止め
        try {
          var msgData = JSON.parse(msgRes.getContentText());
          var messageName = msgData.name || ''; // 例: spaces/XXXXX/messages/YYYYY
          if (messageName && target.chat_space_id) {
            pinResult = pinChatMessage_(messageName);
          } else {
            pinResult = { ok: true, pinned: false, reason: 'message_name_not_found' };
          }
        } catch(parseErr) {
          pinResult = { ok: true, pinned: false, reason: 'parse_error: ' + parseErr.message };
        }
      } else {
        pinResult = { ok: false, reason: 'webhook_http_' + msgCode };
      }
    } catch(pinErr) {
      pinResult = { ok: false, reason: pinErr.message };
    }
  }

  Logger.log('TELスペース割り当て完了: ' + mentee.name + ' → ' + target.stock_id);
  return jsonResponse({
    ok: true, stock_id: target.stock_id,
    mentee_id: mentee_id, mentee_name: mentee.name,
    chat_url: target.chat_url || '', webhook_url: target.webhook_url || '',
    meet_url: finalMeetUrl,
    rename_result: renameResult, pin_result: pinResult,
  });
}

function handleDeleteTelStock(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);
  var stock_id = (body.stock_id || '').trim();
  if (!stock_id) return errorResponse('MISSING_STOCK_ID', 400);

  var sheet = getSheet('tel_space_stock');
  if (!sheet) return errorResponse('STOCK_SHEET_NOT_FOUND', 404);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('stock_id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === stock_id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return errorResponse('STOCK_NOT_FOUND', 404);
}

function renameChatSpace_(spaceId, newDisplayName) {
  try {
    var token = ScriptApp.getOAuthToken();
    var url   = 'https://chat.googleapis.com/v1/spaces/' + spaceId + '?updateMask=displayName';
    var res   = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ displayName: newDisplayName }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    Logger.log('スペース名変更: ' + newDisplayName + ' → HTTP ' + code);
    return code === 200 ? { ok: true } : { ok: false, reason: 'HTTP_' + code };
  } catch(e) {
    return { ok: false, reason: e.message };
  }
}

// ── Chat API: メッセージをピン止め ──
// messageName: "spaces/XXXXX/messages/YYYYY" 形式
function pinChatMessage_(messageName) {
  try {
    var token = ScriptApp.getOAuthToken();
    // Chat API: messages.patch で pinned: true を設定
    var url = 'https://chat.googleapis.com/v1/' + messageName
      + '?updateMask=pinned';
    var res = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json'
      },
      payload: JSON.stringify({ pinned: true }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    Logger.log('メッセージピン止め: ' + messageName + ' → HTTP ' + code);
    if (code === 200) {
      return { ok: true, pinned: true };
    }
    // ピン止め失敗でもメッセージ送信は成功しているので部分成功として返す
    var errBody = '';
    try { errBody = JSON.parse(res.getContentText()).error.message || ''; } catch(e2) {}
    Logger.log('ピン止め失敗詳細: ' + errBody);
    return { ok: true, pinned: false, reason: 'pin_http_' + code + ': ' + errBody };
  } catch(e) {
    Logger.log('pinChatMessage_ error: ' + e.message);
    return { ok: true, pinned: false, reason: e.message };
  }
}

// ============================================================
// Mentor: 全menteeレポート一覧
// GET api/mentor/all-reports
// ============================================================
function handleMentorAllReports(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);

  var users   = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });
  var userMap = {};
  users.forEach(function(u){ userMap[u.user_id] = u; });

  var reports = sheetToObjects(getSheet('mentor_reports')).filter(function(r){ return r.report_id; });
  var result  = reports.map(function(r) {
    var mentee  = userMap[r.mentee_id] || {};
    var mentor  = userMap[r.mentor_id] || {};
    return {
      report_id:               r.report_id,
      mentee_id:               r.mentee_id,
      mentee_name:             mentee.name || r.mentee_id || '—',
      mentor_id:               r.mentor_id,
      mentor_name:             mentor.name || r.mentor_id || '—',
      is_published:            r.is_published,
      ai_summary:              r.ai_summary              || '',
      ai_advice:               r.ai_advice               || '',
      next_goal:               r.next_goal               || '',
      next_month_project_goal: r.next_month_project_goal || '',  // ★
      next_month_study_goal:   r.next_month_study_goal   || '',  // ★
      pdf_url:                 r.pdf_url                 || '',
      created_at:              r.created_at              || '',
      published_at:            r.published_at            || '',
    };
  }).sort(function(a,b){ return (b.created_at||'').localeCompare(a.created_at||''); });

  return jsonResponse({ ok: true, reports: result });
}

// ============================================================
// Mentor: Menteeの半年目標を設定
// POST api/mentor/set-mentee-goals
// ============================================================
function handleMentorSetMenteeGoals(e) {
  var auth = requireAuth(e, 'mentor');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body = parseBody_(e);

  var mentee_id        = body.mentee_id || '';
  var goal_work_6m     = body.goal_work_6m     !== undefined ? String(body.goal_work_6m)     : null;
  var goal_skill_6m    = body.goal_skill_6m    !== undefined ? String(body.goal_skill_6m)    : null;
  var goal_start_month = body.goal_start_month !== undefined ? String(body.goal_start_month) : null;
  var goal_end_month   = body.goal_end_month   !== undefined ? String(body.goal_end_month)   : null;

  if (!mentee_id) return errorResponse('MISSING_MENTEE_ID', 400);

  // ★ キャッシュを完全にバイパスしてスプレッドシートを直接操作
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName('users');
  if (!sheet) return errorResponse('USERS_SHEET_NOT_FOUND', 500);

  // ヘッダーを直接取得
  var lastCol  = sheet.getLastColumn();
  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 必要列が存在しなければ追加
  var requiredCols = ['goal_work_6m','goal_skill_6m','goal_start_month','goal_end_month'];
  requiredCols.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(col);
      headers.push(col);
      Logger.log('users 列自動追加: ' + col + ' → 列' + lastCol);
    }
  });

  // ★ goal_start_month / goal_end_month 列を文字列型に強制（日付型変換防止）
  ['goal_start_month','goal_end_month','goal_work_6m','goal_skill_6m'].forEach(function(col) {
    var idx = headers.indexOf(col);
    if (idx >= 0) {
      sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
    }
  });

  // ★ goal_start_month / goal_end_month を YYYY-MM 形式に正規化
  // （スプレッドシートが日付型に変換した場合でも正しい文字列に戻す）
  function normalizeYearMonth(val) {
    if (!val) return '';
    var s = String(val).trim();
    // 既に YYYY-MM 形式ならそのまま
    if (/^\d{4}-\d{2}$/.test(s)) return s;
    // Date オブジェクト由来の文字列 "Wed Apr 01 2026..." など → 解析
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      var y  = d.getFullYear();
      var mo = ('0' + (d.getMonth() + 1)).slice(-2);
      return y + '-' + mo;
    }
    return s;
  }
  if (goal_start_month !== null) goal_start_month = normalizeYearMonth(goal_start_month);
  if (goal_end_month   !== null) goal_end_month   = normalizeYearMonth(goal_end_month);

  // 更新対象フィールド
  var updates = { updated_at: new Date().toISOString() };
  if (goal_work_6m     !== null) updates.goal_work_6m     = goal_work_6m;
  if (goal_skill_6m    !== null) updates.goal_skill_6m    = goal_skill_6m;
  if (goal_start_month !== null) updates.goal_start_month = goal_start_month;
  if (goal_end_month   !== null) updates.goal_end_month   = goal_end_month;

  // ★ スプレッドシートを再取得して直接書き込み（キャッシュ経由しない）
  var allData  = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var uidIdx   = headers.indexOf('user_id');
  if (uidIdx < 0) return errorResponse('USER_ID_COL_NOT_FOUND', 500);

  var updated  = false;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][uidIdx]) === String(mentee_id)) {
      // 通常フィールドはまとめて書き込み
      Object.keys(updates).forEach(function(key) {
        var idx = headers.indexOf(key);
        if (idx >= 0) allData[i][idx] = updates[key];
      });
      sheet.getRange(i + 1, 1, 1, allData[i].length).setValues([allData[i]]);

      // ★ 期間フィールドは個別セルに文字列として上書き（日付型変換を防ぐ）
      ['goal_start_month','goal_end_month'].forEach(function(col) {
        if (updates[col] === undefined) return;
        var idx = headers.indexOf(col);
        if (idx >= 0) {
          var cell = sheet.getRange(i + 1, idx + 1);
          cell.setNumberFormat('@STRING@');
          cell.setValue(String(updates[col])); // 文字列として明示保存
        }
      });

      updated = true;
      Logger.log('handleMentorSetMenteeGoals: ' + mentee_id + ' 行' + (i+1) + ' 更新完了');
      Logger.log('  goal_work_6m='     + (updates.goal_work_6m     || ''));
      Logger.log('  goal_skill_6m='    + (updates.goal_skill_6m    || ''));
      Logger.log('  goal_start_month=' + (updates.goal_start_month || ''));
      Logger.log('  goal_end_month='   + (updates.goal_end_month   || ''));
      break;
    }
  }

  if (!updated) return errorResponse('MENTEE_NOT_FOUND', 404);

  // キャッシュ無効化
  invalidateCache_('users');
  return jsonResponse({ ok: true });
}


// ============================================================
// デバッグ用: bookingsのcalendar_event_idからhtmlLinkを取得して確認
// GASエディタから手動実行 → ログでURLを確認する
// ============================================================
function debugCalendarUrl() {
  var bookings = cachedSheetToObjects_('bookings').filter(function(b){
    return b.booking_id && b.calendar_event_id;
  });
  Logger.log('calendar_event_id を持つ予約: ' + bookings.length + '件');
  bookings.slice(0, 3).forEach(function(b) {
    var eid = b.calendar_event_id;
    Logger.log('booking_id: ' + b.booking_id);
    Logger.log('  calendar_event_id: ' + eid);
    // GASのCalendarAppで実際のイベントを取得してhtmlLinkを確認
    try {
      var event = CalendarApp.getEventById(eid);
      if (event) {
        Logger.log('  getHtmlLink: ' + event.getHtmlLink());
        Logger.log('  getId: ' + event.getId());
      } else {
        Logger.log('  イベントが見つかりません（別アカウントのカレンダー？）');
      }
    } catch(e) {
      Logger.log('  取得エラー: ' + e.message);
    }
  });
}

// ============================================================
// サポートチャット機能
// ============================================================

var SUPPORT_SPACE_URL = 'https://chat.google.com/room/AAQArGxf8PQ?cls=1';
var SUPPORT_SPACE_ID  = 'AAQArGxf8PQ';
var SUPPORT_SPACE_NAME = 'spaces/AAQArGxf8PQ';  // Chat API用スペース名
var SUPPORT_ADMIN_EMAIL = 'test.admin@socialshift.work'; // 集計除外メール

// support_id 採番（SP-001形式）
function generateSupportId_() {
  var sheet = getOrCreateSheet_('support_raw');
  var last  = sheet.getLastRow();
  if (last <= 1) return 'SP-001';
  var data  = sheet.getRange(2, 1, last - 1, 1).getValues();
  var nums  = data.map(function(r){ return parseInt((r[0]||'').replace('SP-',''))||0; });
  var max   = Math.max.apply(null, nums);
  return 'SP-' + String(max + 1).padStart(3, '0');
}

// シート取得 or 作成
function getOrCreateSheet_(name) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  // ヘッダー追加
  var headers = {
    support_raw:    ['support_id','received_at','sender_name','sender_email','message','message_id'],
    support_status: ['support_id','message_excerpt','sender_email','status','response','responded_at','notified'],
    support_history:['log_id','support_id','changed_at','changed_by','field','old_value','new_value'],
  };
  if (headers[name]) sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
  return sheet;
}

// ============================================================
// ★ サポートチャット ポーリング関数
// GASのトリガーで10分毎に実行する
// ============================================================
function pollSupportChat() {
  Logger.log('=== pollSupportChat 開始 ===');
  try {
    var token = ScriptApp.getOAuthToken();

    // Chat API: 直近24時間のメッセージを取得
    var since = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
    var url = 'https://chat.googleapis.com/v1/' + SUPPORT_SPACE_NAME + '/messages'
      + '?orderBy=createTime+asc'
      + '&filter=createTime+%3E+%22' + since.replace(/:/g,'%3A') + '%22'
      + '&pageSize=100';

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('Chat API エラー: ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 200));
      return;
    }

    var data = JSON.parse(resp.getContentText());
    var messages = data.messages || [];
    Logger.log('取得メッセージ数: ' + messages.length);

    if (messages.length === 0) { Logger.log('新着なし'); return; }

    // 既存の message_id セットを取得（重複防止）
    var rawSheet = getOrCreateSheet_('support_raw');
    var existing = sheetToObjects(rawSheet);
    var existingIds = {};
    existing.forEach(function(r){ if (r.message_id) existingIds[r.message_id] = true; });

    var added = 0;
    messages.forEach(function(msg) {
      var msgId     = msg.name || '';           // "spaces/xxx/messages/yyy"
      var sender    = msg.sender || {};
      var email     = (sender.email || '').toLowerCase();
      var name      = sender.displayName || '';
      var text      = msg.text || msg.argumentText || '';
      var createdAt = msg.createTime || new Date().toISOString();
      var senderType = sender.type || '';

      // 除外条件
      if (!msgId)                              return; // IDなし
      if (existingIds[msgId])                  return; // 既に処理済み
      if (!text.trim())                        return; // 空メッセージ
      if (senderType === 'BOT')                return; // Botのメッセージ
      if (email === SUPPORT_ADMIN_EMAIL.toLowerCase()) return; // 管理者

      // スプシに書き込み
      var supportId = generateSupportId_();
      var now       = new Date().toISOString();
      var excerpt   = text.trim().slice(0, 50) + (text.trim().length > 50 ? '…' : '');

      appendRow('support_raw', {
        support_id:   supportId,
        received_at:  createdAt,
        sender_name:  name,
        sender_email: email,
        message:      text.trim(),
        message_id:   msgId,
      });

      appendRow('support_status', {
        support_id:      supportId,
        message_excerpt: excerpt,
        sender_email:    email,
        status:          '未対応',
        response:        '',
        responded_at:    '',
        notified:        'FALSE',
      });

      existingIds[msgId] = true; // 同一ポーリング内での重複防止
      added++;
      Logger.log('追加: ' + supportId + ' from ' + email + ' msg=' + text.slice(0, 30));
    });

    if (added > 0) {
      invalidateCache_('support_raw');
      invalidateCache_('support_status');
    }
    Logger.log('=== pollSupportChat 完了: ' + added + '件追加 ===');

  } catch(err) {
    Logger.log('pollSupportChat エラー: ' + err.message);
  }
}

// ============================================================
// ポーリングトリガーのセットアップ（初回のみ手動実行）
// GASエディタで setupSupportPollingTrigger() を1回実行する
// ============================================================
function setupSupportPollingTrigger() {
  // 既存のpollSupportChatトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pollSupportChat') {
      ScriptApp.deleteTrigger(t);
      Logger.log('既存トリガー削除: pollSupportChat');
    }
  });
  // 10分毎のトリガーを作成
  ScriptApp.newTrigger('pollSupportChat')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('✅ pollSupportChat トリガー設定完了（10分毎）');
}

// doPost の Chat Webhook ルーティングは廃止（ポーリング方式に移行）
// 旧: handleSupportWebhook は削除


// ── 一覧取得 ──
// GET api/support/list
function handleSupportList(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var userEmail = auth.payload.email || '';

  var status  = sheetToObjects(getOrCreateSheet_('support_status')).filter(function(r){ return r.support_id; });
  var raw     = sheetToObjects(getOrCreateSheet_('support_raw')).filter(function(r){ return r.support_id; });
  var rawMap  = {};
  raw.forEach(function(r){ rawMap[r.support_id] = r; });

  var list = status
    .sort(function(a, b){ return (b.support_id||'').localeCompare(a.support_id||''); })
    .map(function(s) {
      var r        = rawMap[s.support_id] || {};
      var isMine   = (r.sender_email || '').toLowerCase() === userEmail.toLowerCase();
      return {
        support_id:       s.support_id,
        received_at:      r.received_at     || '',
        message_excerpt:  s.message_excerpt || '',
        status:           s.status          || '未対応',
        response:         s.response        || '',
        responded_at:     s.responded_at    || '',
        is_mine:          isMine,            // ★ 自分の投稿フラグ
      };
    });

  return jsonResponse({ ok: true, list: list });
}

// ── 詳細＋履歴取得 ──
// GET api/support/detail  (support_id は body.support_id または e.parameter.support_id)
function handleSupportDetail(e) {
  var auth = requireAuth(e);
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var supportId = body.support_id
    || (e.parameter && e.parameter.support_id)
    || '';
  if (!supportId) return errorResponse('MISSING_SUPPORT_ID', 400);
  var userEmail = auth.payload.email || '';

  var status  = sheetToObjects(getOrCreateSheet_('support_status')).find(function(r){ return r.support_id === supportId; });
  var raw     = sheetToObjects(getOrCreateSheet_('support_raw')).find(function(r){ return r.support_id === supportId; });
  var history = sheetToObjects(getOrCreateSheet_('support_history'))
    .filter(function(r){ return r.support_id === supportId; })
    .sort(function(a, b){ return (a.changed_at||'').localeCompare(b.changed_at||''); });

  if (!status) return errorResponse('NOT_FOUND', 404);
  var isMine = (raw && (raw.sender_email||'').toLowerCase() === userEmail.toLowerCase());
  var isAdmin = auth.payload.role === 'admin';

  return jsonResponse({
    ok:          true,
    support_id:  supportId,
    received_at: raw ? (raw.received_at||'') : '',
    sender_email: isAdmin ? (raw ? (raw.sender_email||'') : '') : '', // adminのみメール返す
    message:     (isMine || isAdmin) ? (raw ? (raw.message||'') : '') : '',  // 自分またはadminのみ全文
    message_excerpt: status.message_excerpt || '',
    status:      status.status    || '未対応',
    response:    status.response  || '',
    responded_at:status.responded_at || '',
    is_mine:     isMine,
    history:     history.map(function(h){ return {
      changed_at: h.changed_at,
      field:      h.field,
      old_value:  h.old_value,
      new_value:  h.new_value,
    }; }),
  });
}

// ── ステータス・対応内容更新（admin） ──
// POST api/admin/support/update { support_id, status?, response?, responded_at? }
function handleSupportUpdate(e) {
  var auth = requireAuth(e, 'admin');
  if (auth.error) return errorResponse(auth.error, auth.status);
  var body      = parseBody_(e);
  var supportId = (body.support_id || '').trim();
  if (!supportId) return errorResponse('MISSING_SUPPORT_ID', 400);

  var statusSheet  = getOrCreateSheet_('support_status');
  var historySheet = getOrCreateSheet_('support_history');
  var rows         = sheetToObjects(statusSheet);
  var current      = rows.find(function(r){ return r.support_id === supportId; });
  if (!current) return errorResponse('NOT_FOUND', 404);

  var now      = new Date().toISOString();
  var changedBy= auth.payload.email || 'admin';
  var updates  = { updated_at: now };
  var changed  = [];

  if (body.status !== undefined && body.status !== current.status) {
    changed.push({ field: 'status', old: current.status, new: body.status });
    updates.status = body.status;
  }
  if (body.response !== undefined && body.response !== current.response) {
    changed.push({ field: 'response', old: current.response||'', new: body.response });
    updates.response = body.response;
  }
  if (updates.status || updates.response) {
    updates.responded_at = now;
    updates.notified     = 'FALSE'; // ★ 通知フラグをリセット（再通知対象に）
    updateRowWhere('support_status', 'support_id', supportId, updates);
    // 変更履歴を記録
    changed.forEach(function(c) {
      var logId = 'hl-' + Date.now() + '-' + Math.random().toString(36).substring(2, 6);
      appendRow('support_history', {
        log_id:     logId,
        support_id: supportId,
        changed_at: now,
        changed_by: changedBy,
        field:      c.field,
        old_value:  c.old,
        new_value:  c.new,
      });
    });
    invalidateCache_('support_status');
  }

  return jsonResponse({ ok: true, support_id: supportId, changed: changed.length });
}

// ── チャット通知送信（ポーリング or 手動実行） ──
// 10分毎のトリガーに追加するか、GASエディタから手動実行
function sendSupportNotifications() {
  var WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SUPPORT_SPACE_WEBHOOK');
  if (!WEBHOOK_URL) { Logger.log('SUPPORT_SPACE_WEBHOOK が未設定です'); return; }

  var status = sheetToObjects(getOrCreateSheet_('support_status'))
    .filter(function(r){
      return r.support_id
        && (r.status === '対応中' || r.status === '解決済み')
        && String(r.notified || '').toUpperCase() !== 'TRUE'
        && r.response;
    });

  if (status.length === 0) { Logger.log('通知対象なし'); return; }

  var raw = sheetToObjects(getOrCreateSheet_('support_raw'));
  var rawMap = {};
  raw.forEach(function(r){ rawMap[r.support_id] = r; });

  var users = cachedSheetToObjects_('users').filter(function(u){ return u.user_id; });

  status.forEach(function(s) {
    var r       = rawMap[s.support_id];
    if (!r) return;
    var email   = r.sender_email || '';
    // Google Chat のメンション形式：<users/メールアドレス>
    var mention = '<users/' + email + '>';
    var statusEmoji = s.status === '解決済み' ? '✅' : '⏳';
    var msg = mention + ' 【' + s.support_id + ' 対応状況更新】\n'
      + statusEmoji + ' ステータス：' + s.status + '\n'
      + '📝 対応内容：' + s.response + '\n'
      + '対応状況の詳細は https://socialshift-svg.github.io/potal/mentee.html のサポートページからご確認ください。';

    try {
      UrlFetchApp.fetch(WEBHOOK_URL, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ text: msg }),
        muteHttpExceptions: true,
      });
      updateRowWhere('support_status', 'support_id', s.support_id, { notified: 'TRUE' });
      Logger.log('通知送信: ' + s.support_id + ' → ' + email);
    } catch(err) {
      Logger.log('通知送信エラー: ' + s.support_id + ' ' + err.message);
    }
  });
  invalidateCache_('support_status');
}
