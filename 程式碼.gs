/**
 * 陳地瓜走到哪 — 徒步環島打卡系統後端 (Google Apps Script)
 * BUILD: R1.1
 *
 * 部署方式：部署 → 新增部署作業 → 網頁應用程式
 *   執行身分：我
 *   誰可以存取：所有人
 * 前端是獨立的靜態網頁（GitHub Pages / Cloudflare Pages），
 * 本檔只當 API：寫試算表、存 Drive 照片、寄信。
 */

const BUILD = 'R1.1';   // ← 改版時三處同步：本行、每日摘要信頁尾、前端 web/*.html 的 BUILD

// ────────────────────────────────────────────────
// 設定：第一次部署前把這裡改完
// ────────────────────────────────────────────────
const CONFIG = {
  SHEET_ID: '',                              // 留空＝用綁定的試算表；獨立指令碼請填試算表 ID
  PHOTO_FOLDER_ID: '',                       // 留空＝自動在雲端硬碟建「環島照片」資料夾
  CHECKIN_TOKEN: '換成一組長亂碼_例如k7Qm2xR9vLp4',   // 只有朋友本人知道，放在打卡頁網址列
  SITE_URL: 'https://your-name.github.io/digua',   // 前端網址，結尾不要斜線
  WALKER_NAME: '陳地瓜',
  SITE_NAME: '陳地瓜走到哪',                  // 寄件者名稱與信件抬頭
  START_DATE: '2026-09-15',                  // 出發日，用來算「第幾天」
  SUMMARY_HOUR: 20,                          // 每日摘要寄出時間（整點）
  TIMEZONE: 'Asia/Taipei'
};

// 打卡類型 → 是否即時寄信
const TYPES = {
  '出發':   { instant: true,  color: '#0B6E4F' },
  '休息':   { instant: false, color: '#C99A1E' },
  '收工':   { instant: true,  color: '#1F2622' },
  '報平安': { instant: true,  color: '#B4472E' }
};

const SHEET_CHECKIN = '打卡';
const SHEET_SUB = '訂閱';
const HEAD_CHECKIN = ['打卡ID', '時間', '類型', '緯度', '經度', '誤差(m)', '電量(%)',
                      '留言', '地址', '照片ID', '照片網址', '距上點(km)'];
const HEAD_SUB = ['訂閱時間', 'Email', '稱呼', '通知等級', '退訂Token', '狀態'];

// ────────────────────────────────────────────────
// 路由
// ────────────────────────────────────────────────

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.api === 'feed') {
      out = { ok: true, build: BUILD, data: buildFeed_() };
    } else {
      out = { ok: true, build: BUILD, msg: '陳地瓜走到哪 API' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  // 支援 JSONP，萬一瀏覽器擋 CORS 時前端可切換
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(out);
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '格式錯誤' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    switch (body.action) {
      case 'checkin':   return json_(handleCheckin_(body));
      case 'subscribe': return json_(handleSubscribe_(body));
      case 'unsub':     return json_(handleUnsub_(body));
      default:          return json_({ ok: false, error: '未知的 action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────
// 打卡
// ────────────────────────────────────────────────

function handleCheckin_(b) {
  if (b.token !== CONFIG.CHECKIN_TOKEN) return { ok: false, error: '通行碼不符' };
  if (!TYPES[b.type]) return { ok: false, error: '類型不存在' };
  if (typeof b.lat !== 'number' || typeof b.lng !== 'number') return { ok: false, error: '缺少座標' };

  const sh = sheet_(SHEET_CHECKIN, HEAD_CHECKIN);
  const cid = b.cid || Utilities.getUuid();

  // 冪等：離線補送可能重複，用 cid 擋掉
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(Math.max(2, last - 299), 1, Math.min(300, last - 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === cid) return { ok: true, dup: true, cid: cid };
    }
  }

  const when = b.ts ? new Date(b.ts) : new Date();
  const addr = reverseGeocode_(b.lat, b.lng);

  let photoId = '', photoUrl = '';
  if (b.photo) {
    const saved = savePhoto_(b.photo, when, b.type);
    photoId = saved.id;
    photoUrl = saved.url;
  }

  const prev = lastPoint_(sh);
  const seg = prev ? round_(haversine_(prev.lat, prev.lng, b.lat, b.lng), 2) : 0;

  sh.appendRow([cid, when, b.type, b.lat, b.lng,
                b.acc == null ? '' : Math.round(b.acc),
                b.battery == null ? '' : b.battery,
                b.note || '', addr, photoId, photoUrl, seg]);

  try { CacheService.getScriptCache().remove('feed'); } catch (err) {}

  if (TYPES[b.type].instant) {
    try { sendInstant_(b.type, when, addr, b.note || '', photoUrl); }
    catch (err) { console.error('寄信失敗：' + err); }
  }

  return { ok: true, cid: cid, addr: addr, seg: seg };
}

function lastPoint_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const r = sh.getRange(last, 4, 1, 2).getValues()[0];
  if (typeof r[0] !== 'number') return null;
  return { lat: r[0], lng: r[1] };
}

function savePhoto_(dataUrl, when, type) {
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+);base64,(.+)$/i);
  if (!m) return { id: '', url: '' };
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1],
    Utilities.formatDate(when, CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss') + '_' + type + '.jpg');
  const file = photoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    id: file.getId(),
    url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600'
  };
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  let id = CONFIG.PHOTO_FOLDER_ID || props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (err) {}
  }
  const folder = DriveApp.createFolder('陳地瓜走到哪 照片');
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function reverseGeocode_(lat, lng) {
  try {
    const res = Maps.newGeocoder().setLanguage('zh-TW').reverseGeocode(lat, lng);
    if (res.status === 'OK' && res.results.length) {
      return String(res.results[0].formatted_address).replace(/^\d{3,5}\s*/, '');
    }
  } catch (err) {
    console.warn('反向地理編碼失敗：' + err);
  }
  return '';
}

// ────────────────────────────────────────────────
// 訂閱
// ────────────────────────────────────────────────

function handleSubscribe_(b) {
  const email = String(b.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Email 格式不對' };

  const level = (b.level === '每日') ? '每日' : '即時';
  const sh = sheet_(SHEET_SUB, HEAD_SUB);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === email) {
      sh.getRange(i + 1, 3, 1, 4).setValues([[b.name || rows[i][2], level, rows[i][4] || Utilities.getUuid(), '有效']]);
      return { ok: true, updated: true };
    }
  }

  const token = Utilities.getUuid();
  sh.appendRow([new Date(), email, b.name || '', level, token, '有效']);
  try { sendWelcome_(email, b.name || '', token); } catch (err) { console.error(err); }
  return { ok: true };
}

function handleUnsub_(b) {
  const sh = sheet_(SHEET_SUB, HEAD_SUB);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]) === String(b.token)) {
      sh.getRange(i + 1, 6).setValue('已退訂');
      return { ok: true };
    }
  }
  return { ok: false, error: '找不到這個訂閱' };
}

function subscribers_(level) {
  const sh = sheet_(SHEET_SUB, HEAD_SUB);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][5] !== '有效') continue;
    if (level === '即時' && rows[i][3] !== '即時') continue;
    out.push({ email: rows[i][1], name: rows[i][2], token: rows[i][4] });
  }
  return out;
}

// ────────────────────────────────────────────────
// 寄信
// ────────────────────────────────────────────────

function sendInstant_(type, when, addr, note, photoUrl) {
  const list = subscribers_('即時');
  if (!list.length) return;
  if (MailApp.getRemainingDailyQuota() < list.length) {
    console.warn('寄信額度不足，跳過即時通知');
    return;
  }

  const t = Utilities.formatDate(when, CONFIG.TIMEZONE, 'M月d日 HH:mm');
  const subject = '［第 ' + dayNo_(when) + ' 天］' + CONFIG.WALKER_NAME + type;

  list.forEach(function (s) {
    const html = mailShell_(
      type,
      '<p style="margin:0 0 4px;font-size:15px;color:#5b5f59">' + t + '</p>' +
      '<p style="margin:0 0 16px;font-size:20px;font-weight:700;color:#1F2622">' + esc_(addr || '（定位中）') + '</p>' +
      (note ? '<p style="margin:0 0 16px;font-size:16px;line-height:1.7">' + esc_(note) + '</p>' : '') +
      (photoUrl ? '<img src="' + photoUrl + '" style="width:100%;border-radius:6px;margin-bottom:16px">' : ''),
      s.token);
    MailApp.sendEmail({ to: s.email, subject: subject, htmlBody: html, name: CONFIG.SITE_NAME });
  });
}

function sendWelcome_(email, name, token) {
  MailApp.sendEmail({
    to: email,
    subject: '已訂閱「' + CONFIG.SITE_NAME + '」',
    name: CONFIG.SITE_NAME,
    htmlBody: mailShell_('訂閱成功',
      '<p style="margin:0 0 16px;font-size:16px;line-height:1.7">' +
      esc_(name || '嗨') + '，之後' + esc_(CONFIG.WALKER_NAME) +
      '出發、收工、報平安時你會收到信，每天晚上也有一封當日摘要。</p>', token)
  });
}

/** 每日摘要：用時間驅動觸發器每天跑一次 */
function dailySummary() {
  const rows = todayRows_();
  const list = subscribers_();
  if (!list.length) return;

  const today = new Date();
  let body, subject;

  if (!rows.length) {
    subject = '［第 ' + dayNo_(today) + ' 天］地瓜今天沒有打卡';
    body = '<p style="margin:0 0 16px;font-size:16px;line-height:1.7">今天沒有收到' +
           esc_(CONFIG.WALKER_NAME) + '的打卡。可能在沒訊號的路段，也可能是休息日。</p>';
  } else {
    let km = 0;
    rows.forEach(function (r) { km += Number(r[11]) || 0; });
    const last = rows[rows.length - 1];

    subject = '［第 ' + dayNo_(today) + ' 天］' + esc_(last[8] || '今日摘要') +
              '，直線 ' + round_(km, 1) + ' 公里';

    body = '<table style="width:100%;border-collapse:collapse;margin:0 0 20px">' +
           '<tr>' +
           statCell_('第 ' + dayNo_(today) + ' 天', '出發至今') +
           statCell_(round_(km, 1) + ' km', '今日直線距離') +
           statCell_(rows.length + ' 次', '今日打卡') +
           '</tr></table>';

    body += rows.map(function (r) {
      const time = Utilities.formatDate(new Date(r[1]), CONFIG.TIMEZONE, 'HH:mm');
      const c = (TYPES[r[2]] || {}).color || '#1F2622';
      return '<div style="border-left:3px solid ' + c + ';padding:2px 0 2px 14px;margin:0 0 18px">' +
             '<div style="font-size:13px;color:#5b5f59">' + time + '　' + esc_(r[2]) + '</div>' +
             '<div style="font-size:16px;font-weight:600;color:#1F2622">' + esc_(r[8] || '') + '</div>' +
             (r[7] ? '<div style="font-size:15px;line-height:1.7;margin-top:6px">' + esc_(r[7]) + '</div>' : '') +
             (r[10] ? '<img src="' + r[10] + '" style="width:100%;border-radius:6px;margin-top:10px">' : '') +
             '</div>';
    }).join('');
  }

  list.forEach(function (s) {
    if (MailApp.getRemainingDailyQuota() < 1) return;
    MailApp.sendEmail({
      to: s.email, subject: subject, name: CONFIG.SITE_NAME,
      htmlBody: mailShell_('今日摘要', body, s.token)
    });
  });
}

function statCell_(big, small) {
  return '<td style="padding:12px 14px;background:#F7F5EF;border-radius:6px;width:33%">' +
         '<div style="font-size:22px;font-weight:800;color:#0B6E4F">' + big + '</div>' +
         '<div style="font-size:12px;color:#5b5f59;margin-top:2px">' + small + '</div></td>';
}

function mailShell_(title, inner, token) {
  return '<div style="font-family:-apple-system,\'PingFang TC\',\'Noto Sans TC\',sans-serif;' +
         'max-width:560px;margin:0 auto;padding:24px;color:#1F2622">' +
         '<div style="font-size:13px;letter-spacing:.08em;color:#0B6E4F;font-weight:700;margin-bottom:14px">' +
         esc_(CONFIG.SITE_NAME) + '　·　' + esc_(title) + '</div>' +
         inner +
         '<p style="margin:24px 0 0"><a href="' + CONFIG.SITE_URL +
         '" style="display:inline-block;background:#0B6E4F;color:#fff;text-decoration:none;' +
         'padding:12px 20px;border-radius:6px;font-weight:600">看地圖</a></p>' +
         '<hr style="border:0;border-top:1px solid #E3E0D6;margin:28px 0 12px">' +
         '<p style="font-size:12px;color:#8a8d86;margin:0">' +
         '<a href="' + CONFIG.SITE_URL + '/subscribe.html?unsub=' + token +
         '" style="color:#8a8d86">不想再收信</a>　·　' + BUILD + '</p></div>';
}

// ────────────────────────────────────────────────
// 公開資料
// ────────────────────────────────────────────────

function buildFeed_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('feed');
  if (hit) { try { return JSON.parse(hit); } catch (err) {} }
  const out = readFeed_();
  // 公開頁每分鐘輪詢，快取擋掉重複讀表；超過 100KB 就放棄快取
  try { cache.put('feed', JSON.stringify(out), 45); } catch (err) {}
  return out;
}

function readFeed_() {
  const sh = sheet_(SHEET_CHECKIN, HEAD_CHECKIN);
  const last = sh.getLastRow();
  const points = [];
  let total = 0;

  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, HEAD_CHECKIN.length).getValues();
    rows.forEach(function (r) {
      if (typeof r[3] !== 'number') return;
      total += Number(r[11]) || 0;
      points.push({
        t: new Date(r[1]).toISOString(),
        type: r[2], lat: r[3], lng: r[4],
        acc: r[5] || null, note: r[7] || '', addr: r[8] || '',
        photo: r[10] || '', km: Number(r[11]) || 0
      });
    });
  }

  return {
    walker: CONFIG.WALKER_NAME,
    site: CONFIG.SITE_NAME,
    startDate: CONFIG.START_DATE,
    day: dayNo_(new Date()),
    totalKm: round_(total, 1),
    updated: new Date().toISOString(),
    points: points
  };
}

// ────────────────────────────────────────────────
// 工具
// ────────────────────────────────────────────────

function ss_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
                         : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, head) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

function todayRows_() {
  const sh = sheet_(SHEET_CHECKIN, HEAD_CHECKIN);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return sh.getRange(2, 1, last - 1, HEAD_CHECKIN.length).getValues().filter(function (r) {
    if (!r[1]) return false;
    return Utilities.formatDate(new Date(r[1]), CONFIG.TIMEZONE, 'yyyy-MM-dd') === today;
  });
}

function dayNo_(d) {
  const start = new Date(CONFIG.START_DATE + 'T00:00:00+08:00');
  return Math.max(1, Math.floor((d - start) / 86400000) + 1);
}

/** 兩點直線距離（公里）。實際步行距離會比這個大，信裡已標明「直線」 */
function haversine_(lat1, lon1, lat2, lon2) {
  const R = 6371, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * rad) * Math.cos(lat2 * rad) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function round_(n, d) { const p = Math.pow(10, d); return Math.round(n * p) / p; }

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

// ────────────────────────────────────────────────
// 一次性安裝：在編輯器裡手動執行這個函式
// ────────────────────────────────────────────────

function setup() {
  sheet_(SHEET_CHECKIN, HEAD_CHECKIN);
  sheet_(SHEET_SUB, HEAD_SUB);
  photoFolder_();

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailySummary').timeBased()
    .atHour(CONFIG.SUMMARY_HOUR).nearMinute(0).everyDays(1).create();

  console.log('安裝完成 ' + BUILD + '，每日摘要 ' + CONFIG.SUMMARY_HOUR + ' 點寄出');
}

/** 測試用：不寫入資料，只確認寄信通道正常 */
function testMail() {
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: '「' + CONFIG.SITE_NAME + '」系統測試 ' + BUILD,
    htmlBody: mailShell_('測試', '<p>看到這封信就表示寄信正常。</p>', 'TEST')
  });
}
