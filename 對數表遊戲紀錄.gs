/**
 * 納皮爾的對數表 · 學生遊戲紀錄（Google Apps Script）
 * 設計：余佳霖老師 · Valerie Yu Studio
 *
 * 使用步驟
 * 1. 新建一個 Google 試算表 → 擴充功能 → Apps Script，刪掉原本的內容，貼上這整個檔案，存檔。
 * 2. 在上方函式選單選「setup」→ 執行，依畫面完成授權（只需要一次）。
 *    試算表會出現三個分頁：學生總覽、作答紀錄、題目統計。
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    執行身分：我；誰可以存取：所有人 → 部署 → 複製「網頁應用程式網址」（結尾是 /exec）。
 * 4. 把網址貼進遊戲 HTML 檔開頭的 const LOG_URL='…'; 再上傳到 GitHub Pages。
 *
 * 之後如果修改了這份程式碼，要用「部署 → 管理部署作業 → 編輯 → 版本：新版本」更新，網址才會維持不變。
 * 想看「哪一題最多人卡關」，在試算表上方選單按「對數表遊戲 → 更新題目統計」。
 */

const SHEET_SUM = '學生總覽';
const SHEET_LOG = '作答紀錄';
const SHEET_STAT = '題目統計';

const SKILLS = ['同底數相乘', '同底數相除', '次方的次方', '分數指數與根號', '指數函數圖形', '指數式與對數式互換',
  '求對數值', '積的對數', '商的對數', '冪的對數', '換底公式', '首數與科學記號', '查對數表', '反查對數表',
  '用對數算乘除', '用對數算次方與開根號'];
const ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII'];

const LOG_HEAD = ['收到時間', '學生時間', '班級', '座號', '姓名', '事件', '模式', '章節', '房間', '題號', '技能',
  '題目', '學生作答（依序）', '正確答案', '第一次就答對', '最後答對', '看解答', '作答次數', '買提示', '作答秒數', '備註'];
const SUM_HEAD = ['班級', '座號', '姓名', '最後上線', '目前進度', '等級', '經驗值', '銀幣', '完成章節', '總星數']
  .concat(ROMAN.map(r => '第 ' + r + ' 章'))
  .concat(['委託完成', '快算最高分', '成就數'])
  .concat(SKILLS);
const SKILL_COL0 = 10 + 8 + 3 + 1; // 第一個技能欄（1 起算）

const EVENT_NAME = {
  login: '登入', answer: '作答', chapter_start: '進入房間', chapter_clear: '逃脫成功',
  rush: '風暴快算', badge: '獲得成就', levelup: '升級', reset: '清除進度'
};

/* ---------- 接收遊戲傳來的紀錄 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets_(ss);

    const events = (data.events || []).filter(x => x && x.id);
    const cache = CacheService.getScriptCache();
    const seen = events.length ? cache.getAll(events.map(x => String(x.id))) : {};
    const now = new Date();
    const rows = [];
    events.forEach(x => {
      if (seen[x.id]) return; // 同一筆重複送達時略過
      const st = x.st || {};
      rows.push([
        now, x.t ? new Date(x.t) : '', st.cls || '', st.seat || '', st.name || '',
        EVENT_NAME[x.type] || x.type || '', x.mode || '', x.ch ? '第 ' + ROMAN[x.ch - 1] + ' 章' : '',
        x.room || '', x.q || '', x.skillName || '', x.question || '', x.given || '', x.correct || '',
        yn_(x.first), yn_(x.ok), yn_(x.revealed), x.tries === undefined ? '' : x.tries,
        yn_(x.hint), x.sec === undefined ? '' : x.sec, x.note || ''
      ].map(safe_));
    });
    if (rows.length) {
      const sh = ss.getSheetByName(SHEET_LOG);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_HEAD.length).setValues(rows);
      const put = {};
      events.forEach(x => { put[String(x.id)] = '1'; });
      cache.putAll(put, 21600);
    }
    if (data.summary && data.summary.st) upsertStudent_(ss, data.summary);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

/* 用瀏覽器打開網址時，確認服務有在運作 */
function doGet() {
  return ContentService.createTextOutput('納皮爾的對數表 · 紀錄服務運作中');
}

/* ---------- 學生總覽：每位學生一列，隨時更新 ---------- */
function upsertStudent_(ss, s) {
  const sh = ss.getSheetByName(SHEET_SUM);
  const st = s.st;
  const key = String(st.cls) + '|' + String(st.seat);
  const last = sh.getLastRow();
  const when = new Date(s.ts || Date.now());
  let r = -1;
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 4).getValues();
    for (let i = 0; i < v.length; i++) {
      if (String(v[i][0]) + '|' + String(v[i][1]) === key) {
        r = i + 2;
        if (v[i][3] instanceof Date && v[i][3] > when) return; // 較舊的資料晚到，不覆蓋
        break;
      }
    }
  }
  const stars = (s.stars || []).slice(0, 8);
  while (stars.length < 8) stars.push(0);
  const skills = (s.skills || []).slice(0, 16).map(k =>
    k && k.n ? ['未接觸', '銅', '銀', '金'][k.lv] + ' ' + k.c + '/' + k.n : '');
  while (skills.length < 16) skills.push('');
  const row = [st.cls, st.seat, st.name, when, s.cur || '', s.lv || 1, s.xp || 0, s.coins || 0,
    s.done || 0, stars.reduce((a, b) => a + b, 0)]
    .concat(stars.map(x => x ? '★'.repeat(x) : ''))
    .concat([s.dn || 0, s.rush || 0, s.badges || 0])
    .concat(skills)
    .map(safe_);
  if (r < 0) r = last + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

/* ---------- 題目統計：找出最多人卡關的題目與技能 ---------- */
function updateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets_(ss);
  const log = ss.getSheetByName(SHEET_LOG);
  const n = log.getLastRow() - 1;
  const out = ss.getSheetByName(SHEET_STAT);
  out.clear();
  if (n < 1) { out.getRange(1, 1).setValue('還沒有作答紀錄。'); return; }
  const v = log.getRange(2, 1, n, LOG_HEAD.length).getValues();
  const C = {};
  LOG_HEAD.forEach((h, i) => { C[h] = i; });

  const byQ = {}, bySkill = {};
  v.forEach(r => {
    if (r[C['事件']] !== '作答') return;
    const first = r[C['第一次就答對']] === '是';
    const tries = Number(r[C['作答次數']]) || 0;
    const sec = Number(r[C['作答秒數']]) || 0;
    const reveal = r[C['看解答']] === '是';
    const hint = r[C['買提示']] === '是';
    const add = (map, key, base) => {
      const o = map[key] || (map[key] = Object.assign({ n: 0, first: 0, tries: 0, sec: 0, reveal: 0, hint: 0 }, base));
      o.n++; if (first) o.first++; o.tries += tries; o.sec += sec; if (reveal) o.reveal++; if (hint) o.hint++;
    };
    const skill = r[C['技能']];
    if (skill && skill !== '暖身') add(bySkill, skill, { skill });
    if (r[C['模式']] === '劇情' && r[C['題號']]) {
      add(byQ, r[C['題號']], { q: r[C['題號']], ch: r[C['章節']], room: r[C['房間']], skill, text: r[C['題目']] });
    }
  });

  const pct = (a, b) => b ? a / b : '';
  const qRows = Object.values(byQ)
    .map(o => [o.q, o.ch, o.room, o.skill, o.text, o.n, pct(o.first, o.n), o.tries / o.n, o.sec / o.n, o.reveal, o.hint])
    .sort((a, b) => a[6] - b[6]);
  const sRows = SKILLS.filter(k => bySkill[k])
    .map(k => { const o = bySkill[k]; return [k, o.n, pct(o.first, o.n), o.tries / o.n, o.sec / o.n, o.reveal, o.hint]; })
    .sort((a, b) => a[2] - b[2]);

  out.getRange(1, 1).setValue('劇情題目（第一次就答對比例由低到高；越上面越多人卡關）').setFontWeight('bold');
  const qHead = ['題號', '章節', '房間', '技能', '題目', '作答人次', '第一次就答對', '平均作答次數', '平均秒數', '看解答', '買提示'];
  out.getRange(2, 1, 1, qHead.length).setValues([qHead]).setFontWeight('bold').setBackground('#efe3c4');
  if (qRows.length) {
    out.getRange(3, 1, qRows.length, qHead.length).setValues(qRows);
    out.getRange(3, 7, qRows.length, 1).setNumberFormat('0%');
    out.getRange(3, 8, qRows.length, 2).setNumberFormat('0.0');
  }
  const top = 4 + qRows.length + 1;
  out.getRange(top, 1).setValue('技能（含劇情、港口委託、風暴快算）').setFontWeight('bold');
  const sHead = ['技能', '作答人次', '第一次就答對', '平均作答次數', '平均秒數', '看解答', '買提示'];
  out.getRange(top + 1, 1, 1, sHead.length).setValues([sHead]).setFontWeight('bold').setBackground('#efe3c4');
  if (sRows.length) {
    out.getRange(top + 2, 1, sRows.length, sHead.length).setValues(sRows);
    out.getRange(top + 2, 3, sRows.length, 1).setNumberFormat('0%');
    out.getRange(top + 2, 4, sRows.length, 2).setNumberFormat('0.0');
  }
  out.setColumnWidth(5, 380);
  out.getRange(3, 5, Math.max(1, qRows.length), 1).setWrap(true);
  out.getRange(1, 13).setValue('更新時間：' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm'));
}

/* ---------- 第一次使用：建立分頁 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  ensureSheets_(ss);
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 3 && def.getLastRow() === 0) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('設定完成！接下來請「部署 → 新增部署作業 → 網頁應用程式」，存取權選「所有人」。');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('對數表遊戲')
    .addItem('更新題目統計', 'updateStats')
    .addItem('初始化分頁', 'setup')
    .addToUi();
}

function ensureSheets_(ss) {
  let sum = ss.getSheetByName(SHEET_SUM);
  if (!sum) {
    sum = ss.insertSheet(SHEET_SUM, 0);
    sum.getRange(1, 1, 1, SUM_HEAD.length).setValues([SUM_HEAD]).setFontWeight('bold').setBackground('#efe3c4').setWrap(true);
    sum.setFrozenRows(1); sum.setFrozenColumns(3);
    sum.getRange('A:B').setNumberFormat('@');
    sum.getRange('D:D').setNumberFormat('yyyy/MM/dd HH:mm');
    sum.setColumnWidth(5, 220);
    const skillRange = sum.getRange(2, SKILL_COL0, 999, SKILLS.length);
    const rule = (txt, color) => SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(txt).setBackground(color).setRanges([skillRange]).build();
    sum.setConditionalFormatRules([rule('金', '#f6d77a'), rule('銀', '#dfe4ea'), rule('銅', '#f0c9a3')]);
  }
  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG, 1);
    log.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold').setBackground('#efe3c4');
    log.setFrozenRows(1);
    log.getRange('C:D').setNumberFormat('@');
    log.getRange('A:B').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    log.setColumnWidth(12, 320);
    log.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('否').setFontColor('#b0443a').setRanges([log.getRange('O2:P')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('是').setFontColor('#28695e').setRanges([log.getRange('O2:P')]).build()
    ]);
    log.getRange(1, 1, 1, LOG_HEAD.length).createFilter();
  }
  if (!ss.getSheetByName(SHEET_STAT)) ss.insertSheet(SHEET_STAT, 2);
}

function yn_(v) { return v === undefined || v === null || v === '' ? '' : (v ? '是' : '否'); }
/* 防止學生輸入的文字被當成公式（例如以 = 開頭） */
function safe_(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }
