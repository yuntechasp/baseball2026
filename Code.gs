// ═══════════════════════════════════════════════════════════
//  比賽報名系統 — Google Apps Script 後端  v1.2
//  部署方式：Extensions > Apps Script > 貼上此程式碼
//            Deploy > New deployment > Web App
//            Execute as: Me / Who has access: Anyone
// ═══════════════════════════════════════════════════════════

const SHEET_ID   = '17RHesPHN6OfHAJ48_yFOzurbFWN-QAVcEyFAZQeRVtQ';          // ← 貼上您的 Google Sheets ID
const ADMIN_PASS = 147258369'; // ← 修改為您的管理員密碼

const SH_SETTINGS = '賽事設定';
const SH_TEAMS    = '隊伍報名';
const SH_MEMBERS  = '球員名單';
const SH_NEWS     = '最新消息';
const SH_PAYMENT  = '繳費紀錄';
const SH_RULES    = '競賽規程';
const SH_DRAW     = '抽籤紀錄';
const SH_MATCHES  = '對戰時間表';

// ── 統一回應 ───────────────────────────────────────────────
function resp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORS Preflight ─────────────────────────────────────────
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── 路由 ───────────────────────────────────────────────────
function doGet(e) {
  // Handle POST-like requests via GET payload (CORS workaround)
  if (e.parameter.payload) {
    const body = JSON.parse(decodeURIComponent(e.parameter.payload));
    return doPostInternal(body);
  }
  const action = e.parameter.action || '';
  try {
    switch(action) {
      case 'getSettings': return resp(getSettings());
      case 'getNews':     return resp(getNews());
      case 'getTeams':    return resp(getTeams());
      case 'getMembers':  return resp(getMembersByTeam(e.parameter.teamId));
      case 'getAllData':  return resp(getAllData());
      case 'getPayments': return resp(getPayments());
      case 'getRules':    return resp(getRules());
      case 'getDrawData': return resp(getDrawData());
      default:            return resp({ error: 'Unknown action' });
    }
  } catch(err) {
    return resp({ error: err.message });
  }
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  return doPostInternal(body);
}

function doPostInternal(body) {
  const action = body.action || '';
  try {
    switch(action) {
      case 'submitTeam':     return resp(submitTeam(body));
      case 'submitPayment':  return resp(submitPayment(body));
      case 'adminLogin':     return resp(adminLogin(body.password));
      case 'updateStatus':   return resp(updateTeamStatus(body));
      case 'updateSettings': return resp(updateSettings(body));
      case 'addNews':        return resp(addNews(body));
      case 'deleteNews':     return resp(deleteNews(body.rowIndex));
      case 'updateMember':   return resp(updateMember(body));
      case 'deleteMember':   return resp(deleteMember(body));
      case 'deleteTeam':     return resp(deleteTeam(body.teamId));
      case 'updatePaymentStatus': return resp(updatePaymentStatus(body));
      case 'saveDraw':    return resp(saveDraw(body));
      case 'saveMatches': return resp(saveMatches(body));
      case 'resetDraw':   return resp(resetDraw());
      case 'toggleDraw':  return resp(toggleDraw(body.open));
      case 'saveRule':   return resp(saveRule(body));
      case 'deleteRule': return resp(deleteRule(body.ruleId));
      default:               return resp({ error: 'Unknown action' });
    }
  } catch(err) {
    return resp({ error: err.message });
  }
}

// ── 取得 / 初始化 Sheet ────────────────────────────────────
function getSheet(name) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); initSheet(sh, name); }
  return sh;
}

function initSheet(sh, name) {
  if (name === SH_SETTINGS) {
    sh.appendRow(['key','value']);
    const defaults = [
      ['title',       '2026全國大專夏季校際棒球錦標賽'],
      ['subtitle',    '大專男生組'],
      ['venue',       '雲林科大棒球場'],
      ['deadline',    '2026-07-20'],
      ['gamedate',    '2026-08-17'],
      ['gameend',     '2026-08-28'],
      ['drawdate',    ''],
      ['maxTeams',    '16'],
      ['open',        'true'],
      ['regFee',      '3000'],
      ['regFeeNote',  '每隊新台幣3,000元'],
      ['bankAccount', ''],
      ['bankName',    ''],
    ];
    defaults.forEach(r => sh.appendRow(r));
  } else if (name === SH_TEAMS) {
    sh.appendRow(['teamId','teamname','sport','group','manager','managerTel','coach','coachTel','email','status','createdAt','note']);
  } else if (name === SH_MEMBERS) {
    sh.appendRow(['teamId','role','name','number','studentId']);
  } else if (name === SH_NEWS) {
    sh.appendRow(['type','title','date','createdAt']);
    sh.appendRow(['info','系統啟用，歡迎各隊報名', new Date().toLocaleDateString('zh-TW'), new Date()]);
  } else if (name === SH_PAYMENT) {
    sh.appendRow(['payId','teamId','teamname','code','amount','payerName','note','status','createdAt']);
  } else if (name === SH_RULES) {
    sh.appendRow(['ruleId','title','content','pdfUrl','icon','order']);
  } else if (name === SH_DRAW) {
    sh.appendRow(['teamId','teamname','slot','drawnAt']);
  } else if (name === SH_MATCHES) {
    sh.appendRow(['matchNum','phase','teamA','teamB','date','time','court']);
  }
}

// ── 日期格式化 ─────────────────────────────────────────────
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return String(val);
}

// ── 設定 ───────────────────────────────────────────────────
function getSettings() {
  const sh = getSheet(SH_SETTINGS);
  const rows = sh.getDataRange().getValues().slice(1);
  const obj = {};
  const dateKeys = ['deadline','gamedate','gameend','drawdate'];
  rows.forEach(r => {
    if (!r[0]) return;
    obj[r[0]] = dateKeys.includes(r[0]) ? formatDate(r[1]) : String(r[1]);
  });
  return { ok: true, data: obj };
}

function updateSettings(body) {
  const sh = getSheet(SH_SETTINGS);
  const rows = sh.getDataRange().getValues();
  const keys = Object.keys(body.settings || {});
  keys.forEach(k => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === k) { sh.getRange(i+1,2).setValue(body.settings[k]); found = true; break; }
    }
    if (!found) sh.appendRow([k, body.settings[k]]);
  });
  return { ok: true };
}

// ── 最新消息 ────────────────────────────────────────────────
function getNews() {
  const sh = getSheet(SH_NEWS);
  const rows = sh.getDataRange().getValues().slice(1);
  return { ok: true, data: rows.filter(r=>r[0]).map((r,i) => ({
    rowIndex: i+2, type: r[0], title: r[1], date: formatDate(r[2])
  })).reverse() };
}

function addNews(body) {
  getSheet(SH_NEWS).appendRow([body.type, body.title, body.date, new Date()]);
  return { ok: true };
}

function deleteNews(rowIndex) {
  getSheet(SH_NEWS).deleteRow(Number(rowIndex));
  return { ok: true };
}

// ── 隊伍 ───────────────────────────────────────────────────
function getTeams() {
  const sh = getSheet(SH_TEAMS);
  const rows = sh.getDataRange().getValues().slice(1);
  return { ok: true, data: rows.filter(r=>r[0]).map(r => ({
    teamId:r[0], teamname:r[1], sport:r[2], group:r[3],
    manager:r[4], managerTel:r[5], coach:r[6], coachTel:r[7],
    email:r[8], status:r[9]||'pending', createdAt:r[10], note:r[11]||''
  })) };
}

function getAllData() {
  return {
    ok: true,
    settings: getSettings().data,
    news:     getNews().data,
    teams:    getTeams().data,
  };
}

function submitTeam(body) {
  const settings = getSettings().data;
  if (settings.open === 'false') return { ok: false, error: '報名已截止' };
  const teamsSh = getSheet(SH_TEAMS);
  const existing = teamsSh.getDataRange().getValues().slice(1);
  if (existing.find(r => r[1] === body.teamname)) return { ok: false, error: '此隊伍名稱已報名' };
  const teamId = 'T' + Date.now();
  const now = new Date().toLocaleString('zh-TW');
  teamsSh.appendRow([
    teamId, body.teamname, body.sport, body.group,
    body.manager, body.managerTel, body.coach, body.coachTel,
    body.email, 'pending', now, ''
  ]);
  const memSh = getSheet(SH_MEMBERS);
  if (body.manager) memSh.appendRow([teamId,'領隊',body.manager,'','']);
  memSh.appendRow([teamId,'教練',body.coach,'','']);
  memSh.appendRow([teamId,'隊長',body.captain.name,body.captain.num,body.captain.sid]);
  (body.members||[]).forEach((m,i) => {
    if (m.name) memSh.appendRow([teamId,`隊員${i+1}`,m.name,m.num,m.sid]);
  });
  // 寄送報名確認信
  if (body.email) {
    try { sendConfirmEmail(body, teamId); } catch(e) { Logger.log('Email error: ' + e); }
  }

  return { ok: true, teamId };
}

// ── 寄送報名確認信 ─────────────────────────────────────────
function sendConfirmEmail(body, teamId) {
  const settings = getSettings().data;
  const title    = settings.title || '比賽報名系統';
  const deadline = settings.deadline || '';
  const venue    = settings.venue    || '';
  const gamedate = settings.gamedate || '';
  const gameend  = settings.gameend  || '';
  const bankName = settings.bankName || '';
  const bankAcc  = settings.bankAccount || '';
  const fee      = settings.regFee   || '';
  const feeNote  = settings.regFeeNote || '';

  // 組合名單純文字
  let memberList = '';
  memberList += '【領隊】' + (body.manager || '未填') + '\n';
  memberList += '【教練】' + body.coach + '\n';
  memberList += '【隊長】' + body.captain.name
    + '  背號：' + body.captain.num
    + '  學號：' + body.captain.sid + '\n';
  (body.members || []).forEach((m, i) => {
    if (m.name) {
      memberList += '【隊員' + (i+1) + '】' + m.name + '  背號：' + m.num + '  學號：' + m.sid + '\n';
    }
  });

  const subject = `【${title}】${body.teamname} 報名確認通知`;

  var lines = [];
  lines.push(title);
  lines.push('');
  lines.push('親愛的 ' + body.teamname + ' 領隊/教練您好，');
  lines.push('');
  lines.push('感謝貴隊完成報名，以下為您的報名資料，請確認是否正確。');
  lines.push('如有任何問題，請聯絡主辦單位。');
  lines.push('');
  lines.push('===========================');
  lines.push('報名資料確認');
  lines.push('===========================');
  lines.push('隊伍名稱：' + body.teamname);
  lines.push('運動種類：' + (body.sport || '棒球'));
  lines.push('報名組別：' + (body.group || '大專男生組'));
  lines.push('教練姓名：' + body.coach);
  lines.push('聯絡信箱：' + body.email);
  lines.push('');
  lines.push('比賽資訊：');
  lines.push('・比賽日期：' + gamedate + (gameend ? ' ~ ' + gameend : ''));
  lines.push('・比賽場地：' + venue);
  lines.push('・報名截止：' + deadline);
  lines.push('');
  lines.push('---------------------------');
  lines.push('球員名單');
  lines.push('---------------------------');
  lines.push(memberList);
  lines.push('===========================');
  if (fee) {
    lines.push('保證金資訊');
    lines.push('・金額：NT$ ' + Number(fee).toLocaleString());
    if (feeNote)  lines.push('・說明：' + feeNote);
    if (bankName) lines.push('・匯款銀行：' + bankName);
    if (bankAcc)  lines.push('・匯款帳號：' + bankAcc);
    lines.push('');
    lines.push('請至報名頁面「繳費回報」填寫匯款資訊。');
    lines.push('===========================');
  }
  lines.push('主辦單位：雲科大體育室');
  lines.push('聯絡人：蔡小姐');
  lines.push('電話：05-5342601#2704');
  lines.push('Email：wanjan@yuntech.edu.tw');
  lines.push('');
  lines.push('此為系統自動發送郵件，請勿直接回覆。');
  var textBody = lines.join('\n');

  MailApp.sendEmail({
    to:      body.email,
    subject: subject,
    body:    textBody,
  });
}

function updateTeamStatus(body) {
  const sh = getSheet(SH_TEAMS);
  const rows = sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === body.teamId) {
      sh.getRange(i+1,10).setValue(body.status);
      if (body.note !== undefined) sh.getRange(i+1,12).setValue(body.note);
      return { ok: true };
    }
  }
  return { ok: false, error: '找不到隊伍' };
}

function deleteTeam(teamId) {
  const tsh = getSheet(SH_TEAMS);
  const trows = tsh.getDataRange().getValues();
  for (let i=trows.length-1; i>=1; i--) {
    if (trows[i][0]===teamId) { tsh.deleteRow(i+1); break; }
  }
  const msh = getSheet(SH_MEMBERS);
  const mrows = msh.getDataRange().getValues();
  for (let i=mrows.length-1; i>=1; i--) {
    if (mrows[i][0]===teamId) msh.deleteRow(i+1);
  }
  return { ok: true };
}

// ── 球員 ───────────────────────────────────────────────────
function getMembersByTeam(teamId) {
  const sh = getSheet(SH_MEMBERS);
  const rows = sh.getDataRange().getValues().slice(1);
  return { ok: true, data: rows.filter(r=>r[0]===teamId).map((r,i) => ({
    rowIndex:i, teamId:r[0], role:r[1], name:r[2], number:r[3], studentId:r[4]
  })) };
}

function updateMember(body) {
  const sh = getSheet(SH_MEMBERS);
  const rows = sh.getDataRange().getValues();
  let count = 0;
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===body.teamId) {
      if (count===body.memberIndex) {
        sh.getRange(i+1,3).setValue(body.name);
        sh.getRange(i+1,4).setValue(body.number);
        sh.getRange(i+1,5).setValue(body.studentId);
        return { ok: true };
      }
      count++;
    }
  }
  return { ok: false, error: '找不到球員' };
}

function deleteMember(body) {
  const sh = getSheet(SH_MEMBERS);
  const rows = sh.getDataRange().getValues();
  let count = 0;
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===body.teamId) {
      if (count===body.memberIndex) { sh.deleteRow(i+1); return { ok: true }; }
      count++;
    }
  }
  return { ok: false, error: '找不到球員' };
}

// ── 繳費紀錄 ────────────────────────────────────────────────
function submitPayment(body) {
  const sh = getSheet(SH_PAYMENT);
  const payId = 'P' + Date.now();
  const now = new Date().toLocaleString('zh-TW');
  sh.appendRow([
    payId, body.teamId, body.teamname,
    body.code, body.amount, body.payerName||'', body.note||'',
    'pending', now
  ]);
  return { ok: true, payId };
}

function getPayments() {
  const sh = getSheet(SH_PAYMENT);
  const rows = sh.getDataRange().getValues().slice(1);
  return { ok: true, data: rows.filter(r=>r[0]).map(r => ({
    payId:r[0], teamId:r[1], teamname:r[2],
    code:r[3], amount:r[4], payerName:r[5], note:r[6],
    status:r[7]||'pending', createdAt:r[8]
  })) };
}

function updatePaymentStatus(body) {
  const sh = getSheet(SH_PAYMENT);
  const rows = sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===body.payId) {
      sh.getRange(i+1,8).setValue(body.status);
      return { ok: true };
    }
  }
  return { ok: false, error: '找不到繳費紀錄' };
}


// ── 競賽規程 ────────────────────────────────────────────────
function getRules() {
  const sh = getSheet(SH_RULES);
  const rows = sh.getDataRange().getValues().slice(1);
  return { ok: true, data: rows.filter(r=>r[0]).map(r => ({
    ruleId:r[0], title:r[1], content:r[2], pdfUrl:r[3]||'', icon:r[4]||'📄', order:r[5]||0
  })).sort((a,b)=>a.order-b.order) };
}

function saveRule(body) {
  const sh = getSheet(SH_RULES);
  const rows = sh.getDataRange().getValues();
  // Update existing
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === body.ruleId) {
      sh.getRange(i+1,2).setValue(body.title);
      sh.getRange(i+1,3).setValue(body.content);
      sh.getRange(i+1,4).setValue(body.pdfUrl||'');
      sh.getRange(i+1,5).setValue(body.icon||'📄');
      sh.getRange(i+1,6).setValue(body.order||0);
      return { ok: true };
    }
  }
  // New rule
  const ruleId = 'R' + Date.now();
  sh.appendRow([ruleId, body.title, body.content, body.pdfUrl||'', body.icon||'📄', body.order||0]);
  return { ok: true, ruleId };
}

function deleteRule(ruleId) {
  const sh = getSheet(SH_RULES);
  const rows = sh.getDataRange().getValues();
  for (let i=rows.length-1; i>=1; i--) {
    if (rows[i][0]===ruleId) { sh.deleteRow(i+1); return { ok:true }; }
  }
  return { ok:false, error:'找不到規程' };
}


// ── 抽籤 ────────────────────────────────────────────────────
function getDrawData() {
  const drawSh = getSheet(SH_DRAW);
  const drawRows = drawSh.getDataRange().getValues().slice(1);
  const draws = drawRows.filter(r=>r[0]).map(r => ({
    teamId: r[0], teamname: r[1], slot: r[2]
  }));

  const matchSh = getSheet(SH_MATCHES);
  const matchRows = matchSh.getDataRange().getValues().slice(1);
  const matches = matchRows.filter(r=>r[0]).map(r => ({
    matchNum:r[0], phase:r[1], teamA:r[2], teamB:r[3],
    date:r[4], time:r[5], court:r[6]
  }));

  // Get draw open status from settings
  const settings = getSettings().data;
  const drawOpen = settings.drawOpen === 'true';

  return { ok: true, draws, matches, drawOpen };
}

function saveDraw(body) {
  // Check if already drawn
  const sh = getSheet(SH_DRAW);
  const rows = sh.getDataRange().getValues().slice(1);
  if (rows.find(r => r[0] === body.teamId)) return { ok: false, error: '已抽過籤' };
  if (rows.find(r => r[2] === body.slot))   return { ok: false, error: '此籤號已被抽走，請重試' };
  sh.appendRow([body.teamId, body.teamname, body.slot, new Date().toLocaleString('zh-TW')]);
  return { ok: true };
}

function saveMatches(body) {
  const sh = getSheet(SH_MATCHES);
  // Clear existing
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, 7).clearContent();
  // Write new
  (body.matches || []).forEach(m => {
    sh.appendRow([m.matchNum, m.phase, m.teamA, m.teamB, m.date, m.time, m.court]);
  });
  return { ok: true };
}

function resetDraw() {
  const sh = getSheet(SH_DRAW);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, 4).clearContent();
  const msh = getSheet(SH_MATCHES);
  const mlastRow = msh.getLastRow();
  if (mlastRow > 1) msh.getRange(2, 1, mlastRow-1, 7).clearContent();
  return { ok: true };
}

function toggleDraw(open) {
  const sh = getSheet(SH_SETTINGS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'drawOpen') {
      sh.getRange(i+1, 2).setValue(open ? 'true' : 'false');
      return { ok: true };
    }
  }
  sh.appendRow(['drawOpen', open ? 'true' : 'false']);
  return { ok: true };
}

// ── 管理員登入 ──────────────────────────────────────────────
function adminLogin(password) {
  return { ok: password === ADMIN_PASS };
}
