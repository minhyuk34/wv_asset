/**
 * 주주클럽 기부 현황 백엔드 (Google Apps Script)
 *
 * 주식기부 약정서와 증권사 이체 안내는 월드비전 홈페이지에서 받으므로,
 * 이 시트에는 개인정보(연락처·계좌 등)를 넣지 않습니다.
 * 자산기금팀이 입고를 확인할 때마다 '약정' 탭에 한 줄씩 입력하면 캠페인 페이지 그래프에 반영됩니다.
 *
 *   GET  ?action=summary  → 합계만 반환 (5분 캐시)
 *   GET  ?action=qna      → 답변이 게시된 Q&A 목록 (5분 캐시)
 *   POST {action:'ask'}   → 질문 접수 ('Q&A' 탭에 한 줄 추가, 검토 후 '공개' 선택 시 페이지에 게시)
 *
 * 사용법: 구글 시트에서 확장 프로그램 > Apps Script 에 붙여넣고 setup 실행 후 웹 앱으로 배포
 */

const SHEET_LEDGER = '약정';
const SHEET_SET    = '설정';
const CACHE_KEY    = 'summary_v2';
const CACHE_SEC    = 300;
const SHEET_QNA    = 'Q&A';
const QNA_KEY      = 'qna_v1';
const QNA_HEAD     = ['접수번호','접수일시','작성자','질문','공개용질문','답변','공개','답변일','메모','clientId'];
const QCOL = {}; QNA_HEAD.forEach((h, i) => QCOL[h] = i);

const HEAD = ['접수일','이름','기금','유형','종목','수량','평가액(원)','상태','입고일','공개표시명','메모'];
const COL = {}; HEAD.forEach((h, i) => COL[h] = i);

const BIZ    = {domestic:'국내 꿈성장기금', global:'해외 자립마을 기금', relief:'긴급구호 기금'};
const TYPES  = ['기부형', '기금형'];
const STATUS = ['약정', '입고완료', '취소'];

/* ───────── 최초 1회 실행 ───────── */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getSheetByName(SHEET_LEDGER) || ss.insertSheet(SHEET_LEDGER);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#FFE9DA');
  sh.setFrozenRows(1);
  const dv = (list) => SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sh.getRange(2, COL['기금'] + 1, 1000, 1).setDataValidation(dv(Object.keys(BIZ).map(k => BIZ[k])));
  sh.getRange(2, COL['유형'] + 1, 1000, 1).setDataValidation(dv(TYPES));
  sh.getRange(2, COL['상태'] + 1, 1000, 1).setDataValidation(dv(STATUS));
  sh.getRange(2, COL['접수일'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COL['입고일'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COL['평가액(원)'] + 1, 1000, 1).setNumberFormat('#,##0');
  sh.setColumnWidths(1, HEAD.length, 110);
  sh.getRange(1, COL['이름'] + 1).setNote('입고 확인 때 기부자를 찾기 위한 내부용입니다. 주민번호·주소·연락처는 이 시트에 넣지 마세요. (공개 API로는 나가지 않습니다)');
  sh.getRange(1, COL['공개표시명'] + 1).setNote('참여자가 이름·팀명 공개에 동의한 경우에만 적으세요. 비워두면 익명으로 집계됩니다.');
  sh.getRange(1, COL['평가액(원)'] + 1).setNote('이체일 KRX 종가(취득가가 더 낮으면 취득가) × 수량. 비워두면 인원·주식 수에만 반영됩니다.');

  const st = ss.getSheetByName(SHEET_SET) || ss.insertSheet(SHEET_SET);
  if (st.getLastRow() < 2) {
    st.getRange(1, 1, 2, 3).setValues([
      ['항목', '값', '설명'],
      ['목표금액', 100000000, '그래프 목표 (원)'],
    ]);
    st.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#FFE9DA');
    st.setColumnWidth(1, 140); st.setColumnWidth(2, 200); st.setColumnWidth(3, 420);
  }
  ensureQna_();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  try { SpreadsheetApp.getUi().alert('설정 완료. 이제 배포 > 새 배포 > 웹 앱으로 배포하세요.'); }
  catch (e) { Logger.log('설정 완료. 이제 배포 > 배포 관리에서 새 버전으로 배포하세요.'); }   // 편집기에서 실행하면 getUi()를 쓸 수 없어 로그로 대신 안내
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('주주클럽')
    .addItem('기부 현황 새로고침(캐시 비우기)', 'clearCache')
    .addItem('초기 설정(setup)', 'setup')
    .addToUi();
}
function clearCache() { const c = CacheService.getScriptCache(); c.remove(CACHE_KEY); c.remove(QNA_KEY); }

/* ───────── 웹 앱 진입점 ───────── */
function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'summary';
  if (a === 'summary') return json_(getSummary_());
  if (a === 'qna') return json_(getQna_());
  return json_({ok: false, error: 'unknown action'});
}
function doPost(e) {
  try {
    const p = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (p.action === 'ask') return json_(askQuestion_(p));
    return json_({ok: false, error: 'unknown action'});
  } catch (err) {
    return json_({ok: false, error: String(err && err.message || err)});
  }
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function setting_(key) {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SET);
  if (!st) return '';
  const v = st.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) return v[i][1];
  return '';
}

/* ───────── 집계 (합계만, 개인정보 없음) ───────── */
function getSummary_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const s = buildSummary_();
  try { cache.put(CACHE_KEY, JSON.stringify(s), CACHE_SEC); } catch (e) {}
  return s;
}

function buildSummary_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEDGER);
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues() : [];
  const key = {}; Object.keys(BIZ).forEach(k => key[BIZ[k]] = k);
  const byBiz = {domestic: {p: 0, r: 0}, global: {p: 0, r: 0}, relief: {p: 0, r: 0}};
  let pledged = 0, received = 0, participants = 0, shares = 0;
  const weeks = {}, sup = [];

  rows.forEach(r => {
    const status = r[COL['상태']] || '약정';
    const k = key[r[COL['기금']]];
    const qty = Number(r[COL['수량']]) || 0, amt = Number(r[COL['평가액(원)']]) || 0;
    if (status === '취소' || (!k && !qty && !amt)) return;      // 빈 줄·취소 제외

    participants++; shares += qty; pledged += amt;
    if (k) byBiz[k].p += amt;
    if (status === '입고완료') { received += amt; if (k) byBiz[k].r += amt; }

    const d = r[COL['접수일']];
    if (d instanceof Date) {
      const mon = new Date(d); mon.setHours(0, 0, 0, 0); mon.setDate(mon.getDate() - ((mon.getDay() + 6) % 7));
      weeks[mon.getTime()] = (weeks[mon.getTime()] || 0) + amt;
    }
    const label = String(r[COL['공개표시명']] || '').trim();
    if (label) sup.push({t: d instanceof Date ? d.getTime() : 0, label: label.slice(0, 20), biz: r[COL['기금']] || '', qty: qty});
  });

  let acc = 0;
  const weekly = Object.keys(weeks).map(Number).sort((a, b) => a - b).map(w => acc += weeks[w]).slice(-12);
  sup.sort((a, b) => b.t - a.t);
  return {
    ok: true,
    goal: Number(setting_('목표금액')) || 100000000,
    pledged: pledged, received: received, participants: participants, shares: shares,
    byBiz: byBiz, weekly: weekly,
    supporters: sup.slice(0, 12).map(s => ({label: s.label, biz: s.biz, qty: s.qty})),
    updated: new Date().toISOString()
  };
}

/* ───────── Q&A 게시판 ───────── */
function ensureQna_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_QNA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_QNA);
    sh.getRange(1, 1, 1, QNA_HEAD.length).setValues([QNA_HEAD]).setFontWeight('bold').setBackground('#FFE9DA');
    sh.setFrozenRows(1);
    sh.getRange(2, QCOL['공개'] + 1, 1000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['공개', '숨김'], true).setAllowInvalid(false).build());
    sh.getRange(2, QCOL['접수일시'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(2, QCOL['답변일'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
    sh.setColumnWidths(1, QNA_HEAD.length, 110);
    sh.setColumnWidth(QCOL['질문'] + 1, 320); sh.setColumnWidth(QCOL['공개용질문'] + 1, 320); sh.setColumnWidth(QCOL['답변'] + 1, 420);
    sh.getRange(1, QCOL['공개용질문'] + 1).setNote('개인정보를 지우거나 다듬은 질문을 적으면 페이지에는 이 내용이 나옵니다. 비워두면 질문 칸 그대로 나옵니다.');
    sh.getRange(1, QCOL['공개'] + 1).setNote('답변을 적고 이 칸을 "공개"로 고르면 페이지에 게시됩니다. 비워두거나 "숨김"이면 게시되지 않습니다.');
    sh.getRange(1, QCOL['clientId'] + 1).setNote('중복 접수 방지용. 건드리지 마세요.');
  }
  return sh;
}

function safe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }   // 수식 주입 방지
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ' ').trim().slice(0, max); }

function askQuestion_(p) {
  if (clean_(p.website, 50)) return {ok: true, id: 'Q-000'};                 // 허니팟: 봇은 조용히 무시
  const q = clean_(p.question, 500), name = clean_(p.name, 30), cid = clean_(p.clientId, 64);
  if (q.length < 5) throw new Error('질문을 5자 이상 적어주세요.');
  if (!cid) throw new Error('요청이 올바르지 않습니다.');

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let id;
  try {
    const cache = CacheService.getScriptCache();
    const day = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
    const dayKey = 'qd_' + day, cliKey = 'qc_' + cid;
    const dayN = Number(cache.get(dayKey) || 0), cliN = Number(cache.get(cliKey) || 0);
    if (dayN >= 200) throw new Error('오늘 접수 가능한 질문 수를 넘었습니다. 내일 다시 남겨주세요.');
    if (cliN >= 5) throw new Error('질문을 짧은 시간에 너무 많이 남기셨어요. 잠시 후 다시 시도해주세요.');

    const sh = ensureQna_();
    const hit = sh.getRange(1, QCOL['clientId'] + 1, Math.max(sh.getLastRow(), 1), 1).createTextFinder(cid).matchEntireCell(true).findNext();
    if (hit) return {ok: true, id: sh.getRange(hit.getRow(), 1).getValue(), duplicate: true};   // 재시도 중복 방지

    id = 'Q-' + ('000' + sh.getLastRow()).slice(-3);
    const row = QNA_HEAD.map(() => '');
    row[QCOL['접수번호']] = id;
    row[QCOL['접수일시']] = new Date();
    row[QCOL['작성자']] = safe_(name);
    row[QCOL['질문']] = safe_(q);
    row[QCOL['clientId']] = cid;
    sh.appendRow(row);
    SpreadsheetApp.flush();
    cache.put(dayKey, String(dayN + 1), 86400);
    cache.put(cliKey, String(cliN + 1), 3600);
  } finally {
    lock.releaseLock();
  }
  return {ok: true, id: id};
}

function getQna_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(QNA_KEY);
  if (hit) return JSON.parse(hit);
  const out = buildQna_();
  try { cache.put(QNA_KEY, JSON.stringify(out), CACHE_SEC); } catch (e) {}
  return out;
}

function buildQna_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QNA);
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, QNA_HEAD.length).getValues() : [];
  const items = [];
  rows.forEach(r => {
    const a = String(r[QCOL['답변']] || '').trim();
    if (r[QCOL['공개']] !== '공개' || !a) return;                       // 공개 선택 + 답변이 있을 때만 게시
    const q = String(r[QCOL['공개용질문']] || '').trim() || String(r[QCOL['질문']] || '').trim();
    const d = r[QCOL['답변일']] instanceof Date ? r[QCOL['답변일']] : (r[QCOL['접수일시']] instanceof Date ? r[QCOL['접수일시']] : null);
    items.push({t: d ? d.getTime() : 0, q: q.slice(0, 500), a: a.slice(0, 2000), date: d ? Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd') : ''});
  });
  items.sort((x, y) => y.t - x.t);
  return {ok: true, items: items.slice(0, 50).map(i => ({q: i.q, a: i.a, date: i.date})), updated: new Date().toISOString()};
}
