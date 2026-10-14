/**
 * 주주클럽 기부 현황 백엔드 (Google Apps Script)
 *
 * 주식기부 약정서와 증권사 이체 안내는 월드비전 홈페이지에서 받으므로,
 * 이 시트에는 개인정보(연락처·계좌 등)를 넣지 않습니다.
 * 자산기금팀이 입고를 확인할 때마다 '약정' 탭에 한 줄씩 입력하면 캠페인 페이지 그래프에 반영됩니다.
 *
 *   GET  ?action=summary  → 합계만 반환 (5분 캐시)
 *
 * 사용법: 구글 시트에서 확장 프로그램 > Apps Script 에 붙여넣고 setup 실행 후 웹 앱으로 배포
 */

const SHEET_LEDGER = '약정';
const SHEET_SET    = '설정';
const CACHE_KEY    = 'summary_v2';
const CACHE_SEC    = 300;

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
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('설정 완료. 이제 배포 > 새 배포 > 웹 앱으로 배포하세요.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('주주클럽')
    .addItem('기부 현황 새로고침(캐시 비우기)', 'clearCache')
    .addItem('초기 설정(setup)', 'setup')
    .addToUi();
}
function clearCache() { CacheService.getScriptCache().remove(CACHE_KEY); }

/* ───────── 웹 앱 진입점 ───────── */
function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'summary';
  if (a === 'summary') return json_(getSummary_());
  return json_({ok: false, error: 'unknown action'});
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
