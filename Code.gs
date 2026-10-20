/**
 * 하이커그라운드 층별 비품 보드 — Google Apps Script 백엔드
 * 구글 시트 4개 탭(공간목록 / 비품목록 / 초기배치 / 이동기록)을 데이터베이스로 사용합니다.
 * 탭 구조는 제공된 엑셀 '하이커그라운드_층별_비품관리대장.xlsx'와 동일합니다.
 */

const SHEETS = { spaces: '공간목록', items: '비품목록', init: '초기배치', moves: '이동기록' };
const TZ = 'Asia/Seoul';

const HTML_URL = 'https://raw.githubusercontent.com/parkjh3737/hikr-board/main/Index.html';

/** 화면(Index.html)은 GitHub 저장소에서 읽어옵니다 → 화면 변경은 GitHub에 올리기만 하면 반영 (최대 5분 캐시) */
function doGet(e) {
  let html;
  const cache = CacheService.getScriptCache();
  const nocache = e && e.parameter && e.parameter.nocache;
  if (!nocache) html = cache.get('index_html');
  if (!html) {
    const res = UrlFetchApp.fetch(HTML_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">화면 파일을 GitHub에서 불러오지 못했습니다 (HTTP ' + res.getResponseCode() + '). 잠시 후 다시 시도하세요.</p>');
    }
    html = res.getContentText();
    cache.put('index_html', html, 300);
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('하이커 비품 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(key) {
  const s = ss().getSheetByName(SHEETS[key]);
  if (!s) throw new Error("'" + SHEETS[key] + "' 탭이 없습니다. 시트 탭 이름을 확인하세요.");
  return s;
}
function fmtDate(v) {
  if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v ? String(v).slice(0, 10) : '';
}
function toDate(s) { // 'yyyy-mm-dd' → Date (시트에 날짜로 저장)
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
}
function rows(key) { // 헤더 제외 데이터 행 (시트 행 번호 포함)
  const s = sh(key);
  const last = s.getLastRow();
  if (last < 2) return [];
  const vals = s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();
  return vals.map((r, i) => ({ row: i + 2, v: r }));
}

/** 보드가 읽는 전체 데이터 */
function getData() {
  const spaces = rows('spaces').filter(r => r.v[2]).map(r => ({ floor: String(r.v[0]), name: String(r.v[2]).trim() }));
  const items = rows('items').filter(r => r.v[0]).map(r => ({
    row: r.row, id: String(r.v[0]).trim(), name: String(r.v[1]), cat: String(r.v[2] || ''), spec: String(r.v[3] || ''),
    unit: String(r.v[4] || ''), min: Number(r.v[5]) || 0, note: String(r.v[6] || '')
  }));
  const init = rows('init').filter(r => r.v[1] && r.v[3]).map(r => ({
    row: r.row, date: fmtDate(r.v[0]), id: String(r.v[1]).trim(), space: String(r.v[3]).trim(), qty: Number(r.v[4]) || 0,
    state: String(r.v[5] || ''), note: String(r.v[6] || '')
  }));
  const moves = rows('moves').filter(r => r.v[1] && (r.v[3] || r.v[4])).map(r => ({
    row: r.row, date: fmtDate(r.v[0]), id: String(r.v[1]).trim(), from: String(r.v[3]).trim(), to: String(r.v[4]).trim(),
    qty: Number(r.v[5]) || 0, who: String(r.v[6] || ''), why: String(r.v[7] || ''), ok: String(r.v[8] || 'O').trim().toUpperCase() || 'O'
  }));
  return { spaces, items, init, moves, sheetUrl: ss().getUrl(), updatedAt: new Date().toISOString() };
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function nameFormula(r) { return '=IF(B' + r + '="","",IFERROR(INDEX(비품목록!$B:$B,MATCH(B' + r + ',비품목록!$A:$A,0)),"ID없음"))'; }
function itemExists(id) { return rows('items').some(r => String(r.v[0]).trim() === id); }
function spaceExists(name) { return rows('spaces').some(r => String(r.v[2]).trim() === name); }

/** 특정 공간의 현재 보유 수량 (초기배치 + 도착 − 출발, 적용 O만) */
function stockAt(id, space) {
  let n = 0;
  rows('init').forEach(r => { if (String(r.v[1]).trim() === id && String(r.v[3]).trim() === space) n += Number(r.v[4]) || 0; });
  rows('moves').forEach(r => {
    if (String(r.v[1]).trim() !== id || String(r.v[8] || 'O').trim().toUpperCase() !== 'O') return;
    if (String(r.v[4]).trim() === space) n += Number(r.v[5]) || 0;
    if (String(r.v[3]).trim() === space) n -= Number(r.v[5]) || 0;
  });
  return n;
}

/** 이동 기록 추가 */
function addMove(m) {
  return withLock(() => {
    if (!m || !m.id || !m.from || !m.to) throw new Error('비품·출발·도착 공간이 필요합니다.');
    if (!itemExists(m.id)) throw new Error('등록되지 않은 비품ID: ' + m.id);
    if (!spaceExists(m.from) || !spaceExists(m.to)) throw new Error('공간목록에 없는 공간입니다.');
    const qty = Number(m.qty); if (!(qty > 0)) throw new Error('수량은 1 이상이어야 합니다.');
    const have = stockAt(m.id, m.from);
    if (qty > have) throw new Error(m.from + '에 ' + m.id + '은(는) ' + have + '개뿐입니다. 수량을 ' + have + ' 이하로 입력하세요.');
    const s = sh('moves'); const r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 9).setValues([[toDate(m.date), m.id, '', m.from, m.to, qty, m.who || '', m.why || '', 'O']]);
    s.getRange(r, 1).setNumberFormat('yyyy-mm-dd');
    s.getRange(r, 3).setFormula(nameFormula(r));
    return getData();
  });
}

/** 이동 기록 삭제 (행 번호 + 내용 대조 후 삭제) */
function deleteMove(row, m) {
  return withLock(() => {
    const s = sh('moves');
    const v = s.getRange(row, 1, 1, 9).getValues()[0];
    const same = String(v[1]).trim() === m.id && String(v[3]).trim() === m.from && String(v[4]).trim() === m.to && Number(v[5]) === Number(m.qty) && fmtDate(v[0]) === m.date;
    if (!same) throw new Error('다른 사람이 먼저 수정했습니다. 새로고침 후 다시 시도하세요.');
    s.deleteRow(row);
    return getData();
  });
}

/** 입고(초기배치) 추가 */
function addStock(r) {
  return withLock(() => {
    if (!r || !r.id || !r.space) throw new Error('비품과 공간이 필요합니다.');
    if (!itemExists(r.id)) throw new Error('등록되지 않은 비품ID: ' + r.id);
    if (!spaceExists(r.space)) throw new Error('공간목록에 없는 공간입니다.');
    const qty = Number(r.qty); if (!(qty > 0)) throw new Error('수량은 1 이상이어야 합니다.');
    const s = sh('init'); const row = s.getLastRow() + 1;
    s.getRange(row, 1, 1, 7).setValues([[toDate(r.date), r.id, '', r.space, qty, r.state || '배치', r.note || '']]);
    s.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
    s.getRange(row, 3).setFormula(nameFormula(row));
    return getData();
  });
}

/** 신규 비품 등록 */
function addItem(it) {
  return withLock(() => {
    const id = String(it && it.id || '').trim();
    if (!/^[A-Za-z0-9_.-]+$/.test(id)) throw new Error('비품ID는 영문·숫자·-_.만 사용하세요 (예: E008).');
    if (!it.name) throw new Error('품명을 입력하세요.');
    if (itemExists(id)) throw new Error(id + '는 이미 등록된 ID입니다.');
    const s = sh('items'); const row = s.getLastRow() + 1;
    s.getRange(row, 1, 1, 7).setValues([[id, it.name, it.cat || '', it.spec || '', it.unit || '개', Number(it.min) || 0, it.note || '']]);
    return getData();
  });
}

/** 비품 정보 수정 (품명·분류·최소보유) — ID는 바꾸지 않음 */
function updateItem(id, patch) {
  return withLock(() => {
    const s = sh('items');
    const hit = rows('items').find(r => String(r.v[0]).trim() === id);
    if (!hit) throw new Error('비품을 찾을 수 없습니다: ' + id);
    if (!patch.name) throw new Error('품명을 입력하세요.');
    s.getRange(hit.row, 2).setValue(patch.name);
    s.getRange(hit.row, 3).setValue(patch.cat || '');
    s.getRange(hit.row, 6).setValue(Number(patch.min) || 0);
    return getData();
  });
}

/** 처음 설치 시 빈 시트에 탭·헤더·기본 공간을 만드는 함수 (Apps Script 편집기에서 1회 실행) */
function setupSheets() {
  const spec = {
    '공간목록': [['층', '공간코드', '공간명'],
      ['지하3층', 'B301', '지하3층 창고'], ['1층', 'S101', '1층 웰컴데스크'], ['1층', 'S102', '1.5층 계단뒤'], ['2층', 'S201', '2층 서브인포'], ['2층', 'S202', '2층 인포 창고'],
      ['2층', 'S203', '마이스테이지 시스템실'], ['2층', 'S204', '런드리 뒤'], ['2층', 'S205', '2층 비품창고'], ['2층', 'S206', '2층 화물엘베 앞'], ['3층', 'S301', '3층 사무국'],
      ['3층', 'S302', '3층 휴게실'], ['3층', 'S303', '3층 창고'], ['4층', 'S401', '4층 창고'], ['4층', 'S402', '4층 시스템실'],
      ['4층', 'S403', '이든사무실'], ['5층', 'S501', '5층 굿즈샵'], ['5층', 'S502', '5층 휴게공간'], ['-', 'S999', '폐기']],
    '비품목록': [['비품ID', '품명', '분류', '규격/상세', '단위', '최소보유수량', '비고']],
    '초기배치': [['등록일', '비품ID', '품명(자동)', '공간', '수량', '상태', '비고']],
    '이동기록': [['이동일', '비품ID', '품명(자동)', '출발공간', '도착공간', '수량', '담당자', '사유/비고', '적용(O/X)']]
  };
  const book = ss();
  Object.keys(spec).forEach(name => {
    let s = book.getSheetByName(name);
    if (!s) s = book.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, spec[name].length, spec[name][0].length).setValues(spec[name]);
      s.getRange(1, 1, 1, spec[name][0].length).setFontWeight('bold');
      s.setFrozenRows(1);
    }
  });
}
