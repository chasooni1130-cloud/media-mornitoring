/**
 * 15_WebApp.gs — 웹앱 진입점 (제어판 + 대시보드)
 * 배포(배포 > 새 배포 > 웹 앱) 후 발급되는 /exec URL로 접속하는 외부 웹페이지.
 * ?page=dashboard 쿼리로 대시보드, 그 외(기본값)는 제어판(수집 실행 버튼)을 보여준다.
 *
 * ?action=... 쿼리가 있으면 HTML 대신 JSON API로 응답한다(사용자 요청 2026-10: Vercel에 올리는
 * 정적 페이지가 google.script.run 대신 fetch()로 이 Apps Script를 백엔드 삼아 호출하기 위함).
 * 배포 접근권한이 "모든 사용자"여야 로그인 없는 방문자의 fetch 요청도 통과한다.
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action) return handleApiGet_(action);

  const page = (e && e.parameter && e.parameter.page) || 'index';
  const file = page === 'dashboard' ? 'Dashboard' : 'Index';
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle('언론보도 모니터링')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** JSON API 분기. action=status|dashboard|collect. 항상 {ok, data|error} 형태로 응답한다. */
function handleApiGet_(action) {
  let payload;
  try {
    if (action === 'status') {
      payload = { ok: true, data: getStatusForWeb_() };
    } else if (action === 'dashboard') {
      payload = { ok: true, data: getDashboardData_() };
    } else if (action === 'collect') {
      payload = { ok: true, data: runCollectionFromWeb_() };
    } else {
      payload = { ok: false, error: '알 수 없는 action: ' + action };
    }
  } catch (err) {
    payload = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 웹 제어판의 "지금 수집 실행" 버튼에서 호출. shouldRunNow_ 스로틀을 무시하고 즉시 1회 수집한다. */
function runCollectionFromWeb_() {
  const startTime = Date.now();
  const before = getLedgerCount_();
  runCollection_();
  const after = getLedgerCount_();
  return {
    savedCount: after - before,
    durationMs: Date.now() - startTime,
    finishedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}

function getLedgerCount_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  return lastRow >= 2 ? lastRow - 1 : 0;
}

/** 웹 제어판에서 마지막 수집 시각·오늘 수집 건수 등 현재 상태를 보여줄 때 호출. */
function getStatusForWeb_() {
  const tz = Session.getScriptTimeZone();
  const today = stripTime_(new Date());
  const tomorrow = addDays_(today, 1);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  const allRows = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues() : [];
  const todayCount = allRows.filter(row => {
    const pub = row[LEDGER_COL.보도일시 - 1];
    return pub instanceof Date && pub >= today && pub < tomorrow;
  }).length;
  return {
    lastCollectAt: Utilities.formatDate(getLastCollectTs_(), tz, 'yyyy-MM-dd HH:mm:ss'),
    totalCount: allRows.length,
    todayCount: todayCount
  };
}

/**
 * 대시보드 웹페이지에서 호출. 09_Dashboard.gs의 refreshDashboard()와 같은 집계 로직을
 * 시트에 쓰는 대신 JSON으로 그대로 반환한다(시트 쪽 refreshDashboard는 그대로 둔다).
 */
function getDashboardData_() {
  const tz = Session.getScriptTimeZone();
  const today = stripTime_(new Date());
  const tomorrow = addDays_(today, 1);

  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = ledger.getLastRow();
  const allRows = lastRow >= 2 ? ledger.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues() : [];

  const todayRows = allRows.filter(row => {
    const pub = row[LEDGER_COL.보도일시 - 1];
    return pub instanceof Date && pub >= today && pub < tomorrow;
  });
  const unconfirmedCount = allRows.filter(row => row[LEDGER_COL.상태 - 1] === '미확인').length;

  const last7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = addDays_(today, -i);
    const next = addDays_(d, 1);
    const cnt = allRows.filter(row => {
      const pub = row[LEDGER_COL.보도일시 - 1];
      return pub instanceof Date && pub >= d && pub < next;
    }).length;
    last7.push({ date: Utilities.formatDate(d, tz, 'MM-dd(E)'), count: cnt });
  }

  const crisisCounts = {};
  OPTS.위기등급.forEach(g => crisisCounts[g] = 0);
  allRows.forEach(row => {
    const g = row[LEDGER_COL.위기등급 - 1] || '해당없음';
    crisisCounts[g] = (crisisCounts[g] || 0) + 1;
  });

  const toRow = (row) => ({
    pubDate: row[LEDGER_COL.보도일시 - 1] instanceof Date
      ? Utilities.formatDate(row[LEDGER_COL.보도일시 - 1], tz, 'MM-dd HH:mm') : '',
    media: row[LEDGER_COL.매체명 - 1],
    title: row[LEDGER_COL.제목 - 1],
    url: row[LEDGER_COL.URL - 1],
    grade: row[LEDGER_COL.위기등급 - 1],
    status: row[LEDGER_COL.상태 - 1],
    responseStatus: row[LEDGER_COL.대응상태 - 1],
    tone: row[LEDGER_COL.논조확정 - 1] || row[LEDGER_COL.논조자동 - 1]
  });

  const negRows = allRows.filter(row =>
    row[LEDGER_COL.보도일시 - 1] instanceof Date && row[LEDGER_COL.보도일시 - 1] >= addDays_(today, -6) &&
    (row[LEDGER_COL.논조확정 - 1] === '부정' || row[LEDGER_COL.논조자동 - 1] === '부정'));

  const ongoing = allRows.filter(row =>
    row[LEDGER_COL.위기등급 - 1] && row[LEDGER_COL.위기등급 - 1] !== '해당없음' &&
    row[LEDGER_COL.대응상태 - 1] !== '종결');

  const recentRows = allRows
    .filter(row => row[LEDGER_COL.보도일시 - 1] instanceof Date)
    .sort((a, b) => b[LEDGER_COL.보도일시 - 1] - a[LEDGER_COL.보도일시 - 1])
    .slice(0, 30);

  return {
    generatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    todayCount: todayRows.length,
    totalCount: allRows.length,
    unconfirmedCount: unconfirmedCount,
    last7: last7,
    crisisCounts: crisisCounts,
    negativeRecent: negRows.map(toRow),
    ongoingCrisis: ongoing.map(toRow),
    recent: recentRows.map(toRow)
  };
}
