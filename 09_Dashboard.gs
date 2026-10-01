/**
 * 09_Dashboard.gs — F-404 실시간 대시보드
 * 당일 보도 건수, 미확인 건수, 최근 7일 추이, 부정 보도 현황, 진행 중 위기 이슈를 한 화면에 표시한다.
 */

function refreshDashboard() {
  const sh = getOrCreateSheet_(SHEETS.DASHBOARD);
  sh.clear();
  const tz = Session.getScriptTimeZone();
  const today = stripTime_(new Date());
  const tomorrow = addDays_(today, 1);

  let r = 1;
  r = writeTitle_(sh, r, '실시간 대시보드 — 기준시각 ' + nowStr_());

  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = ledger.getLastRow();
  const allRows = lastRow >= 2 ? ledger.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues() : [];

  const todayRows = allRows.filter(row => {
    const pub = row[LEDGER_COL.보도일시 - 1];
    return pub instanceof Date && pub >= today && pub < tomorrow;
  });
  const unconfirmedCount = allRows.filter(row => row[LEDGER_COL.상태 - 1] === '미확인').length;

  r = writeKv_(sh, r, [
    ['당일 보도 건수', todayRows.length],
    ['미확인 건수(전체 누적)', unconfirmedCount]
  ]);
  r++;

  const last7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = addDays_(today, -i);
    const next = addDays_(d, 1);
    const cnt = allRows.filter(row => {
      const pub = row[LEDGER_COL.보도일시 - 1];
      return pub instanceof Date && pub >= d && pub < next;
    }).length;
    last7.push([Utilities.formatDate(d, tz, 'MM-dd(E)'), cnt]);
  }
  r = writeTable_(sh, r, '최근 7일 추이', ['날짜', '건수'], last7);

  const negRows = allRows.filter(row =>
    row[LEDGER_COL.보도일시 - 1] instanceof Date && row[LEDGER_COL.보도일시 - 1] >= addDays_(today, -6) &&
    (row[LEDGER_COL.논조확정 - 1] === '부정' || row[LEDGER_COL.논조자동 - 1] === '부정'));
  r = writeTable_(sh, r, '최근 7일 부정 보도 현황', ['매체명', '보도일시', '제목', '상태', '대응상태'],
    negRows.map(row => [row[LEDGER_COL.매체명 - 1], fmtDateTimeCell_(row[LEDGER_COL.보도일시 - 1]),
      row[LEDGER_COL.제목 - 1], row[LEDGER_COL.상태 - 1], row[LEDGER_COL.대응상태 - 1]]));

  const ongoing = allRows.filter(row =>
    row[LEDGER_COL.위기등급 - 1] && row[LEDGER_COL.위기등급 - 1] !== '해당없음' &&
    row[LEDGER_COL.대응상태 - 1] !== '종결');
  r = writeTable_(sh, r, '진행 중 위기 이슈', ['위기등급', '매체명', '제목', '이슈그룹ID', '대응상태', '링크'],
    ongoing.map(row => [row[LEDGER_COL.위기등급 - 1], row[LEDGER_COL.매체명 - 1], row[LEDGER_COL.제목 - 1],
      row[LEDGER_COL.이슈그룹ID - 1], row[LEDGER_COL.대응상태 - 1], row[LEDGER_COL.URL - 1]]));

  formatStatsSheet_(sh);
}

function fmtDateTimeCell_(d) {
  return d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM-dd HH:mm') : '';
}
