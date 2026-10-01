/**
 * 12_PressMatch.gs — F-208 보도자료 매칭
 * 02_보도자료의 배포 건과 01_모니터링현황의 기사를 제목 유사도 + 배포일 근접성으로 매칭하여
 * 채택 여부를 표시한다. F-406(보도자료 성과)의 집계는 08_Stats.gs의 computePressStats_가 담당한다.
 */

const PRESS_MATCH_WINDOW_DAYS = 14;
const PRESS_MATCH_SIM_THRESHOLD = 0.45;

/** 메뉴/트리거에서 호출. 02_보도자료 각 건에 대해 모니터링 현황에서 채택 기사를 찾아 갱신한다. */
function matchPressReleases() {
  const pressSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRESS);
  const pressLast = pressSh.getLastRow();
  if (pressLast < 2) return;
  const pressData = pressSh.getRange(2, 1, pressLast - 1, PRESS_HEADERS.length).getValues();

  const ledgerSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const ledgerLast = ledgerSh.getLastRow();
  if (ledgerLast < 2) return;
  const ledgerData = ledgerSh.getRange(2, 1, ledgerLast - 1, LEDGER_LAST_COL).getValues();
  const assignedRows = new Set(); // 이번 실행 중 이미 다른 보도자료에 매칭된 행은 재매칭하지 않는다

  pressData.forEach((pRow, pIdx) => {
    const prId = pRow[0];
    const distDate = pRow[1];
    const title = pRow[2];
    if (!prId || !(distDate instanceof Date)) return;

    const windowStart = distDate;
    const windowEnd = addDays_(distDate, PRESS_MATCH_WINDOW_DAYS);
    const adoptedMedia = new Set();

    ledgerData.forEach((lRow, lIdx) => {
      const pub = lRow[LEDGER_COL.보도일시 - 1];
      if (!(pub instanceof Date) || pub < windowStart || pub > windowEnd) return;
      const sim = titleSimilarity_(title, lRow[LEDGER_COL.제목 - 1]);
      if (sim < PRESS_MATCH_SIM_THRESHOLD) return;

      adoptedMedia.add(lRow[LEDGER_COL.매체명 - 1]);
      if (!lRow[LEDGER_COL.보도자료ID - 1] && !assignedRows.has(lIdx)) {
        assignedRows.add(lIdx);
        ledgerSh.getRange(lIdx + 2, LEDGER_COL.보도자료ID).setValue(prId);
        if (!lRow[LEDGER_COL.보도유형 - 1]) ledgerSh.getRange(lIdx + 2, LEDGER_COL.보도유형).setValue('보도자료 기반');
      }
    });

    const distributed = Number(pRow[4]) || 0;
    const adoptedCount = adoptedMedia.size;
    const rate = distributed ? round1_(adoptedCount / distributed * 100) : 0;
    pressSh.getRange(pIdx + 2, 6, 1, 3).setValues([[adoptedCount, Array.from(adoptedMedia).join(', '), rate]]);
  });
}
