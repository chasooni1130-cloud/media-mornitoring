/**
 * 11_ManualEntry.gs — F-104 수동 등록
 * 지면·방송 등 API로 수집되지 않는 보도는 담당자가 직접 입력한다.
 * 01_모니터링현황 시트에 행을 직접 입력해도 되지만(수집경로=수동으로 표기),
 * 메뉴에서 실행하면 기사ID 자동 생성 등 형식을 맞춰 준다.
 */

function addManualArticle() {
  const ui = SpreadsheetApp.getUi();
  const title = promptText_(ui, '제목', '보도 제목을 입력하세요.');
  if (title === null) return;
  const media = promptText_(ui, '매체명', '지면/방송 매체명을 입력하세요.');
  if (media === null) return;
  const pubDateText = promptText_(ui, '보도일시', 'YYYY-MM-DD 또는 YYYY-MM-DD HH:mm 형식으로 입력하세요.');
  if (pubDateText === null) return;
  const note = promptText_(ui, '비고(선택)', '지면 단독/방송 리포트 등 특이사항을 입력하세요. 없으면 비워두세요.');

  const pubDate = new Date(pubDateText.replace(/-/g, '/'));
  if (isNaN(pubDate)) { ui.alert('보도일시 형식을 확인해 주세요.'); return; }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  const existingIds = lastRow >= 2 ? sh.getRange(2, LEDGER_COL.기사ID, lastRow - 1, 1).getValues().flat() : [];
  const articleId = generateArticleId_(nextSeqForToday_(existingIds));

  const row = new Array(LEDGER_LAST_COL).fill('');
  row[LEDGER_COL.기사ID - 1] = articleId;
  row[LEDGER_COL.수집일시 - 1] = new Date();
  row[LEDGER_COL.보도일시 - 1] = pubDate;
  row[LEDGER_COL.매체명 - 1] = media;
  row[LEDGER_COL.제목 - 1] = title;
  row[LEDGER_COL.수집경로 - 1] = '수동';
  row[LEDGER_COL.상태 - 1] = '미확인';
  row[LEDGER_COL.대응상태 - 1] = '해당없음';
  row[LEDGER_COL.비고 - 1] = note || '';

  sh.appendRow(row);
  sortLedgerByReportDate_(sh);
  ui.alert('추가되었습니다: ' + articleId + ' — 매체유형·매체영향력·논조 등은 01_모니터링현황에서 직접 확정해 주세요.');
}

function promptText_(ui, title, message) {
  const resp = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  return resp.getResponseText().trim();
}
