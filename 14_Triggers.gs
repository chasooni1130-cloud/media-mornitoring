/**
 * 14_Triggers.gs — 시간 기반 트리거 일괄 설치
 * 메뉴 [모니터링 시스템 → 2. 트리거 설치]에서 실행한다(최초 1회, 이후 필요 시 재실행해도 안전 — 중복 없이 재설치됨).
 */

const MANAGED_TRIGGER_FUNCTIONS = [
  'collectAll', 'refreshDashboard', 'sendMorningDigestForNightWarnings_',
  'sendDailyInterestSummary_', 'sendWeeklyReportEmail', 'dailyCheckMonthlyReportDue_',
  'backupLedger', 'matchPressReleases', 'checkMessageBalance_'
];

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (MANAGED_TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });

  // F-101/102/105: 5분마다 깨어나되, 실제 수집 여부는 collectAll() 내부 shouldRunNow_()가 결정한다.
  ScriptApp.newTrigger('collectAll').timeBased().everyMinutes(5).create();

  // F-404: 대시보드는 15분마다 갱신
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyMinutes(15).create();

  // F-508: 야간 주의등급 익일 07:00 일괄 발송
  ScriptApp.newTrigger('sendMorningDigestForNightWarnings_').timeBased().atHour(7).everyDays(1).create();

  // 6.4항: 관심 등급 일일 요약 08:30
  ScriptApp.newTrigger('sendDailyInterestSummary_').timeBased().atHour(8).nearMinute(30).everyDays(1).create();

  // F-408: 주간 리포트 — 매주 월요일 09:00
  ScriptApp.newTrigger('sendWeeklyReportEmail').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();

  // F-408: 월간 리포트 — 매일 09:00에 확인 후 3영업일에만 실제 발송
  ScriptApp.newTrigger('dailyCheckMonthlyReportDue_').timeBased().atHour(9).nearMinute(0).everyDays(1).create();

  // F-305: 주 1회 백업 — 매주 월요일 06:00
  ScriptApp.newTrigger('backupLedger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();

  // F-208: 보도자료 매칭 — 매일 06:30 (통계 산출 전에 최신화)
  ScriptApp.newTrigger('matchPressReleases').timeBased().atHour(6).nearMinute(30).everyDays(1).create();

  // F-604: 메시지 잔액 확인 — 매일 08:00
  ScriptApp.newTrigger('checkMessageBalance_').timeBased().atHour(8).everyDays(1).create();

  SpreadsheetApp.getUi().alert('트리거 설치 완료',
    '총 ' + MANAGED_TRIGGER_FUNCTIONS.length + '개 자동 실행 작업이 등록되었습니다.\n' +
    'Apps Script 편집기 좌측 [트리거] 메뉴에서 목록을 확인할 수 있습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** 문제가 생겼을 때 모든 관리 트리거를 제거하는 안전장치 (메뉴에는 노출하지 않음, 필요 시 편집기에서 직접 실행). */
function uninstallTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (MANAGED_TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}
