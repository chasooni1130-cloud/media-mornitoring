/**
 * 10_Report.gs — F-408 리포트 자동 생성 및 메일 발송
 * 주간: 매주 월요일 09:00 / 월간: 익월 3영업일 09:00 (installTriggers()가 스케줄을 등록한다)
 * 수신자는 00_설정!수신자 블록에서 '수신등급'에 "리포트"를 포함한 사람으로 한다.
 */

function sendWeeklyReportEmail() {
  buildWeeklyStats();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const day = Number(Utilities.formatDate(now, tz, 'u'));
  const weekStart = stripTime_(addDays_(now, -(day - 1)));
  const weekEnd = addDays_(weekStart, 7);
  const prevStart = addDays_(weekStart, -7);

  const cur = computeLedgerStats_(weekStart, weekEnd);
  const prev = computeLedgerStats_(prevStart, weekStart);

  const html = [
    '<h2>주간 언론보도 모니터링 리포트</h2>',
    '<p>기간: ' + fmtDate_(weekStart) + ' ~ ' + fmtDate_(addDays_(weekEnd, -1)) + '</p>',
    '<ul>',
    '<li>총 보도 건수: <b>' + cur.total + '건</b> (전주 대비 ' + pctChange_(cur.total, prev.total) + '%)</li>',
    '<li>순 이슈 건수: ' + cur.issueCount + '건</li>',
    '<li>부정 보도: ' + (cur.byTone['부정'] || 0) + '건</li>',
    '</ul>',
    buildDistHtmlTable_('매체유형별 분포', cur.byMediaType, cur.total),
    buildDistHtmlTable_('논조별 분포', cur.byTone, cur.total),
    '<p>주요 보도 하이라이트</p><ul>' +
      cur.highlights.map(h => '<li>' + h.media + ' - <a href="' + h.link + '">' + h.title + '</a></li>').join('') + '</ul>',
    '<p>상세 표는 스프레드시트 03_주간통계 시트를 확인해 주세요: ' +
      '<a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">시트 열기</a></p>'
  ].join('\n');

  sendReportEmail_('[주간 리포트] 언론보도 모니터링 (' + fmtDate_(weekStart) + '~' + fmtDate_(addDays_(weekEnd, -1)) + ')', html);
}

function sendMonthlyReportEmail() {
  buildMonthlyStats();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(now, tz, 'M'));
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd = new Date(y, m, 1);
  const prevMonth = computeLedgerStats_(new Date(y, m - 2, 1), monthStart);
  const cur = computeLedgerStats_(monthStart, monthEnd);

  const html = [
    '<h2>월간 언론보도 모니터링 리포트</h2>',
    '<p>기간: ' + y + '년 ' + m + '월</p>',
    '<ul>',
    '<li>총 보도 건수: <b>' + cur.total + '건</b> (전월 대비 ' + pctChange_(cur.total, prevMonth.total) + '%)</li>',
    '<li>부정 보도: ' + (cur.byTone['부정'] || 0) + '건</li>',
    '</ul>',
    buildDistHtmlTable_('매체유형별 분포', cur.byMediaType, cur.total),
    '<p>상세 표는 스프레드시트 04_월간통계 시트를 확인해 주세요: ' +
      '<a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">시트 열기</a></p>'
  ].join('\n');

  sendReportEmail_('[월간 리포트] 언론보도 모니터링 (' + y + '년 ' + m + '월)', html);
}

/** 매일 09:00 트리거에서 호출 — 오늘이 이번 달의 3번째 영업일이면 월간 리포트를 발송한다. */
function dailyCheckMonthlyReportDue_() {
  const tz = Session.getScriptTimeZone();
  const today = stripTime_(new Date());
  if (today.getTime() !== stripTime_(get3rdBusinessDayOfMonth_(today)).getTime()) return;
  sendMonthlyReportEmail();
}

function get3rdBusinessDayOfMonth_(anyDateInMonth) {
  const y = anyDateInMonth.getFullYear(), m = anyDateInMonth.getMonth();
  let count = 0;
  let d = new Date(y, m, 1);
  while (count < 3) {
    const dow = d.getDay(); // 0=Sun,6=Sat
    if (dow !== 0 && dow !== 6) count++;
    if (count === 3) break;
    d = addDays_(d, 1);
  }
  return d;
}

function buildDistHtmlTable_(title, map, total) {
  const rows = Object.keys(map).map(k => [k, map[k], pctOf_(map[k], total)]).sort((a, b) => b[1] - a[1]);
  return '<p><b>' + title + '</b></p><table border="1" cellpadding="4" style="border-collapse:collapse">' +
    '<tr><th>구분</th><th>건수</th><th>비율(%)</th></tr>' +
    rows.map(r => '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td>' + r[2] + '</td></tr>').join('') +
    '</table>';
}

function sendReportEmail_(subject, htmlBody) {
  const recipients = readConfigBlock_('수신자').filter(r => {
    const agree = r['수신동의'] === 'Y' || r['수신동의'] === true;
    const grades = String(r['수신등급'] || '').split(',').map(s => s.trim());
    return agree && (grades.indexOf('리포트') >= 0 || grades.indexOf('전체') >= 0) && String(r['연락처']).indexOf('@') > 0;
  }).map(r => r['연락처']);
  if (!recipients.length) {
    writeLog_('sendReportEmail_', 0, 0, '리포트 수신자 없음(00_설정!수신자의 수신등급에 "리포트" 추가 필요)', 0);
    return;
  }
  MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: htmlBody });
}
