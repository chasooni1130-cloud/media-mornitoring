/**
 * 07_Alert.gs — 위기 알림 발송 (F-504~512, 명세서 6.4~6.5항)
 *
 * 필요한 스크립트 속성:
 *   SOLAPI_API_KEY, SOLAPI_API_SECRET, SOLAPI_SENDER (발신번호, 예: 0212345678)
 *   KAKAO_PF_ID (카카오 비즈니스 발신 프로필 ID), KAKAO_TEMPLATE_CODE (승인된 알림톡 템플릿 코드)
 * 알림톡은 사전 승인된 템플릿 변수만 사용 가능하므로, 템플릿 심사 시
 * sendKakaoAlimtalk_()가 만드는 변수명(#{매체} #{등급} #{제목} #{보도일시} #{링크})으로 신청해야 한다.
 */

const GRADE_RANK = { '해당없음': 0, '관심': 1, '주의': 2, '경계': 3, '심각': 4 };

/** 파이프라인에서 신규로 위기등급이 붙은 기사들에 대해 알림 발송 여부를 판정하고 처리한다. */
function evaluateAndSendAlerts_(candidates) {
  const th = getThresholds_();
  candidates.forEach(c => {
    if (c.crisis.grade === '해당없음') return;

    if (c.crisis.grade === '관심') return; // F-508: 관심은 즉시 발송 없이 일일 요약에만 포함

    if (shouldSuppressForCooldown_(c.issueGroupId, c.crisis.grade, th)) return; // F-507

    const night = isNight_(th);
    if (c.crisis.grade === '주의' && night) {
      return; // F-508: 주의 등급의 야간 발생 건은 익일 07:00 일괄 발송(sendMorningDigestForNightWarnings_)이 처리
    }
    dispatchAlert_(c, c.crisis.grade);
  });
}

/** F-507 중복 알림 억제: 최초 1회 + 등급 상향 시에만 발송, 동일 등급 반복은 쿨다운 내 억제. */
function shouldSuppressForCooldown_(issueGroupId, grade, th) {
  const last = getLastAlertForGroup_(issueGroupId);
  if (!last) return false;
  if (GRADE_RANK[grade] > GRADE_RANK[last.grade]) return false; // 등급 상향 시 항상 발송
  const cooldownH = Number(th['알림_쿨다운_시간']) || 2;
  const elapsedH = (Date.now() - last.sentAt.getTime()) / 3600000;
  return elapsedH < cooldownH;
}

/** 07_알림이력에서 해당 이슈그룹의 가장 최근 발송 기록을 찾는다 (최근 2000행만 스캔).
 *  07_알림이력은 최신순(내림차순)으로 위쪽에 쌓이므로 2행부터 스캔한다. */
function getLastAlertForGroup_(issueGroupId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ALERT_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const scan = Math.min(2000, lastRow - 1);
  const startRow = 2;
  const idx = { 발송일시: 2, 이슈그룹ID: 4, 위기등급: 5 };
  const data = sh.getRange(startRow, 1, scan, ALERT_LOG_HEADERS.length).getValues();
  let best = null;
  data.forEach(r => {
    if (r[idx.이슈그룹ID - 1] !== issueGroupId) return;
    const sentAt = r[idx.발송일시 - 1];
    if (!(sentAt instanceof Date)) return;
    if (!best || sentAt > best.sentAt) best = { sentAt: sentAt, grade: r[idx.위기등급 - 1] };
  });
  return best;
}

function isNight_(th) {
  const tz = Session.getScriptTimeZone();
  const hm = Utilities.formatDate(new Date(), tz, 'HH:mm');
  const start = th['야간_시작'] || '22:00';
  const end = th['야간_종료'] || '07:00';
  // 22:00~07:00처럼 자정을 넘는 구간을 가정
  return hm >= start || hm < end;
}

/** 등급별 수신자 목록 (6.4항). 00_설정!수신자 블록의 '수신등급' 칸에 쉼표로 등급을 나열해 관리한다. */
function getRecipientsForGrade_(grade) {
  return readConfigBlock_('수신자').filter(r => {
    const agree = r['수신동의'] === 'Y' || r['수신동의'] === true;
    if (!agree) return false;
    const grades = String(r['수신등급'] || '').split(',').map(s => s.trim());
    return grades.indexOf(grade) >= 0 || grades.indexOf('전체') >= 0;
  });
}

function dispatchAlert_(candidate, grade) {
  const recipients = getRecipientsForGrade_(grade);
  const message = buildAlertMessage_(candidate, grade);
  if (!recipients.length) {
    appendAlertLog_(candidate, grade, '(수신자 없음)', '', '실패', message);
    return;
  }
  recipients.forEach(r => {
    const result = sendAlertToRecipient_(r, message, candidate, grade);
    appendAlertLog_(candidate, grade, r['이름'] + '(' + r['연락처'] + ')', result.channel, result.ok ? '성공' : '실패: ' + result.error, message);
  });
}

/**
 * 연락처가 이메일(@ 포함)이면 메일로, 그렇지 않으면 전화번호로 간주해 카카오 알림톡/SMS로 보낸다.
 * SOLAPI(카카오·SMS) 계약·API 키가 없는 담당자도 이메일만으로 실시간 알림을 받을 수 있게 하기 위한 대체 경로.
 */
function sendAlertToRecipient_(recipient, message, candidate, grade) {
  const contact = String(recipient['연락처'] || '').trim();
  if (contact.indexOf('@') > 0) return sendAlertEmail_(contact, message, candidate, grade);
  return sendAlimtalkWithSmsFallback_(recipient, message, candidate, grade);
}

function sendAlertEmail_(email, message, candidate, grade) {
  try {
    MailApp.sendEmail({ to: email, subject: '[언론 모니터링] ' + grade + ' 등급 감지 — ' + String(candidate.title).substring(0, 40), body: message });
    return { ok: true, channel: '이메일' };
  } catch (e) {
    return { ok: false, channel: '이메일', error: e.message };
  }
}

/** 6.5항 알림 메시지 양식. */
function buildAlertMessage_(candidate, grade) {
  const tz = Session.getScriptTimeZone();
  const pubDateStr = candidate.pubDate instanceof Date
    ? Utilities.formatDate(candidate.pubDate, tz, 'yyyy-MM-dd HH:mm') : String(candidate.pubDate);
  const lines = [
    '[언론 모니터링] ' + grade + ' 등급 감지',
    '▪ 매체: ' + candidate.media.표준매체명 + ' (' + (candidate.media.매체영향력 || '-') + ')',
    '▪ 보도: ' + pubDateStr,
    '▪ 제목: ' + candidate.title,
    '▪ 감지 키워드: ' + (candidate.crisis.reasonText || '-'),
    '▪ 위기점수: ' + candidate.crisis.score + '점',
    '▪ 링크: ' + candidate.link,
    '※ 확인 후 모니터링 현황 시트에서 대응상태를 갱신해 주십시오.'
  ];
  return lines.join('\n');
}

/** 사용자 요청: 07_알림이력은 항상 최신 발송건이 맨 위(2행)에 오도록, appendRow 대신 2행에 삽입한다. */
function prependAlertLogRow_(rowValues) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ALERT_LOG);
  sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
}

function appendAlertLog_(candidate, grade, recipientLabel, channel, resultLabel, message) {
  prependAlertLogRow_([
    'AL-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000),
    new Date(), candidate.articleId, candidate.issueGroupId, grade, candidate.crisis.score,
    recipientLabel, channel, resultLabel, message
  ]);
}

// ── 발송 채널 (카카오 알림톡 → 실패 시 SMS 대체, F-504) ──────────────────

function sendAlimtalkWithSmsFallback_(recipient, message, candidate, grade) {
  const phone = normalizePhone_(recipient['연락처']);
  if (!phone) return { ok: false, channel: '-', error: '연락처 형식 오류' };

  const kakaoResult = sendKakaoAlimtalk_(phone, candidate, grade);
  if (kakaoResult.ok) return { ok: true, channel: '알림톡' };

  const smsResult = sendSms_(phone, message);
  return smsResult.ok
    ? { ok: true, channel: 'SMS(대체발송)' }
    : { ok: false, channel: '알림톡+SMS 모두실패', error: kakaoResult.error + ' / ' + smsResult.error };
}

function sendKakaoAlimtalk_(phone, candidate, grade) {
  const apiKey = getSecret_('SOLAPI_API_KEY');
  const apiSecret = getSecret_('SOLAPI_API_SECRET');
  const sender = getSecret_('SOLAPI_SENDER');
  const pfId = getSecret_('KAKAO_PF_ID');
  const templateCode = getSecret_('KAKAO_TEMPLATE_CODE');
  if (!apiKey || !apiSecret || !sender || !pfId || !templateCode) {
    return { ok: false, error: '카카오 알림톡 설정(스크립트 속성) 미완료' };
  }
  const tz = Session.getScriptTimeZone();
  const variables = {
    '#{매체}': candidate.media.표준매체명,
    '#{등급}': grade,
    '#{제목}': candidate.title.substring(0, 60),
    '#{보도일시}': candidate.pubDate instanceof Date ? Utilities.formatDate(candidate.pubDate, tz, 'MM-dd HH:mm') : '',
    '#{링크}': candidate.link
  };
  const payload = {
    message: {
      to: phone, from: sender, kakaoOptions: {
        pfId: pfId, templateId: templateCode, variables: variables, disableSms: true
      }
    }
  };
  return solapiSend_(payload);
}

function sendSms_(phone, message) {
  const apiKey = getSecret_('SOLAPI_API_KEY');
  const apiSecret = getSecret_('SOLAPI_API_SECRET');
  const sender = getSecret_('SOLAPI_SENDER');
  if (!apiKey || !apiSecret || !sender) return { ok: false, error: 'SOLAPI 설정(스크립트 속성) 미완료' };
  const payload = { message: { to: phone, from: sender, text: message.substring(0, 2000) } };
  return solapiSend_(payload);
}

/**
 * SOLAPI(메시지 발송 대행) v4 단건 발송 API 호출.
 * 인증 방식(HMAC-SHA256)과 엔드포인트는 SOLAPI 공식 문서를 기준으로 하되,
 * 계약·API 버전에 따라 바뀔 수 있으므로 실 연동 전 https://solapi.com 문서와 대조 확인이 필요하다.
 */
function solapiSend_(payload) {
  const apiKey = getSecret_('SOLAPI_API_KEY');
  const apiSecret = getSecret_('SOLAPI_API_SECRET');
  try {
    const resp = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: solapiAuthHeader_(apiKey, apiSecret) },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) return { ok: true };
    return { ok: false, error: 'HTTP ' + code + ' ' + resp.getContentText().substring(0, 200) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function solapiAuthHeader_(apiKey, apiSecret) {
  const date = new Date().toISOString();
  const salt = Utilities.getUuid();
  const raw = Utilities.computeHmacSha256Signature(date + salt, apiSecret);
  const signature = raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  return 'HMAC-SHA256 apiKey=' + apiKey + ', date=' + date + ', salt=' + salt + ', signature=' + signature;
}

function normalizePhone_(phone) {
  if (!phone) return '';
  return String(phone).replace(/[^0-9]/g, '');
}

// ── F-508 야간·일일 배치 알림 ─────────────────────────────────────────

/** 트리거: 매일 07:00. 야간(22:00~07:00)에 발생한 '주의' 등급을 모아 일괄 발송한다. */
function sendMorningDigestForNightWarnings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const scan = Math.min(3000, lastRow - 1);
  const data = sh.getRange(2, 1, scan, LEDGER_LAST_COL).getValues(); // 01_모니터링현황은 보도일시 내림차순 정렬 — 최신은 상단
  const since = new Date(Date.now() - 12 * 3600 * 1000);
  const alreadySent = getArticleIdsInAlertLog_();

  const targets = data.filter(r =>
    r[LEDGER_COL.위기등급 - 1] === '주의' &&
    r[LEDGER_COL.수집일시 - 1] instanceof Date && r[LEDGER_COL.수집일시 - 1] >= since &&
    alreadySent.indexOf(r[LEDGER_COL.기사ID - 1]) < 0);
  if (!targets.length) return;

  const recipients = getRecipientsForGrade_('주의');
  const body = '[언론 모니터링 야간 요약] 주의 등급 ' + targets.length + '건\n\n' +
    targets.map(r => '▪ ' + r[LEDGER_COL.매체명 - 1] + ' | ' + r[LEDGER_COL.제목 - 1] + '\n  ' + r[LEDGER_COL.URL - 1]).join('\n\n');

  recipients.forEach(rec => {
    const contact = String(rec['연락처'] || '').trim();
    const result = contact.indexOf('@') > 0
      ? (() => { try { MailApp.sendEmail({ to: contact, subject: '[언론 모니터링 야간 요약] 주의 등급 ' + targets.length + '건', body: body }); return { ok: true }; } catch (e) { return { ok: false, error: e.message }; } })()
      : sendSms_(normalizePhone_(contact), body.substring(0, 2000));
    targets.forEach(r => {
      appendAlertLogRaw_(r[LEDGER_COL.기사ID - 1], r[LEDGER_COL.이슈그룹ID - 1], '주의', '',
        rec['이름'], contact.indexOf('@') > 0 ? '이메일(익일일괄)' : '알림톡(익일일괄)', result.ok ? '성공' : '실패: ' + result.error, body);
    });
  });
}

/** 트리거: 매일 08:30. '관심' 등급을 일일 요약 메일로 발송한다(6.4항). */
function sendDailyInterestSummary_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const scan = Math.min(3000, lastRow - 1);
  const data = sh.getRange(2, 1, scan, LEDGER_LAST_COL).getValues(); // 01_모니터링현황은 보도일시 내림차순 정렬 — 최신은 상단
  const since = new Date(Date.now() - 24 * 3600 * 1000);
  const alreadySent = getArticleIdsInAlertLog_();

  const targets = data.filter(r =>
    r[LEDGER_COL.위기등급 - 1] === '관심' &&
    r[LEDGER_COL.수집일시 - 1] instanceof Date && r[LEDGER_COL.수집일시 - 1] >= since &&
    alreadySent.indexOf(r[LEDGER_COL.기사ID - 1]) < 0);
  if (!targets.length) return;

  const recipients = getRecipientsForGrade_('관심').map(r => r['연락처']).filter(v => String(v).indexOf('@') > 0);
  const htmlBody = '<p>[언론 모니터링] 관심 등급 일일 요약 — ' + targets.length + '건</p><ul>' +
    targets.map(r => '<li>' + r[LEDGER_COL.매체명 - 1] + ' | ' + r[LEDGER_COL.제목 - 1] +
      ' — <a href="' + r[LEDGER_COL.URL - 1] + '">기사보기</a></li>').join('') + '</ul>';

  if (recipients.length) {
    MailApp.sendEmail({ to: recipients.join(','), subject: '[모니터링] 관심 등급 일일 요약 (' + nowStr_() + ')', htmlBody: htmlBody });
  }
  targets.forEach(r => {
    appendAlertLogRaw_(r[LEDGER_COL.기사ID - 1], r[LEDGER_COL.이슈그룹ID - 1], '관심', 0,
      recipients.join(', '), '일일요약', recipients.length ? '성공' : '실패: 수신자 없음', htmlBody);
  });
}

/** 07_알림이력은 최신순(내림차순)으로 위쪽에 쌓이므로 2행부터 스캔한다. */
function getArticleIdsInAlertLog_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ALERT_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const scan = Math.min(5000, lastRow - 1);
  return sh.getRange(2, 3, scan, 1).getValues().map(r => r[0]);
}

function appendAlertLogRaw_(articleId, issueGroupId, grade, score, recipientLabel, channel, resultLabel, message) {
  prependAlertLogRow_([
    'AL-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000),
    new Date(), articleId, issueGroupId, grade, score, recipientLabel, channel, resultLabel, message
  ]);
}

// ── F-509 수동 알림 / F-510 대응상태(시트 드롭다운) / F-512 오탐 학습 ─────

/** 메뉴: 01_모니터링현황에서 기사 행을 선택한 상태로 실행하면 해당 기사로 즉시 알림을 발송한다. */
function sendManualAlertForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== SHEETS.LEDGER) {
    ui.alert('01_모니터링현황 시트에서 기사 행을 클릭한 뒤 다시 실행해 주세요.');
    return;
  }
  const row = activeSheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('기사 데이터 행을 클릭한 뒤 다시 실행해 주세요.'); return; }
  const sh = activeSheet;
  const values = sh.getRange(row, 1, 1, LEDGER_LAST_COL).getValues()[0];
  const grade = values[LEDGER_COL.위기등급 - 1] || '주의';
  const effectiveGrade = grade === '해당없음' ? '주의' : grade;
  const candidate = {
    articleId: values[LEDGER_COL.기사ID - 1],
    issueGroupId: values[LEDGER_COL.이슈그룹ID - 1] || ('MANUAL-' + values[LEDGER_COL.기사ID - 1]),
    title: values[LEDGER_COL.제목 - 1],
    link: values[LEDGER_COL.URL - 1],
    pubDate: values[LEDGER_COL.보도일시 - 1],
    media: { 표준매체명: values[LEDGER_COL.매체명 - 1], 매체영향력: values[LEDGER_COL.매체영향력 - 1] },
    crisis: { grade: effectiveGrade, score: 0, reasonText: '담당자 수동 발송' }
  };
  dispatchAlert_(candidate, effectiveGrade);
  ui.alert('수동 알림 발송을 시도했습니다. 07_알림이력 시트에서 결과를 확인하세요.');
}

/** 메뉴: 01_모니터링현황에서 오탐인 기사 행을 선택하고 실행하면 상태를 '오탐'으로 바꾸고 매칭 키워드를 예외 목록에 축적한다. */
function markSelectedRowFalsePositive() {
  const ui = SpreadsheetApp.getUi();
  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== SHEETS.LEDGER) {
    ui.alert('01_모니터링현황 시트에서 기사 행을 클릭한 뒤 다시 실행해 주세요.');
    return;
  }
  const row = activeSheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('기사 데이터 행을 클릭한 뒤 다시 실행해 주세요.'); return; }
  const sh = activeSheet;

  const values = sh.getRange(row, 1, 1, LEDGER_LAST_COL).getValues()[0];
  const title = values[LEDGER_COL.제목 - 1];
  const summary = values[LEDGER_COL.요약 - 1];
  const text = title + ' ' + summary;

  const crisisDict = readConfigBlock_('위기키워드');
  const orgPersonKeywords = readConfigBlock_('키워드')
    .filter(k => k['구분'] === '기관명' || k['구분'] === '상위기관' || k['구분'] === '인물')
    .map(k => k['키워드']);
  const matched = matchCrisisKeywords_(text, crisisDict, orgPersonKeywords, []);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  matched.forEach(m => appendToConfigBlock_('오탐예외', [m.keyword, today, title]));

  sh.getRange(row, LEDGER_COL.상태).setValue('오탐');
  ui.alert('오탐 처리 완료. 예외 키워드 ' + matched.length + '건이 00_설정!오탐예외에 축적되었습니다.');
}
