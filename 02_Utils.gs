/**
 * 02_Utils.gs — 공용 유틸리티
 */

/** 기사ID 생성: 수집일시+일련번호 (예: 20260820-0031). 같은 날 여러 건이면 일련번호 증가. */
function generateArticleId_(seq) {
  const tz = Session.getScriptTimeZone();
  const ymd = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  return ymd + '-' + String(seq).padStart(4, '0');
}

/** 오늘 날짜(yyyyMMdd) 기준으로 모니터링 현황 시트에 이미 쓰인 마지막 일련번호+1을 반환. */
function nextSeqForToday_(existingIds) {
  const tz = Session.getScriptTimeZone();
  const todayPrefix = Utilities.formatDate(new Date(), tz, 'yyyyMMdd') + '-';
  let max = 0;
  existingIds.forEach(id => {
    if (typeof id === 'string' && id.indexOf(todayPrefix) === 0) {
      const n = parseInt(id.substring(todayPrefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max + 1;
}

/** URL과 (매체명+제목)을 결합한 중복판정용 해시 (F-201). */
function dedupKey_(url, mediaName, title) {
  const normUrl = normalizeUrl_(url);
  if (normUrl) return 'U:' + normUrl;
  const raw = (mediaName || '') + '|' + normalizeTitle_(title || '');
  return 'T:' + md5_(raw);
}

/** 추적 파라미터 제거 등 URL 정규화 (?utm_ 등 제거, 트레일링 슬래시 제거). */
function normalizeUrl_(url) {
  if (!url) return '';
  try {
    let u = String(url).trim();
    u = u.split('#')[0];
    const qIdx = u.indexOf('?');
    if (qIdx >= 0) {
      const base = u.substring(0, qIdx);
      const query = u.substring(qIdx + 1);
      const keep = query.split('&').filter(p => p && !/^(utm_|fbclid|gclid|ref=)/i.test(p));
      u = keep.length ? base + '?' + keep.join('&') : base;
    }
    return u.replace(/\/+$/, '').toLowerCase();
  } catch (e) {
    return String(url).toLowerCase();
  }
}

/** 제목 비교용 정규화: 태그·특수문자·공백 제거. */
function normalizeTitle_(title) {
  return String(title)
    .replace(/<[^>]+>/g, '')
    .replace(/[\s　]+/g, '')
    .replace(/[""''「」『』【】\[\]()·,.!?…\-–—:;"']/g, '')
    .toLowerCase();
}

function md5_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

/**
 * 두 제목의 유사도(0~1) — 문자 2-gram 기반 Dice 계수.
 * 통신사 전재로 동일 기사가 여러 매체에 실리는 경우(F-202) 묶는 데 사용.
 */
function titleSimilarity_(a, b) {
  const na = normalizeTitle_(a);
  const nb = normalizeTitle_(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  const gramsA = bigrams_(na);
  const gramsB = bigrams_(nb);
  if (!gramsA.size || !gramsB.size) return 0;
  let common = 0;
  gramsA.forEach(g => { if (gramsB.has(g)) common++; });
  return (2 * common) / (gramsA.size + gramsB.size);
}

function bigrams_(s) {
  const set = new Set();
  for (let i = 0; i < s.length - 1; i++) set.add(s.substring(i, i + 2));
  if (set.size === 0 && s.length) set.add(s);
  return set;
}

/** HTML 태그 제거 + 공백 정리 (RSS description 등에서 요약 추출). */
function stripHtml_(html) {
  if (!html) return '';
  return String(html)
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 사용자 요청(2026-10): URL 컬럼을 따로 두지 않고, 제목 셀 자체를 클릭하면 원문으로 이동하도록
 * =HYPERLINK() 수식을 만든다. getValues()로 읽으면 수식이 아니라 표시되는 제목 텍스트가 그대로
 * 반환되므로, 제목 기반 매칭(이슈그룹·중복판정·보도자료 매칭 등) 로직은 그대로 동작한다.
 */
function titleHyperlinkFormula_(title, url) {
  const t = String(title || '').replace(/"/g, '""');
  if (!url) return t;
  const u = String(url).replace(/"/g, '""');
  return '=HYPERLINK("' + u + '","' + t + '")';
}

/** 본문/요약을 2~3문장(약 200자) 이내로 절단 — 전문 저장 금지 원칙(저작권). */
function truncateSummary_(text, maxLen) {
  maxLen = maxLen || 220;
  const clean = stripHtml_(text);
  if (clean.length <= maxLen) return clean;
  const cut = clean.substring(0, maxLen);
  const lastSentenceEnd = Math.max(cut.lastIndexOf('.'), cut.lastIndexOf('다.'), cut.lastIndexOf('요.'));
  return (lastSentenceEnd > maxLen * 0.5 ? cut.substring(0, lastSentenceEnd + 1) : cut) + '…';
}

/** 08_로그 시트에 실행 기록 (F-106). */
function writeLog_(task, callCount, savedCount, errorMsg, elapsedMs) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (!sh) return;
  sh.appendRow([new Date(), task, callCount || 0, savedCount || 0, errorMsg || '', elapsedMs || 0]);
}

/** 담당자 알림용 이메일 발송 (장애·API 한도·메시지 잔액 등 운영 알림). */
function notifyOps_(subject, body) {
  const receivers = readConfigBlock_('수신자')
    .filter(r => r['수신동의'] === 'Y' || r['수신동의'] === true)
    .map(r => r['연락처'])
    .filter(v => typeof v === 'string' && v.indexOf('@') > 0);
  if (!receivers.length) return;
  try {
    MailApp.sendEmail(receivers.join(','), '[모니터링 시스템] ' + subject, body);
  } catch (e) {
    writeLog_('notifyOps_', 0, 0, 'Ops 메일 발송 실패: ' + e.message, 0);
  }
}

/** 날짜별 카운터 증가 (F-107 API 한도 관리 등). CacheService보다 영속적인 ScriptProperties 사용. */
function incrementDailyCounter_(key) {
  const props = PropertiesService.getScriptProperties();
  const dateKey = key + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const cur = Number(props.getProperty(dateKey) || 0) + 1;
  props.setProperty(dateKey, String(cur));
  return cur;
}
function getDailyCounter_(key) {
  const dateKey = key + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return Number(PropertiesService.getScriptProperties().getProperty(dateKey) || 0);
}

/** 월별 카운터 — NAVER API HUB는 일일이 아닌 월 단위 통합 한도로 관리되므로 F-107에서 사용. */
function incrementMonthlyCounter_(key) {
  const props = PropertiesService.getScriptProperties();
  const monthKey = key + '_M_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  const cur = Number(props.getProperty(monthKey) || 0) + 1;
  props.setProperty(monthKey, String(cur));
  return cur;
}
function getMonthlyCounter_(key) {
  const monthKey = key + '_M_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  return Number(PropertiesService.getScriptProperties().getProperty(monthKey) || 0);
}

/**
 * 스프레드시트 메뉴가 아니라 스크립트 편집기에서 직접 실행할 때는
 * SpreadsheetApp.getUi()가 예외를 던진다. 알림 자체가 목적이 아닌 함수에서
 * "결과를 보여줄 수 있으면 보여주고, 안 되면 조용히 넘어간다"용으로 쓴다.
 */
function alertSafe_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    writeLog_('alertSafe_', 0, 0, message, 0);
  }
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
