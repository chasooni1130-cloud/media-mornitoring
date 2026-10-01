/**
 * 05_Pipeline.gs — 수집 오케스트레이션 + 정제 파이프라인
 * F-201(중복제거) F-202(이슈그룹) F-203(매체정규화) F-204(오수집필터)
 * F-205(논조1차판정) F-206(카테고리분류) F-207(진료과태깅) F-301(모니터링현황 적재)
 *
 * 트리거·메뉴에서 호출하는 최상위 진입점은 collectAll() 하나뿐이다.
 */

function collectAll() {
  if (!shouldRunNow_()) return;
  runCollection_();
}

function runCollection_() {
  const startTime = Date.now();
  let raw = [];

  try { raw = raw.concat(collectNaverNews_()); }
  catch (e) { writeLog_('collectAll:news', 0, 0, e.message, 0); }

  try { raw = raw.concat(collectRSS_()); }
  catch (e) { writeLog_('collectAll:rss', 0, 0, e.message, 0); }

  const th = getThresholds_();
  if (String(th['커뮤니티수집_사용']).toUpperCase() === 'Y') {
    try { raw = raw.concat(collectNaverCommunity_()); }
    catch (e) { writeLog_('collectAll:community', 0, 0, e.message, 0); }
  }

  const savedCount = processAndAppend_(raw);
  setLastCollectTs_(new Date());
  writeLog_('collectAll', raw.length, savedCount, '', Date.now() - startTime);
  checkApiQuota_();
  checkStaleCollection_();
}

/**
 * 원시 수집 결과를 정제 → 위기판정 → 01_모니터링현황 적재까지 한 번에 처리한다.
 * 반환값: 실제로 모니터링 현황에 새로 추가된 행 수.
 */
function processAndAppend_(rawItems) {
  if (!rawItems.length) return 0;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  const context = getRecentLedgerContext_(sh, lastRow); // 최근 최대 3000행 컨텍스트 (이슈그룹/확산판정용 — 72시간 창이라 최근분만으로 충분)
  const existingKeys = new Set(context.map(c => c.dedupKey));
  getPermanentDedupKeys_().forEach(k => existingKeys.add(k)); // F-201: 01_모니터링현황에서 지워진 옛 기사도 재적재되지 않도록 영구 이력과 합친다
  const existingIds = context.map(c => c.articleId);

  const excludeWords = readConfigBlock_('제외어').filter(e => e['사용여부'] === 'Y' || e['사용여부'] === true).map(e => e['제외어']);
  const orgKeywords = readConfigBlock_('키워드').filter(k => k['구분'] === '기관명').map(k => k['키워드']);
  const mediaMap = buildMediaMap_();
  const toneDict = readConfigBlock_('논조키워드');
  const categoryDict = readConfigBlock_('카테고리키워드');
  const deptDict = readConfigBlock_('진료과키워드');
  const crisisDict = readConfigBlock_('위기키워드');
  const orgPersonKeywords = readConfigBlock_('키워드')
    .filter(k => k['구분'] === '기관명' || k['구분'] === '상위기관' || k['구분'] === '인물')
    .map(k => k['키워드']);
  const fpExceptions = readConfigBlock_('오탐예외').map(r => r['키워드']).filter(String);
  const staffDict = readConfigBlock_('의료진사전');
  const simThreshold = Number(getThresholds_()['이슈그룹_유사도_임계']) || 0.62;
  const th = getThresholds_();

  let seq = nextSeqForToday_(existingIds);
  const batchGroupTouch = {}; // issueGroupId -> 이번 배치에서 새로 생긴 항목 수(급증 판정용)
  const newRows = [];
  const newHistoryRows = []; // 09_수집이력에 함께 적재할 [기사ID, dedupKey, 보도일시, 수집일시]
  const crisisCandidates = []; // 알림 대상 후보 {rowData, grade, score, reasons}

  rawItems.forEach(item => {
    if (!item.title || !item.link) return;
    const key = dedupKey_(item.link, item.rawMedia, item.title);
    if (existingKeys.has(key)) return; // F-201 중복 제거
    existingKeys.add(key);

    const media = normalizeMedia_(item, mediaMap);
    if (EXCLUDED_MEDIA.indexOf(media.표준매체명) >= 0) return; // 사용자 요청: 이 매체는 아예 적재하지 않는다
    if (media.도메인제외) return; // 사용자 요청(2026-10): 매핑되지 않은 영문 도메인 매체는 수집하지 않는다

    const bodyText = item.title + ' ' + (item.description || '');
    const excluded = isExcluded_(bodyText, excludeWords, orgKeywords);

    const tone = classifyTone_(bodyText, toneDict);
    const category = classifyCategory_(bodyText, categoryDict);
    const depts = classifyDepartments_(bodyText, deptDict);
    const staff = classifyMedicalStaff_(bodyText, staffDict);

    const issueGroupId = assignIssueGroup_(item.title, item.pubDate, context, simThreshold);
    batchGroupTouch[issueGroupId] = (batchGroupTouch[issueGroupId] || 0) + 1;

    const crisis = evaluateCrisis_(bodyText, media, issueGroupId, item.sourceType, context, batchGroupTouch,
      crisisDict, orgPersonKeywords, th, fpExceptions);

    const articleId = generateArticleId_(seq++);
    const row = new Array(LEDGER_LAST_COL).fill('');
    row[LEDGER_COL.기사ID - 1] = articleId;
    row[LEDGER_COL.수집일시 - 1] = new Date();
    row[LEDGER_COL.보도일시 - 1] = item.pubDate || new Date();
    row[LEDGER_COL.매체명 - 1] = media.표준매체명;
    row[LEDGER_COL.매체유형 - 1] = media.매체유형;
    row[LEDGER_COL.매체영향력 - 1] = media.매체영향력;
    row[LEDGER_COL.기자명 - 1] = '';
    row[LEDGER_COL.제목 - 1] = titleHyperlinkFormula_(item.title, item.link);
    row[LEDGER_COL.URL - 1] = item.link;
    row[LEDGER_COL.요약 - 1] = truncateSummary_(item.description, 220);
    row[LEDGER_COL.검색키워드 - 1] = item.keyword || '';
    row[LEDGER_COL.카테고리 - 1] = category;
    row[LEDGER_COL.진료과센터 - 1] = depts;
    row[LEDGER_COL.논조자동 - 1] = tone;
    row[LEDGER_COL.논조확정 - 1] = '';
    row[LEDGER_COL.위기등급 - 1] = crisis.grade;
    row[LEDGER_COL.이슈그룹ID - 1] = issueGroupId;
    row[LEDGER_COL.보도유형 - 1] = '';
    row[LEDGER_COL.보도자료ID - 1] = '';
    row[LEDGER_COL.노출도 - 1] = '';
    row[LEDGER_COL.수집경로 - 1] = item.collectMethod || 'API';
    row[LEDGER_COL.상태 - 1] = excluded ? '제외' : '미확인';
    row[LEDGER_COL.대응상태 - 1] = crisis.grade === '해당없음' ? '해당없음' : '확인';
    row[LEDGER_COL.조치내역 - 1] = '';
    row[LEDGER_COL.담당자 - 1] = '';
    row[LEDGER_COL.비고 - 1] = media.신규매체 ? '신규매체 - 확인필요' : '';
    row[LEDGER_COL.진료과 - 1] = staff.dept;
    row[LEDGER_COL.의료진 - 1] = staff.staff;

    newRows.push(row);
    newHistoryRows.push([articleId, key, item.pubDate || new Date(), new Date()]);
    context.push({
      articleId: articleId, title: item.title, pubDate: item.pubDate || new Date(),
      collectedAt: new Date(), issueGroupId: issueGroupId, dedupKey: key,
      sourceType: item.sourceType, mediaInfluence: media.매체영향력
    });

    if (!excluded && crisis.grade !== '해당없음') {
      crisisCandidates.push({ articleId, row, media, crisis, issueGroupId, link: item.link, title: item.title, pubDate: item.pubDate });
    }
  });

  if (!newRows.length) return 0;
  sh.getRange(lastRow + 1, 1, newRows.length, LEDGER_LAST_COL).setValues(newRows);
  sortLedgerByReportDate_(sh);
  appendDedupHistory_(newHistoryRows);

  if (crisisCandidates.length) {
    evaluateAndSendAlerts_(crisisCandidates);
  }
  return newRows.length;
}

/** 사용자 요청: 01_모니터링현황은 항상 보도일시 최신순(내림차순)으로 정렬된 상태를 유지한다. */
function sortLedgerByReportDate_(sh) {
  const dataRows = sh.getLastRow() - 1;
  if (dataRows > 1) {
    sh.getRange(2, 1, dataRows, LEDGER_LAST_COL).sort({ column: LEDGER_COL.보도일시, ascending: false });
  }
}

/** 최근 N일치 모니터링 현황 데이터를 컨텍스트 배열로 로드 (중복/이슈그룹/확산판정용). 과도한 스캔을 막기 위해 최대 3000행으로 제한.
 *  01_모니터링현황은 보도일시 내림차순으로 정렬되어 있으므로 최신 데이터는 상단(2행부터)에 있다. */
function getRecentLedgerContext_(sh, lastRow) {
  if (lastRow < 2) return [];
  const maxScan = 3000;
  const numRows = Math.min(maxScan, lastRow - 1);
  const cols = [LEDGER_COL.기사ID, LEDGER_COL.수집일시, LEDGER_COL.보도일시, LEDGER_COL.매체명,
    LEDGER_COL.매체영향력, LEDGER_COL.URL, LEDGER_COL.제목, LEDGER_COL.이슈그룹ID, LEDGER_COL.수집경로];
  const minCol = Math.min.apply(null, cols);
  const maxCol = Math.max.apply(null, cols);
  const values = sh.getRange(2, minCol, numRows, maxCol - minCol + 1).getValues();
  // 주의: 여기서 날짜로 걸러내지 않는다 — 예전엔 "days일 이전 제외"를 했다가, 오래된 기사가
  // 검색 결과에 계속 다시 잡힐 때 중복방지 키(dedupKey) 목록에서도 함께 빠져버려서
  // 매 수집마다 같은 기사가 무한히 재적재되는 사고가 있었다(F-201 실패).
  // 이슈그룹 판단(assignIssueGroup_)은 자체적으로 72시간 창을 다시 확인하므로 여기서 막지 않아도 안전하다.
  // (참고) 01_모니터링현황 자체에서 옛 행을 삭제해도 dedupKey는 09_수집이력에 영구 보관되므로
  // (getPermanentDedupKeys_), 화면에서 지운 기사가 재수집으로 되살아나는 사고는 이제 이 컨텍스트와 무관하게 방지된다.
  const out = [];
  values.forEach(r => {
    const get = (col) => r[col - minCol];
    const pubDate = get(LEDGER_COL.보도일시);
    out.push({
      articleId: get(LEDGER_COL.기사ID),
      collectedAt: get(LEDGER_COL.수집일시),
      pubDate: pubDate,
      title: get(LEDGER_COL.제목),
      issueGroupId: get(LEDGER_COL.이슈그룹ID),
      mediaInfluence: get(LEDGER_COL.매체영향력),
      sourceType: get(LEDGER_COL.수집경로) === '커뮤니티' ? '커뮤니티' : 'news',
      dedupKey: dedupKey_(get(LEDGER_COL.URL), get(LEDGER_COL.매체명), get(LEDGER_COL.제목))
    });
  });
  return out;
}

/**
 * 09_수집이력(절대 삭제하지 않는 dedupKey 영구 저장소)에서 전체 dedupKey를 읽어온다.
 * 01_모니터링현황에서 지운 옛 기사라도 이 목록에 남아있으면 F-201 중복판정에서 계속 걸러진다.
 */
function getPermanentDedupKeys_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.HISTORY);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 2, lastRow - 1, 1).getValues().map(r => r[0]).filter(String);
}

/** 새로 적재된 기사의 dedupKey를 09_수집이력에 이어붙인다(never trimmed). */
function appendDedupHistory_(rows) {
  if (!rows.length) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.HISTORY);
  if (!sh) return; // 마이그레이션 전(예: 초기 설정 미실행) 환경에서도 수집 자체는 막지 않는다
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HISTORY_HEADERS.length).setValues(rows);
}

function buildMediaMap_() {
  const rows = readConfigBlock_('매체마스터');
  const map = {};
  rows.forEach(r => { if (r['수집매체명']) map[String(r['수집매체명']).trim().toLowerCase()] = r; });
  return map;
}

/** F-203 매체 정규화. 매치 실패 시 표준매체명=원문, 신규매체 플래그를 세운다. */
function normalizeMedia_(item, mediaMap) {
  const key = String(item.rawMedia || '').trim().toLowerCase();
  const found = key && mediaMap[key];
  if (found) {
    return { 표준매체명: found['표준매체명'] || item.rawMedia, 매체유형: found['매체유형'] || '', 매체영향력: found['매체영향력'] || '', 신규매체: false };
  }
  if (item.sourceType === '커뮤니티') {
    return { 표준매체명: item.rawMedia || '커뮤니티', 매체유형: '커뮤니티·SNS', 매체영향력: '커뮤니티', 신규매체: false };
  }
  const fb = key && DOMAIN_MEDIA_FALLBACK[key];
  if (fb) {
    return { 표준매체명: fb.name, 매체유형: fb.type, 매체영향력: fb.influence, 신규매체: false };
  }
  // 사용자 요청(2026-10): 매체마스터/DOMAIN_MEDIA_FALLBACK 어디에도 없고 영문 도메인(URL) 형태 그대로인
  // 매체는 "신규매체 - 확인필요"로 쌓아두지 않고 아예 적재하지 않는다(processAndAppend_의 도메인제외 분기 참고).
  // 피드명이 이미 한글 등으로 들어온 경우(도메인 형태가 아님)는 기존대로 신규매체 검토 대상으로 남긴다.
  if (key && isDomainLike_(key)) {
    return { 표준매체명: item.rawMedia, 매체유형: '', 매체영향력: '', 신규매체: true, 도메인제외: true };
  }
  return { 표준매체명: item.rawMedia || '(미상)', 매체유형: '', 매체영향력: '', 신규매체: true };
}

/** 문자열이 "news.sbs.co.kr" 같은 영문 도메인(URL 호스트) 형태인지 판정한다. */
function isDomainLike_(s) {
  return /^(?:[a-z0-9-]+\.)+[a-z]{2,}$/i.test(s);
}

/** F-204: 제외어가 있고 기관명이 본문에 없으면 오수집으로 판정. */
function isExcluded_(text, excludeWords, orgKeywords) {
  const hasExclude = excludeWords.some(w => w && text.indexOf(w) >= 0);
  if (!hasExclude) return false;
  const hasOrg = orgKeywords.some(w => w && text.indexOf(w) >= 0);
  return !hasOrg;
}

/** F-205 논조 1차 판정: 긍/부정 키워드 매칭 수 비교. 동률·매칭없음이면 중립. */
function classifyTone_(text, toneDict) {
  let pos = 0, neg = 0;
  toneDict.forEach(r => {
    if (!r['키워드'] || text.indexOf(r['키워드']) < 0) return;
    if (r['구분(긍정/부정)'] === '긍정') pos++; else if (r['구분(긍정/부정)'] === '부정') neg++;
  });
  if (pos === 0 && neg === 0) return '중립';
  return pos > neg ? '긍정' : (neg > pos ? '부정' : '중립');
}

/** F-206 카테고리 분류: 카테고리별 키워드 매칭 수가 가장 많은 카테고리를 채택. */
function classifyCategory_(text, categoryDict) {
  const counts = {};
  categoryDict.forEach(r => {
    if (!r['키워드'] || text.indexOf(r['키워드']) < 0) return;
    counts[r['카테고리']] = (counts[r['카테고리']] || 0) + 1;
  });
  let best = '', bestCount = 0;
  Object.keys(counts).forEach(cat => { if (counts[cat] > bestCount) { best = cat; bestCount = counts[cat]; } });
  return best || '기타';
}

/** F-207 진료과·센터 태깅: 매칭되는 모든 진료과를 쉼표로 나열. */
function classifyDepartments_(text, deptDict) {
  const set = new Set();
  deptDict.forEach(r => { if (r['키워드'] && text.indexOf(r['키워드']) >= 0) set.add(r['진료과·센터']); });
  return Array.from(set).join(', ');
}

/**
 * 사용자 요청: 기사 제목·요약에 00_설정!의료진 사전에 등록된 이름이 나오면 "소속 이름" 형태로 나열.
 * 이름 인식은 NLP가 아니라 담당자가 등록한 사전과의 단순 텍스트 매칭이므로,
 * 사전에 없는 인물은 잡히지 않는다 — 정확도를 높이려면 사전에 이름을 계속 추가해야 한다.
 */
function classifyMedicalStaff_(text, staffDict) {
  const depts = new Set();
  const names = new Set();
  staffDict.forEach(r => {
    const name = String(r['이름'] || '').trim();
    if (!name || text.indexOf(name) < 0) return;
    const dept = String(r['소속'] || '').trim();
    names.add(name);
    if (dept) depts.add(dept);
  });
  return { dept: Array.from(depts).join(', '), staff: Array.from(names).join(', ') };
}

/** F-202 이슈 그룹 배정: 최근 컨텍스트에서 제목 유사도가 임계 이상이고 보도일시가 72시간 이내인 항목의 그룹에 편입. */
function assignIssueGroup_(title, pubDate, context, threshold) {
  const pd = pubDate instanceof Date ? pubDate : new Date(pubDate || Date.now());
  const windowMs = 72 * 3600 * 1000;
  let best = null, bestSim = 0;
  for (let i = context.length - 1; i >= 0; i--) {
    const c = context[i];
    if (!c.issueGroupId) continue;
    const cPd = c.pubDate instanceof Date ? c.pubDate : new Date(c.pubDate || 0);
    if (Math.abs(pd - cPd) > windowMs) continue;
    const sim = titleSimilarity_(title, c.title);
    if (sim >= threshold && sim > bestSim) { best = c.issueGroupId; bestSim = sim; }
  }
  if (best) return best;
  const tz = Session.getScriptTimeZone();
  return 'IG-' + Utilities.formatDate(pd, tz, 'yyyyMMdd') + '-' + md5_(normalizeTitle_(title)).substring(0, 6);
}

// ── 수집 스케줄 자기-제어 (F-105) ───────────────────────────────────────

function setLastCollectTs_(d) {
  PropertiesService.getScriptProperties().setProperty('LAST_COLLECT_TS', String(d.getTime()));
}
function getLastCollectTs_() {
  const v = PropertiesService.getScriptProperties().getProperty('LAST_COLLECT_TS');
  return v ? new Date(Number(v)) : new Date(0);
}

/**
 * 시간 기반 트리거는 5분 간격으로 collectAll()을 깨우지만, 실제 수집은
 * F-105 규칙(평일 주간 기본/야간·주말 완화/위기 시 전 시간대 단축)에 맞춰 자체적으로 걸러낸다.
 */
function shouldRunNow_() {
  const th = getThresholds_();
  const intervalMin = getCollectIntervalMinutes_(th);
  const elapsedMin = (Date.now() - getLastCollectTs_().getTime()) / 60000;
  return elapsedMin >= intervalMin;
}

function getCollectIntervalMinutes_(th) {
  if (isForceRelaxed_()) return Number(th['수집주기_야간주말_분']) || 45; // F-107: 일일 한도 100% 도달 시 자동 완화
  if (isCrisisModeActive_()) return Number(th['수집주기_위기_분']) || 5;
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const day = Number(Utilities.formatDate(now, tz, 'u')); // 1=Mon..7=Sun
  const hour = Number(Utilities.formatDate(now, tz, 'H'));
  const isWeekend = (day === 6 || day === 7);
  const isDaytime = hour >= 7 && hour < 22;
  if (!isWeekend && isDaytime) return Number(th['수집주기_기본_분']) || 15;
  return Number(th['수집주기_야간주말_분']) || 45;
}

/** 최근 위기모드해제_시간 내에 경계/심각 등급이면서 미종결인 건이 있으면 위기 모드로 간주.
 *  01_모니터링현황은 보도일시 내림차순 정렬 상태이므로 최신 데이터는 상단(2행부터)에 있다. */
function isCrisisModeActive_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const th = getThresholds_();
  const hoursWindow = Number(th['위기모드_해제_시간']) || 12;
  const scanRows = Math.min(1000, lastRow - 1);
  const data = sh.getRange(2, 1, scanRows, LEDGER_LAST_COL).getValues();
  const cutoff = new Date(Date.now() - hoursWindow * 3600 * 1000);
  return data.some(r => {
    const grade = r[LEDGER_COL.위기등급 - 1];
    const status = r[LEDGER_COL.대응상태 - 1];
    const collectedAt = r[LEDGER_COL.수집일시 - 1];
    return (grade === '경계' || grade === '심각') && status !== '종결' &&
      collectedAt instanceof Date && collectedAt >= cutoff;
  });
}

/**
 * 유지보수용 메뉴 함수. 00_설정!매체마스터 또는 DOMAIN_MEDIA_FALLBACK 사전을 갱신한 뒤,
 * 이미 모니터링 현황에 쌓인 기존 행들의 매체명·매체유형·매체영향력을 URL 기준으로 다시 계산하고
 * 보도일시 최신순 정렬도 함께 맞춘다.
 */
/**
 * 사용자 요청: 이미 '네이버뉴스'로 분류된 기존 행들의 실제 매체명을 다시 확인한다.
 * 네이버 뉴스 미러 링크(URL 컬럼)만으로는 원 매체를 알 수 없으므로, 저장된 제목으로
 * 네이버 뉴스 검색을 재호출해 같은 기사(코드+글번호 일치)를 찾고, 그 결과의 originallink(원문 링크)
 * 에서 실제 매체를 재계산한다. HTML을 직접 파싱하는 방식은 네이버 페이지가 클라이언트 렌더링이라
 * 신뢰할 수 없어(원문 링크 비노출 매체도 있음) API 재조회 방식으로 전환했다.
 */
function backfillNaverOriginalMedia() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { alertSafe_('모니터링 현황에 데이터가 없습니다.'); return; }
  const range = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL);
  const data = range.getValues();
  const mediaMap = buildMediaMap_();
  let updated = 0, notFound = 0;
  const notFoundSamples = [];

  data.forEach(row => {
    if (row[LEDGER_COL.매체명 - 1] !== '네이버뉴스') return;
    const url = String(row[LEDGER_COL.URL - 1] || '');
    const codeMatch = url.match(/\/article\/(\d+)\/(\d+)/);
    if (!codeMatch) { notFound++; return; }
    const code = codeMatch[1], aid = codeMatch[2];
    const title = row[LEDGER_COL.제목 - 1];

    const result = fetchNaverNewsOnce_(title, 10);
    const match = !result.error && result.items.find(it => String(it.link || '').indexOf('/article/' + code + '/' + aid) >= 0);
    const rawMedia = match && match.originallink ? guessMediaFromLink_(match.originallink) : '';

    if (!rawMedia || rawMedia === 'n.news.naver.com' || rawMedia === 'news.naver.com') {
      notFound++;
      if (notFoundSamples.length < 15) notFoundSamples.push('code=' + code + ' ' + title);
      return;
    }
    const media = normalizeMedia_({ rawMedia: rawMedia, sourceType: 'news' }, mediaMap);
    row[LEDGER_COL.매체명 - 1] = media.표준매체명;
    row[LEDGER_COL.매체유형 - 1] = media.매체유형;
    row[LEDGER_COL.매체영향력 - 1] = media.매체영향력;
    if (media.신규매체 && !row[LEDGER_COL.비고 - 1]) row[LEDGER_COL.비고 - 1] = '신규매체 - 확인필요';
    updated++;
  });

  range.setValues(data);
  sortLedgerByReportDate_(sh);
  writeLog_('backfillNaverOriginalMedia', updated + notFound, updated, notFoundSamples.join(' | '), 0);
  alertSafe_('네이버뉴스 매체명 재확인 완료 — 갱신 ' + updated + '건, 확인 불가(원문 링크 비공개 등) ' + notFound + '건. 상세는 08_로그 시트 참고.');
}

function renormalizeLedgerMedia() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('모니터링 현황에 데이터가 없습니다.');
    return;
  }
  const mediaMap = buildMediaMap_();
  const range = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL);
  const data = range.getValues();
  data.forEach(row => {
    const url = row[LEDGER_COL.URL - 1];
    const rawMedia = guessMediaFromLink_(url);
    const sourceType = row[LEDGER_COL.수집경로 - 1] === '커뮤니티' ? '커뮤니티' : 'news';
    const media = normalizeMedia_({ rawMedia: rawMedia, sourceType: sourceType }, mediaMap);
    row[LEDGER_COL.매체명 - 1] = media.표준매체명;
    row[LEDGER_COL.매체유형 - 1] = media.매체유형;
    row[LEDGER_COL.매체영향력 - 1] = media.매체영향력;
    if (media.신규매체 && !row[LEDGER_COL.비고 - 1]) row[LEDGER_COL.비고 - 1] = '신규매체 - 확인필요';
  });
  range.setValues(data);
  sortLedgerByReportDate_(sh);
  SpreadsheetApp.getUi().alert('매체 정보 재계산 완료', data.length + '건을 갱신했습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 유지보수용 메뉴 함수. 00_설정!의료진 사전을 채우거나 수정한 뒤 실행하면,
 * 이미 쌓인 기존 행의 제목·요약을 다시 검사해 의료진 컬럼을 갱신한다(재수집 불필요).
 */
function backfillMedicalStaff() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('모니터링 현황에 데이터가 없습니다.');
    return;
  }
  const staffDict = readConfigBlock_('의료진사전');
  if (!staffDict.length) {
    SpreadsheetApp.getUi().alert('00_설정!의료진 사전이 비어 있습니다. 이름·소속을 먼저 입력해 주세요.');
    return;
  }
  const range = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL);
  const data = range.getValues();
  data.forEach(row => {
    const bodyText = row[LEDGER_COL.제목 - 1] + ' ' + (row[LEDGER_COL.요약 - 1] || '');
    const staff = classifyMedicalStaff_(bodyText, staffDict);
    row[LEDGER_COL.진료과 - 1] = staff.dept;
    row[LEDGER_COL.의료진 - 1] = staff.staff;
  });
  range.setValues(data);
  SpreadsheetApp.getUi().alert('의료진 태그 재계산 완료', data.length + '건을 검사했습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}
