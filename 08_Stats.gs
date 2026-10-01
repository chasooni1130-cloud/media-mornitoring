/**
 * 08_Stats.gs — 통계 자동화 (F-401~405, 명세서 7장)
 * 주간/월간/연간 통계는 모니터링 현황(01_모니터링현황)의 상태=확정 행만 집계 대상으로 한다(7.1항).
 * 통계는 QUERY 수식 대신 스크립트가 직접 계산해 값으로 기록한다 —
 * 담당자가 수식을 이해하지 못해도 메뉴 클릭만으로 항상 같은 방식으로 재계산되도록 하기 위함이다.
 */

function buildWeeklyStats() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const day = Number(Utilities.formatDate(now, tz, 'u')); // 1=Mon .. 7=Sun
  const monday = addDays_(now, -(day - 1));
  const weekStart = stripTime_(monday);
  const weekEnd = addDays_(weekStart, 7);
  const prevStart = addDays_(weekStart, -7);
  const prevEnd = weekStart;

  const cur = computeLedgerStats_(weekStart, weekEnd);
  const prev = computeLedgerStats_(prevStart, prevEnd);
  const pressStats = computePressStats_(weekStart, weekEnd);

  const sh = getOrCreateSheet_(SHEETS.WEEKLY);
  sh.clear();
  let r = 1;
  r = writeTitle_(sh, r, '주간 통계 — ' + fmtDate_(weekStart) + ' ~ ' + fmtDate_(addDays_(weekEnd, -1)));

  r = writeKv_(sh, r, [
    ['총 보도 건수', cur.total],
    ['순 이슈 건수(중복제거)', cur.issueCount],
    ['전주 대비 증감률(%)', pctChange_(cur.total, prev.total)]
  ]);
  r++;

  r = writeTable_(sh, r, '매체유형별 분포', ['매체유형', '건수', '비율(%)'], distTable_(cur.byMediaType, cur.total));
  r = writeTable_(sh, r, '논조별 분포', ['논조', '건수', '비율(%)'], distTable_(cur.byTone, cur.total));
  r = writeTable_(sh, r, '카테고리별 분포(상위 3개 ★)', ['카테고리', '건수', '비율(%)'], distTable_(cur.byCategory, cur.total, 3));
  r = writeTable_(sh, r, '진료과별 상위 5', ['진료과·센터', '건수'], topN_(cur.byDept, 5));

  r = writeTable_(sh, r, '보도자료 채택 현황', ['보도자료ID', '제목', '배포매체수', '기사화건수', '채택률(%)'],
    pressStats.map(p => [p.id, p.title, p.distributed, p.adopted, p.rate]));

  r = writeTable_(sh, r, '주요 보도 하이라이트(1급 매체·포털메인, 최대5)', ['매체명', '제목', '링크'],
    cur.highlights.map(h => [h.media, h.title, h.link]));

  r = writeTable_(sh, r, '부정 보도 및 대응 현황', ['매체명', '제목', '논조', '위기등급', '대응상태', '링크'],
    cur.negatives.map(n => [n.media, n.title, n.tone, n.grade, n.status, n.link]));

  formatStatsSheet_(sh);
}

function buildMonthlyStats() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(now, tz, 'M'));
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd = new Date(y, m, 1);
  const prevMonthStart = new Date(y, m - 2, 1);
  const prevMonthEnd = monthStart;
  const lastYearStart = new Date(y - 1, m - 1, 1);
  const lastYearEnd = new Date(y - 1, m, 1);

  const cur = computeLedgerStats_(monthStart, monthEnd);
  const prevMonth = computeLedgerStats_(prevMonthStart, prevMonthEnd);
  const lastYear = computeLedgerStats_(lastYearStart, lastYearEnd);
  const pressStats = computePressStats_(monthStart, monthEnd);

  const sh = getOrCreateSheet_(SHEETS.MONTHLY);
  sh.clear();
  let r = 1;
  r = writeTitle_(sh, r, '월간 통계 — ' + y + '년 ' + m + '월');
  r = writeKv_(sh, r, [
    ['총 보도 건수', cur.total],
    ['전월 대비 증감률(%)', pctChange_(cur.total, prevMonth.total)],
    ['전년 동월 대비 증감률(%)', pctChange_(cur.total, lastYear.total)]
  ]);
  r++;

  r = writeTable_(sh, r, '매체별 보도량 순위(상위 20)', ['매체명', '건수'], topN_(cur.byMedia, 20));
  r = writeTable_(sh, r, '기자별 보도 건수(상위 10)', ['기자명', '건수'], topN_(cur.byReporter, 10));

  const totalDist = pressStats.reduce((s, p) => s + p.distributed, 0);
  const totalAdopt = pressStats.reduce((s, p) => s + p.adopted, 0);
  r = writeKv_(sh, r, [['월간 보도자료 채택률(%)', totalDist ? round1_(totalAdopt / totalDist * 100) : 0]]);
  r = writeTable_(sh, r, '주제별 보도자료 채택률', ['주제·진료과', '배포매체수', '기사화건수', '채택률(%)'], pressByTopic_(pressStats));

  const crisisSummary = computeCrisisSummary_(monthStart, monthEnd);
  r = writeTable_(sh, r, '위기 이슈 발생·처리 현황', ['등급', '건수', '평균 대응소요(시간)', '미종결 건수'], crisisSummary.rows);

  formatStatsSheet_(sh);
}

function buildYearlyStats() {
  const tz = Session.getScriptTimeZone();
  const y = Number(Utilities.formatDate(new Date(), tz, 'yyyy'));
  const yearStart = new Date(y, 0, 1);
  const yearEnd = new Date(y + 1, 0, 1);
  const prevYearStart = new Date(y - 1, 0, 1);
  const prevYearEnd = yearStart;

  const cur = computeLedgerStats_(yearStart, yearEnd);
  const prev = computeLedgerStats_(prevYearStart, prevYearEnd);

  const sh = getOrCreateSheet_(SHEETS.YEARLY);
  sh.clear();
  let r = 1;
  r = writeTitle_(sh, r, '연간 통계 — ' + y + '년');
  r = writeKv_(sh, r, [
    ['연간 총 보도 건수', cur.total],
    ['전년 대비 증감률(%)', pctChange_(cur.total, prev.total)]
  ]);
  r++;

  const monthly = [];
  for (let m = 1; m <= 12; m++) {
    const s = new Date(y, m - 1, 1), e = new Date(y, m, 1);
    monthly.push([m + '월', computeLedgerStats_(s, e).total]);
  }
  r = writeTable_(sh, r, '월별 보도량 추이', ['월', '건수'], monthly);

  r = writeTable_(sh, r, '카테고리 구성비 변화(금년 vs 전년, %)', ['카테고리', '금년', '전년'],
    OPTS.카테고리.map(c => [c, pctOf_(cur.byCategory[c], cur.total), pctOf_(prev.byCategory[c], prev.total)]));

  r = writeTable_(sh, r, '매체유형별 연간 구성비 변화(%)', ['매체유형', '금년', '전년'],
    OPTS.매체유형.map(t => [t, pctOf_(cur.byMediaType[t], cur.total), pctOf_(prev.byMediaType[t], prev.total)]));

  const kpi = computeKpiAchievement_(cur, y);
  r = writeTable_(sh, r, 'KPI 목표 대비 달성률(7.4항)', ['KPI', '금년 실적', '목표', '달성률(%)'], kpi);

  formatStatsSheet_(sh);
}

/** F-405 기간 임의 조회: 시작일·종료일을 입력받아 09_기간조회 시트에 결과를 기록한다. */
function runCustomRangeQuery() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('기간 조회 — 시작일', 'YYYY-MM-DD 형식으로 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('기간 조회 — 종료일(포함)', 'YYYY-MM-DD 형식으로 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  const start = new Date(r1.getResponseText().trim());
  const end = addDays_(new Date(r2.getResponseText().trim()), 1);
  if (isNaN(start) || isNaN(end)) { ui.alert('날짜 형식을 확인해 주세요.'); return; }

  const stats = computeLedgerStats_(start, end);
  const sh = getOrCreateSheet_('09_기간조회');
  sh.clear();
  let r = 1;
  r = writeTitle_(sh, r, '기간 조회 — ' + fmtDate_(start) + ' ~ ' + fmtDate_(addDays_(end, -1)));
  r = writeKv_(sh, r, [['총 보도 건수', stats.total], ['순 이슈 건수', stats.issueCount]]);
  r++;
  r = writeTable_(sh, r, '매체유형별 분포', ['매체유형', '건수'], topN_(stats.byMediaType, 20));
  r = writeTable_(sh, r, '논조별 분포', ['논조', '건수'], topN_(stats.byTone, 3));
  r = writeTable_(sh, r, '카테고리별 분포', ['카테고리', '건수'], topN_(stats.byCategory, 10));
  formatStatsSheet_(sh);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

// ── 핵심 집계 엔진 ────────────────────────────────────────────────────

/** [start, end) 구간에서 상태=확정 모니터링 현황 행을 집계한다. */
function computeLedgerStats_(start, end) {
  const rows = getConfirmedLedgerRows_(start, end);
  const stats = {
    total: rows.length,
    byMediaType: {}, byMediaInfluence: {}, byTone: {}, byCategory: {}, byDept: {}, byMedia: {}, byReporter: {},
    issueGroups: new Set(), highlights: [], negatives: []
  };
  rows.forEach(r => {
    bump_(stats.byMediaType, r[LEDGER_COL.매체유형 - 1]);
    bump_(stats.byMediaInfluence, r[LEDGER_COL.매체영향력 - 1]);
    bump_(stats.byTone, r[LEDGER_COL.논조확정 - 1]);
    bump_(stats.byCategory, r[LEDGER_COL.카테고리 - 1]);
    bump_(stats.byMedia, r[LEDGER_COL.매체명 - 1]);
    bump_(stats.byReporter, r[LEDGER_COL.기자명 - 1]);
    String(r[LEDGER_COL.진료과센터 - 1] || '').split(',').map(s => s.trim()).filter(String)
      .forEach(d => bump_(stats.byDept, d));
    if (r[LEDGER_COL.이슈그룹ID - 1]) stats.issueGroups.add(r[LEDGER_COL.이슈그룹ID - 1]);

    const isTop = r[LEDGER_COL.매체영향력 - 1] === '1급' || r[LEDGER_COL.노출도 - 1] === '포털메인';
    if (isTop && stats.highlights.length < 5) {
      stats.highlights.push({ media: r[LEDGER_COL.매체명 - 1], title: r[LEDGER_COL.제목 - 1], link: r[LEDGER_COL.URL - 1] });
    }
    const isNegative = r[LEDGER_COL.논조확정 - 1] === '부정' || (r[LEDGER_COL.위기등급 - 1] && r[LEDGER_COL.위기등급 - 1] !== '해당없음');
    if (isNegative) {
      stats.negatives.push({
        media: r[LEDGER_COL.매체명 - 1], title: r[LEDGER_COL.제목 - 1], tone: r[LEDGER_COL.논조확정 - 1],
        grade: r[LEDGER_COL.위기등급 - 1], status: r[LEDGER_COL.대응상태 - 1], link: r[LEDGER_COL.URL - 1]
      });
    }
  });
  stats.issueCount = stats.issueGroups.size;
  return stats;
}

/** 위기(등급≠해당없음) 요약 — 발생 건수/평균대응소요/미종결 (월간 리포트용). 확정 여부와 무관하게 전수 포함. */
function computeCrisisSummary_(start, end) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  const rows = [];
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues().forEach(r => {
      const pub = r[LEDGER_COL.보도일시 - 1];
      if (pub instanceof Date && pub >= start && pub < end && r[LEDGER_COL.위기등급 - 1] && r[LEDGER_COL.위기등급 - 1] !== '해당없음') {
        rows.push(r);
      }
    });
  }
  const byGrade = {};
  ['관심', '주의', '경계', '심각'].forEach(g => byGrade[g] = { count: 0, unresolved: 0, totalHours: 0, resolvedCount: 0 });
  rows.forEach(r => {
    const g = r[LEDGER_COL.위기등급 - 1];
    if (!byGrade[g]) return;
    byGrade[g].count++;
    if (r[LEDGER_COL.대응상태 - 1] !== '종결') byGrade[g].unresolved++;
    else {
      const collected = r[LEDGER_COL.수집일시 - 1];
      if (collected instanceof Date) {
        byGrade[g].totalHours += (Date.now() - collected.getTime()) / 3600000; // 근사치(정확한 종결시각 컬럼이 없어 현재시각 기준)
        byGrade[g].resolvedCount++;
      }
    }
  });
  const out = { rows: [] };
  ['심각', '경계', '주의', '관심'].forEach(g => {
    const d = byGrade[g];
    const avg = d.resolvedCount ? round1_(d.totalHours / d.resolvedCount) : '-';
    out.rows.push([g, d.count, avg, d.unresolved]);
  });
  return out;
}

/** 7.4 KPI 달성률. 목표치는 임계값 블록에 없으므로 전년 실적 기반 참고치를 함께 표기한다. */
function computeKpiAchievement_(cur, year) {
  const posRate = pctOf_(cur.byTone['긍정'], cur.total);
  const negRate = pctOf_(cur.byTone['부정'], cur.total);
  const top1 = cur.byMediaInfluence['1급'] || 0;
  return [
    ['보도 총량', cur.total, '전년 대비 증가(목표는 00_설정에서 별도 관리 권장)', '-'],
    ['긍정 보도 비율(%)', posRate, '기저치 설정 후 관리', '-'],
    ['부정 보도 비율(%)', negRate, '전년 대비 감소', '-'],
    ['1급 매체 노출 건수(근사)', top1, '전년 대비 증가', '-']
  ];
}

/** 사용자 요청: 통계 집계는 "보도통계 집계 매체" 목록(STATS_MEDIA_WHITELIST)에 속한 매체만 포함한다. */
function getConfirmedLedgerRows_(start, end) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues();
  return data.filter(r => {
    const pub = r[LEDGER_COL.보도일시 - 1];
    return r[LEDGER_COL.상태 - 1] === '확정' && pub instanceof Date && pub >= start && pub < end &&
      STATS_MEDIA_WHITELIST.indexOf(r[LEDGER_COL.매체명 - 1]) >= 0;
  });
}

function computePressStats_(start, end) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRESS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, PRESS_HEADERS.length).getValues();
  const out = [];
  data.forEach(r => {
    const dist = r[1];
    if (!(dist instanceof Date) || dist < start || dist >= end) return;
    const distributed = Number(r[4]) || 0;
    const adopted = Number(r[5]) || 0;
    out.push({
      id: r[0], title: r[2], topic: r[3], distributed: distributed, adopted: adopted,
      rate: distributed ? round1_(adopted / distributed * 100) : 0
    });
  });
  return out;
}

function pressByTopic_(pressStats) {
  const map = {};
  pressStats.forEach(p => {
    const k = p.topic || '(미분류)';
    if (!map[k]) map[k] = { distributed: 0, adopted: 0 };
    map[k].distributed += p.distributed;
    map[k].adopted += p.adopted;
  });
  return Object.keys(map).map(k => [k, map[k].distributed, map[k].adopted,
    map[k].distributed ? round1_(map[k].adopted / map[k].distributed * 100) : 0]);
}

// ── 시트 출력 헬퍼 ────────────────────────────────────────────────────

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeTitle_(sh, row, text) {
  sh.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(14);
  return row + 2;
}

function writeKv_(sh, row, pairs) {
  pairs.forEach((p, i) => sh.getRange(row + i, 1, 1, 2).setValues([p]));
  return row + pairs.length + 1;
}

function writeTable_(sh, row, title, headers, rows) {
  sh.getRange(row, 1).setValue(title).setFontWeight('bold');
  row++;
  sh.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F1F3F4');
  row++;
  if (rows.length) {
    sh.getRange(row, 1, rows.length, headers.length).setValues(rows);
    row += rows.length;
  }
  return row + 1;
}

function formatStatsSheet_(sh) {
  sh.autoResizeColumns(1, 8);
  sh.setFrozenRows(2);
}

function bump_(map, key) {
  if (!key) return;
  map[key] = (map[key] || 0) + 1;
}
function distTable_(map, total, topFlagN) {
  const entries = Object.keys(map).map(k => [k, map[k], pctOf_(map[k], total)]);
  entries.sort((a, b) => b[1] - a[1]);
  if (topFlagN) entries.forEach((e, i) => { if (i < topFlagN) e[0] = '★ ' + e[0]; });
  return entries;
}
function topN_(map, n) {
  return Object.keys(map).map(k => [k, map[k]]).sort((a, b) => b[1] - a[1]).slice(0, n);
}
function pctOf_(count, total) { return total ? round1_((count || 0) / total * 100) : 0; }
function pctChange_(cur, prev) { return prev ? round1_((cur - prev) / prev * 100) : (cur ? 100 : 0); }
function round1_(n) { return Math.round(n * 10) / 10; }
function addDays_(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function fmtDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
