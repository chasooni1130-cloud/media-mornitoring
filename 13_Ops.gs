/**
 * 13_Ops.gs — 운영 관리 (F-107 API 한도, F-304 아카이브, F-305 백업,
 * F-603 장애 알림, F-604 메시지 잔액 관리)
 * F-306(엑셀 내려받기)·F-602(권한 관리)는 구글 시트 기본 기능/공유 설정으로 처리하며 별도 코드가 필요 없다.
 */

// ── F-107 API 한도 관리 ──────────────────────────────────────────────

function checkApiQuota_() {
  const th = getThresholds_();
  const limit = Number(th['네이버_API_월간한도']) || 775000;
  const warnRatio = Number(th['API_한도_경고_비율']) || 0.8;
  const used = getMonthlyCounter_('NAVER_API_CALLS'); // NAVER API HUB는 월 단위 통합 한도
  const props = PropertiesService.getScriptProperties();
  const monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');

  if (used >= limit) {
    props.setProperty('FORCE_RELAXED_COLLECT', 'Y');
    if (props.getProperty('QUOTA_100_NOTIFIED_' + monthKey) !== 'Y') {
      notifyOps_('네이버 API 월간 한도 도달', '이번 달 호출 건수 ' + used + '/' + limit + '. 수집 주기를 자동으로 완화합니다.');
      props.setProperty('QUOTA_100_NOTIFIED_' + monthKey, 'Y');
    }
  } else if (used >= limit * warnRatio) {
    if (props.getProperty('QUOTA_80_NOTIFIED_' + monthKey) !== 'Y') {
      notifyOps_('네이버 API 월간 한도 80% 도달', '이번 달 호출 건수 ' + used + '/' + limit);
      props.setProperty('QUOTA_80_NOTIFIED_' + monthKey, 'Y');
    }
  }
}

function todayKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

/** getCollectIntervalMinutes_()가 참조하는 강제 완화 플래그. 월 단위 한도이므로 다음 달이 되면 자동 리셋된다. */
function isForceRelaxed_() {
  const monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM');
  return PropertiesService.getScriptProperties().getProperty('FORCE_RELAXED_COLLECT') === 'Y'
    && PropertiesService.getScriptProperties().getProperty('QUOTA_100_NOTIFIED_' + monthKey) === 'Y';
}

// ── F-603 장애 알림 ──────────────────────────────────────────────────

/** 08_로그의 최근 기록을 보고 연속 실패 또는 장시간 무수집 상태를 감지한다. */
function checkStaleCollection_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const th = getThresholds_();
  const failThreshold = Number(th['수집실패_경고_연속횟수']) || 3;
  const staleHours = Number(th['수집중단_경고_시간']) || 6;

  const scan = Math.min(50, lastRow - 1);
  const startRow = lastRow - scan + 1;
  const rows = sh.getRange(startRow, 1, scan, LOG_HEADERS.length).getValues()
    .filter(r => r[1] === 'collectAll');

  let consecutiveFail = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][4]) consecutiveFail++; else break; // 오류내용 컬럼이 채워져 있으면 실패로 간주
  }
  if (consecutiveFail >= failThreshold) {
    notifyOps_('수집 연속 실패 ' + consecutiveFail + '회', '최근 로그를 08_로그 시트에서 확인해 주세요.');
  }

  const lastSuccessRow = rows.filter(r => (r[3] || 0) > 0).pop(); // 신규저장건수 > 0인 마지막 행
  if (lastSuccessRow) {
    const hoursSince = (Date.now() - new Date(lastSuccessRow[0]).getTime()) / 3600000;
    if (hoursSince >= staleHours) {
      notifyOps_('신규 수집 없음 (' + round1_(hoursSince) + '시간)', '수집 로직 또는 API 키 상태를 점검해 주세요.');
    }
  }
}

// ── F-604 메시지 잔액 관리 ────────────────────────────────────────────

/** SOLAPI 잔액 조회(참고 구현). 실제 응답 형식은 SOLAPI 문서를 기준으로 조정이 필요할 수 있다. */
function checkMessageBalance_() {
  const apiKey = getSecret_('SOLAPI_API_KEY');
  const apiSecret = getSecret_('SOLAPI_API_SECRET');
  if (!apiKey || !apiSecret) return;
  try {
    const resp = UrlFetchApp.fetch('https://api.solapi.com/cash/v1/balance', {
      method: 'get', headers: { Authorization: solapiAuthHeader_(apiKey, apiSecret) }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return;
    const json = JSON.parse(resp.getContentText());
    const balance = Number(json.balance || json.point || 0);
    const threshold = Number(getSecret_('MESSAGE_BALANCE_ALERT_THRESHOLD') || 5000);
    if (balance < threshold) {
      const props = PropertiesService.getScriptProperties();
      if (props.getProperty('BALANCE_NOTIFIED_' + todayKey_()) !== 'Y') {
        notifyOps_('메시지 발송 잔액 부족', '현재 잔액 약 ' + balance + '원. 충전이 필요합니다.');
        props.setProperty('BALANCE_NOTIFIED_' + todayKey_(), 'Y');
      }
    }
  } catch (e) {
    writeLog_('checkMessageBalance_', 0, 0, e.message, 0);
  }
}

// ── F-305 백업 ───────────────────────────────────────────────────────

/** 주 1회 트리거로 01_모니터링현황을 백업 스프레드시트로 복사한다(최근 12주 유지). */
function backupLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSh = ss.getSheetByName(SHEETS.LEDGER);
  const folder = getOrCreateBackupFolder_();

  const tz = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const name = '모니터링현황백업_' + stamp;

  const file = DriveApp.getFileById(ss.getId());
  const copy = file.makeCopy(name, folder);
  const copySs = SpreadsheetApp.openById(copy.getId());
  copySs.getSheets().forEach(sh => { if (sh.getName() !== SHEETS.LEDGER) copySs.deleteSheet(sh); });

  pruneOldBackups_(folder, 12);
  writeLog_('backupLedger', 0, 0, '', 0);
}

/**
 * 1회성 정리용(사용자 요청). 보도일시가 "어제" 00:00보다 이전인 행을 전부 삭제해
 * 화면에 어제·오늘 데이터만 남긴다. 실행 전 [모니터링 현황 백업 실행]으로 먼저 백업해 둘 것 —
 * 이 함수는 삭제만 하고 백업은 만들지 않는다.
 * 01_모니터링현황이 보도일시 내림차순 정렬 상태라는 전제로, 위에서부터 훑다가
 * 처음으로 어제 이전 날짜가 나오는 지점부터 끝까지를 지운다.
 */
function trimLedgerBeforeYesterday() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { alertSafe_('모니터링 현황에 데이터가 없습니다.'); return; }

  const tz = Session.getScriptTimeZone();
  const todayStart = new Date(Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd'));
  const yesterdayStart = new Date(todayStart.getTime() - 24 * 3600 * 1000);

  // 정렬이 항상 완벽하게 유지된다고 가정하지 않는다(중복 폭증 등으로 날짜가 뒤섞여 있을 수 있음).
  // 전체를 읽어 어제 이후만 남기고 다시 쓴다 — 행이 매우 많을 경우를 대비해 청크로 읽는다.
  const numRows = lastRow - 1;
  const chunk = 20000;
  const kept = [];
  for (let offset = 0; offset < numRows; offset += chunk) {
    const size = Math.min(chunk, numRows - offset);
    const vals = sh.getRange(2 + offset, 1, size, LEDGER_LAST_COL).getValues();
    vals.forEach(row => {
      const pub = row[LEDGER_COL.보도일시 - 1];
      if (!(pub instanceof Date) || pub >= yesterdayStart) kept.push(row);
    });
  }

  const deleteCount = numRows - kept.length;
  if (deleteCount <= 0) { alertSafe_('어제 이전 데이터가 없습니다. 삭제할 행이 없습니다.'); return; }

  sh.getRange(2, 1, numRows, LEDGER_LAST_COL).clearContent();
  if (kept.length) sh.getRange(2, 1, kept.length, LEDGER_LAST_COL).setValues(kept);
  // 남는 빈 행을 실제로 지워 시트 크기(getLastRow 등)도 함께 줄인다.
  const trailingBlank = numRows - kept.length;
  if (trailingBlank > 0) sh.deleteRows(2 + kept.length, trailingBlank);
  sortLedgerByReportDate_(sh);
  alertSafe_('정리 완료: ' + deleteCount + '건(어제 이전 데이터)을 삭제했습니다. 남은 건수: ' + kept.length);
}

/**
 * 1회성 정리용(사용자 요청, 2026-09). EXCLUDED_MEDIA(00_Config.gs)에 등록된 매체를
 * (1) 00_설정!매체마스터에서 완전히 제거하고 (2) 01_모니터링현황에서 해당 매체의 기존 기사를 전부 삭제한다.
 * 00_설정 시트는 여러 블록이 같은 행을 열(column)만 나눠 공유하므로, 매체마스터 정리는
 * deleteRows가 아니라 H:K 열 범위만 다시 채우는 방식으로 해야 다른 블록(위기키워드 등)이 밀리지 않는다.
 */
function purgeExcludedMedia() {
  const cfgSh = getConfigSheet_();
  const block = CFG_BLOCK.매체마스터;
  const rows = readConfigBlock_('매체마스터');
  const kept = rows.filter(r => EXCLUDED_MEDIA.indexOf(r['표준매체명']) < 0);
  const removedFromMaster = rows.length - kept.length;

  const startRange = cfgSh.getRange(block.dataStart);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();
  const numCols = block.cols.length;
  if (rows.length) cfgSh.getRange(startRow, startCol, rows.length, numCols).clearContent();
  if (kept.length) {
    cfgSh.getRange(startRow, startCol, kept.length, numCols)
      .setValues(kept.map(r => block.cols.map(c => r[c])));
  }

  const ledgerSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = ledgerSh.getLastRow();
  let removedFromLedger = 0;
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    const chunk = 20000;
    const keptLedger = [];
    for (let offset = 0; offset < numRows; offset += chunk) {
      const size = Math.min(chunk, numRows - offset);
      const vals = ledgerSh.getRange(2 + offset, 1, size, LEDGER_LAST_COL).getValues();
      vals.forEach(row => {
        if (EXCLUDED_MEDIA.indexOf(row[LEDGER_COL.매체명 - 1]) >= 0) removedFromLedger++;
        else keptLedger.push(row);
      });
    }
    if (removedFromLedger > 0) {
      ledgerSh.getRange(2, 1, numRows, LEDGER_LAST_COL).clearContent();
      if (keptLedger.length) ledgerSh.getRange(2, 1, keptLedger.length, LEDGER_LAST_COL).setValues(keptLedger);
      const trailingBlank = numRows - keptLedger.length;
      if (trailingBlank > 0) ledgerSh.deleteRows(2 + keptLedger.length, trailingBlank);
      sortLedgerByReportDate_(ledgerSh);
    }
  }

  writeLog_('purgeExcludedMedia', rows.length, removedFromMaster + removedFromLedger, '', 0);
  alertSafe_('제외 매체 정리 완료 — 매체마스터에서 ' + removedFromMaster + '건, 모니터링 현황에서 ' + removedFromLedger + '건 삭제했습니다.');
}

/**
 * 1회성 정리용(사용자 요청, 2026-10). 01_모니터링현황!매체명에 영문 도메인(URL) 형태 그대로 남아있는
 * 행을 정리한다. DOMAIN_MEDIA_FALLBACK(00_Config.gs)에 등록된 도메인은 한글 매체명(+매체유형·매체영향력)
 * 으로 바꿔 쓰고, 등록되지 않은 영문 도메인 매체는 행 자체를 삭제한다. 앞으로의 수집은
 * normalizeMedia_의 "도메인제외" 분기(05_Pipeline.gs)가 같은 기준으로 자동 차단한다.
 */
function renameAndPurgeDomainMedia() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { alertSafe_('모니터링 현황에 데이터가 없습니다.'); return; }

  const numRows = lastRow - 1;
  const chunk = 20000;
  const kept = [];
  let renamed = 0, deleted = 0;

  for (let offset = 0; offset < numRows; offset += chunk) {
    const size = Math.min(chunk, numRows - offset);
    const vals = sh.getRange(2 + offset, 1, size, LEDGER_LAST_COL).getValues();
    vals.forEach(row => {
      const mediaRaw = String(row[LEDGER_COL.매체명 - 1] || '').trim();
      const key = mediaRaw.toLowerCase();
      if (!isDomainLike_(key)) { kept.push(row); return; }
      const mapped = DOMAIN_MEDIA_FALLBACK[key];
      if (mapped) {
        row[LEDGER_COL.매체명 - 1] = mapped.name;
        if (!row[LEDGER_COL.매체유형 - 1]) row[LEDGER_COL.매체유형 - 1] = mapped.type;
        if (!row[LEDGER_COL.매체영향력 - 1]) row[LEDGER_COL.매체영향력 - 1] = mapped.influence;
        renamed++;
        kept.push(row);
      } else {
        deleted++;
      }
    });
  }

  sh.getRange(2, 1, numRows, LEDGER_LAST_COL).clearContent();
  if (kept.length) sh.getRange(2, 1, kept.length, LEDGER_LAST_COL).setValues(kept);
  const trailingBlank = numRows - kept.length;
  if (trailingBlank > 0) sh.deleteRows(2 + kept.length, trailingBlank);
  sortLedgerByReportDate_(sh);

  writeLog_('renameAndPurgeDomainMedia', numRows, renamed + deleted, '', 0);
  alertSafe_('도메인 매체 정리 완료 — 한글명으로 변경 ' + renamed + '건, 미등록 도메인 삭제 ' + deleted + '건.');
}

/**
 * 1회성 등록용(사용자 요청, 2026-10). renameAndPurgeDomainMedia()에서 한글명으로 치환한 16개 영문 도메인
 * 매체를 00_설정!매체마스터에도 정식 등록한다. 매체마스터는 DOMAIN_MEDIA_FALLBACK(00_Config.gs)보다 항상
 * 우선 적용되므로, 등록해두면 이후 수집·정규화가 매체마스터 값을 바로 쓰게 된다(값은 DOMAIN_MEDIA_FALLBACK과
 * 동일하게 가져온다 — 두 군데가 어긋나지 않도록).
 */
function registerDomainMediaToMaster() {
  const domains = [
    'woman.donga.com', 'sedaily.com', 'digitalchosun.dizzo.com', 'sports.donga.com',
    'sportsseoul.com', 'news.tvchosun.com', 'sportsworldi.com', 'news.sbs.co.kr',
    'news.kbs.co.kr', 'biz.chosun.com', 'mbn.co.kr', 'ytn.co.kr', 'ichannela.com',
    'segye.com', 'dt.co.kr', 'yonhapnewstv.co.kr'
  ];
  const cfgSh = getConfigSheet_();
  const block = CFG_BLOCK.매체마스터;
  const existing = readConfigBlock_('매체마스터');
  const existingKeys = new Set(existing.map(r => String(r['수집매체명']).trim().toLowerCase()));

  const toAdd = [];
  domains.forEach(d => {
    if (existingKeys.has(d)) return;
    const fb = DOMAIN_MEDIA_FALLBACK[d];
    if (!fb) return;
    toAdd.push([d, fb.name, fb.type, fb.influence]);
  });

  if (!toAdd.length) { alertSafe_('추가할 매체가 없습니다(이미 모두 등록됨).'); return; }

  const startRange = cfgSh.getRange(block.dataStart);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();
  const targetRow = startRow + existing.length;
  cfgSh.getRange(targetRow, startCol, toAdd.length, block.cols.length).setValues(toAdd);

  writeLog_('registerDomainMediaToMaster', domains.length, toAdd.length, '', 0);
  alertSafe_('매체마스터에 ' + toAdd.length + '건을 추가했습니다(총 ' + (existing.length + toAdd.length) + '건).');
}

/**
 * 1회성 정리용(사용자 요청, 2026-09-30). 보도일시가 2026-09-16 00:00보다 이전인 행을 전부 삭제해
 * 화면에 9월 16일 이후 데이터만 남긴다. trimLedgerBeforeYesterday와 동일한 안전한 방식(위치가 아니라
 * 날짜로 매 행을 개별 필터링 후 재기록)을 쓴다 — 실행 전 [모니터링 현황 백업 실행]을 먼저 하는 것을 권장한다.
 */
function trimLedgerBefore0916() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { alertSafe_('모니터링 현황에 데이터가 없습니다.'); return; }

  const tz = Session.getScriptTimeZone();
  const cutoff = new Date(Utilities.formatDate(new Date(2026, 8, 16), tz, 'yyyy/MM/dd')); // 2026-09-16 00:00

  const numRows = lastRow - 1;
  const chunk = 20000;
  const kept = [];
  for (let offset = 0; offset < numRows; offset += chunk) {
    const size = Math.min(chunk, numRows - offset);
    const vals = sh.getRange(2 + offset, 1, size, LEDGER_LAST_COL).getValues();
    vals.forEach(row => {
      const pub = row[LEDGER_COL.보도일시 - 1];
      if (!(pub instanceof Date) || pub >= cutoff) kept.push(row);
    });
  }

  const deleteCount = numRows - kept.length;
  if (deleteCount <= 0) { alertSafe_('9월 16일 이전 데이터가 없습니다. 삭제할 행이 없습니다.'); return; }

  sh.getRange(2, 1, numRows, LEDGER_LAST_COL).clearContent();
  if (kept.length) sh.getRange(2, 1, kept.length, LEDGER_LAST_COL).setValues(kept);
  const trailingBlank = numRows - kept.length;
  if (trailingBlank > 0) sh.deleteRows(2 + kept.length, trailingBlank);
  sortLedgerByReportDate_(sh);
  alertSafe_('정리 완료: ' + deleteCount + '건(9월 16일 이전 데이터)을 삭제했습니다. 남은 건수: ' + kept.length);
}

/**
 * 1회성 시딩용(사용자 요청, 2026-09-30). 고려대 구로병원 전임·임상·진료교원 명단(2026-07 기준 엑셀)을
 * 00_설정!의료진사전에 채운다. 이미 등록된 (이름,소속) 조합은 건너뛰고 신규만 추가한다.
 */
function seedMedicalStaffDictionary() {
  const cfgSh = getConfigSheet_();
  const block = CFG_BLOCK.의료진사전;
  const existing = readConfigBlock_('의료진사전');
  const existingKeys = new Set(existing.map(r => r['이름'] + '|' + r['소속']));
  const toAdd = MEDICAL_STAFF_SEED_.filter(r => !existingKeys.has(r[0] + '|' + r[1]));
  if (!toAdd.length) { alertSafe_('추가할 신규 의료진이 없습니다(이미 모두 등록됨).'); return; }

  const startRange = cfgSh.getRange(block.dataStart);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();
  const targetRow = startRow + existing.length;
  cfgSh.getRange(targetRow, startCol, toAdd.length, 2).setValues(toAdd);
  alertSafe_('의료진 사전에 ' + toAdd.length + '명을 추가했습니다(총 ' + (existing.length + toAdd.length) + '명).');
}

/**
 * 1회성 마이그레이션(사용자 요청, 2026-09-30). 09_수집이력 시트를 라이브 스프레드시트에 만든다.
 * initializeSystem()에도 추가해뒀지만, 이미 초기 설정이 끝난 기존 시트에는 자동 반영되지 않으므로
 * 이 함수를 한 번 실행해 지금 바로 생성한다(이미 있으면 아무 일도 하지 않음).
 */
function migrateAddHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSimpleSheet_(ss, SHEETS.HISTORY, HISTORY_HEADERS);
  alertSafe_('09_수집이력 시트를 준비했습니다.');
}

/**
 * 1회성 백필(사용자 요청, 2026-09-30). 01_모니터링현황을 9월 16일 이전 삭제한 뒤, 네이버 검색에
 * 다시 잡힌 옛 기사가 "새 기사"로 오인되어 재적재·재알림되는 사고가 있었다(F-201이 화면에 보이는
 * 01_모니터링현황만 기준으로 동작했기 때문). 09_수집이력을 영구 dedup 저장소로 쓰기로 하고,
 * (1) 현재 01_모니터링현황(재유입된 옛 행 포함)과 (2) 트림 직전 백업 스프레드시트의 dedupKey를
 * 모두 모아 09_수집이력에 채운다 — 이렇게 해야 다시 트림해도 같은 사고가 재발하지 않는다.
 */
function backfillDedupHistoryFromCurrentAndBackup() {
  const historySh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.HISTORY);
  if (!historySh) { alertSafe_('09_수집이력 시트가 없습니다. migrateAddHistorySheet()를 먼저 실행하세요.'); return; }

  const existing = new Set(getPermanentDedupKeys_());
  const toAdd = [];
  let fromCurrent = 0, fromBackup = 0;

  const collectFrom = (sh, counterName) => {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const data = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues();
    data.forEach(row => {
      const key = dedupKey_(row[LEDGER_COL.URL - 1], row[LEDGER_COL.매체명 - 1], row[LEDGER_COL.제목 - 1]);
      if (!key || existing.has(key)) return;
      existing.add(key);
      toAdd.push([row[LEDGER_COL.기사ID - 1], key, row[LEDGER_COL.보도일시 - 1], row[LEDGER_COL.수집일시 - 1]]);
      if (counterName === 'current') fromCurrent++; else fromBackup++;
    });
  };

  const currentLedger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  collectFrom(currentLedger, 'current');

  const folder = getOrCreateBackupFolder_();
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  if (files.length) {
    const backupSs = SpreadsheetApp.openById(files[0].getId());
    const backupLedgerSh = backupSs.getSheetByName(SHEETS.LEDGER);
    if (backupLedgerSh) collectFrom(backupLedgerSh, 'backup');
  }

  if (toAdd.length) historySh.getRange(historySh.getLastRow() + 1, 1, toAdd.length, HISTORY_HEADERS.length).setValues(toAdd);
  alertSafe_('09_수집이력 백필 완료 — 현재 시트 ' + fromCurrent + '건, 백업 스프레드시트 ' + fromBackup + '건, 총 ' + toAdd.length + '건 추가했습니다.');
}

/**
 * 1회성 마이그레이션(사용자 요청, 2026-09-30). 그동안 "의료진" 한 컬럼에 "소속 이름" 형태로
 * 합쳐서 기입하던 것을 "진료과"·"의료진" 두 컬럼으로 나눈다. 기존 27번 열(의료진, 구 형식 데이터)
 * 앞에 새 열을 끼워넣어 28번으로 밀어내고, 27번을 새 "진료과" 헤더로 만든다.
 * 실행 후 [의료진 태그 재계산(유지보수)]을 한 번 더 돌려서 실제 값을 다시 채워야 한다.
 */
function migrateLedgerSplitStaffDept() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  if (!sh) { alertSafe_('01_모니터링현황 시트를 찾을 수 없습니다.'); return; }

  if (sh.getRange(1, 27).getValue() === '의료진') {
    sh.insertColumnBefore(27); // 기존 의료진(27) 데이터를 28로 밀어내고 27을 비운다
  }
  if (sh.getMaxColumns() < LEDGER_LAST_COL) {
    sh.insertColumnsAfter(sh.getMaxColumns(), LEDGER_LAST_COL - sh.getMaxColumns());
  }
  sh.getRange(1, LEDGER_COL.진료과).setValue('진료과')
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF')
    .setNote('의료진 컬럼에서 찾은 사람의 00_설정!의료진 사전 상 소속(복수 시 쉼표로 나열).');
  sh.getRange(1, LEDGER_COL.의료진).setValue('의료진')
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF')
    .setNote('00_설정!의료진 사전에 등록된 이름이 제목·요약에서 발견되면 자동으로 채워집니다. 등록 전에는 비어 있습니다.');
  alertSafe_('진료과 컬럼을 추가했습니다. 이어서 [의료진 태그 재계산(유지보수)]을 실행해 기존 데이터를 다시 채워주세요.');
}

/**
 * 1회성 마이그레이션(사용자 요청, 2026-10). 01_모니터링현황의 화면 컬럼 순서를
 * 보도날짜·매체명·의료진·진료과·제목(URL 하이퍼링크) 순으로 재배치하고,
 * URL·매체유형·검색키워드·카테고리·보도유형은 숨긴다(데이터는 내부 로직용으로 유지).
 * 시트를 통째로 지우고 새 컬럼 순서(00_Config.gs의 새 LEDGER_COL)로 다시 만든 뒤 데이터를 옮겨 붓는다
 * — 이렇게 해야 검증(드롭다운)·조건부 서식·숨김 설정이 새 위치에 깨끗하게 다시 적용된다.
 */
function migrateLedgerReorderColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSh = ss.getSheetByName(SHEETS.LEDGER);
  if (!oldSh) { alertSafe_('01_모니터링현황 시트를 찾을 수 없습니다.'); return; }

  // 마이그레이션 직전(오늘 진료과/의료진 분리 이후)의 실제 컬럼 배치 스냅샷 — 00_Config.gs는 이미 새 순서로 바뀌어 있으므로 여기 별도 보관.
  const OLD_COL = {
    기사ID: 1, 수집일시: 2, 보도일시: 3, 매체명: 4, 매체유형: 5, 매체영향력: 6, 기자명: 7, 제목: 8, URL: 9, 요약: 10,
    검색키워드: 11, 카테고리: 12, 진료과센터: 13, 논조자동: 14, 논조확정: 15, 위기등급: 16, 이슈그룹ID: 17, 보도유형: 18,
    보도자료ID: 19, 노출도: 20, 수집경로: 21, 상태: 22, 대응상태: 23, 조치내역: 24, 담당자: 25, 비고: 26, 진료과: 27, 의료진: 28
  };
  const OLD_LAST_COL = 28;

  const lastRow = oldSh.getLastRow();
  const numRows = Math.max(0, lastRow - 1);
  const oldData = numRows > 0 ? oldSh.getRange(2, 1, numRows, OLD_LAST_COL).getValues() : [];

  const sheetIndex = oldSh.getIndex();
  ss.deleteSheet(oldSh);
  setupLedgerSheet_(ss); // 새 LEDGER_COL 순서로 헤더·검증·서식·숨김을 처음부터 다시 만든다
  const newSh = ss.getSheetByName(SHEETS.LEDGER);
  ss.setActiveSheet(newSh);
  ss.moveActiveSheet(sheetIndex);

  if (numRows > 0) {
    const newData = oldData.map(oldRow => {
      const row = new Array(LEDGER_LAST_COL).fill('');
      Object.keys(OLD_COL).forEach(key => { row[LEDGER_COL[key] - 1] = oldRow[OLD_COL[key] - 1]; });
      row[LEDGER_COL.제목 - 1] = titleHyperlinkFormula_(row[LEDGER_COL.제목 - 1], row[LEDGER_COL.URL - 1]);
      return row;
    });
    newSh.getRange(2, 1, newData.length, LEDGER_LAST_COL).setValues(newData);
    sortLedgerByReportDate_(newSh);
  }

  alertSafe_('모니터링 현황 컬럼 재구성 완료 — ' + numRows + '건을 새 순서로 옮겼습니다.');
}

/**
 * 긴급 복구용(2026-10-01). migrateLedgerReorderColumns() 실행 중 "Document ... is missing" 예외로
 * 01_모니터링현황이 시트 삭제 직후 빈 채로 남았다. 시트를 다시 지우지 않고(위험 최소화) 그대로 재사용해
 * 헤더·검증·서식을 처음부터 다시 만들고, 직전에 만든 백업 스프레드시트에서 데이터를 가져와
 * 새 컬럼 순서로 재배치하며 복원한다.
 */
function restoreLedgerFromBackupAndReorder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const OLD_COL = {
    기사ID: 1, 수집일시: 2, 보도일시: 3, 매체명: 4, 매체유형: 5, 매체영향력: 6, 기자명: 7, 제목: 8, URL: 9, 요약: 10,
    검색키워드: 11, 카테고리: 12, 진료과센터: 13, 논조자동: 14, 논조확정: 15, 위기등급: 16, 이슈그룹ID: 17, 보도유형: 18,
    보도자료ID: 19, 노출도: 20, 수집경로: 21, 상태: 22, 대응상태: 23, 조치내역: 24, 담당자: 25, 비고: 26, 진료과: 27, 의료진: 28
  };
  const OLD_LAST_COL = 28;

  const folder = getOrCreateBackupFolder_();
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  if (!files.length) { alertSafe_('백업 파일을 찾을 수 없습니다.'); return; }
  const backupSs = SpreadsheetApp.openById(files[0].getId());
  const backupLedgerSh = backupSs.getSheetByName(SHEETS.LEDGER);
  if (!backupLedgerSh) { alertSafe_('백업 스프레드시트에 01_모니터링현황 시트가 없습니다.'); return; }

  const lastRow = backupLedgerSh.getLastRow();
  const numRows = Math.max(0, lastRow - 1);
  const oldData = numRows > 0 ? backupLedgerSh.getRange(2, 1, numRows, OLD_LAST_COL).getValues() : [];

  let sh = ss.getSheetByName(SHEETS.LEDGER);
  if (!sh) sh = ss.insertSheet(SHEETS.LEDGER);
  sh.clear();
  try { sh.getDataRange().clearDataValidations(); } catch (e) { /* 빈 시트면 무시 */ }
  sh.clearConditionalFormatRules();

  sh.getRange(1, 1, 1, LEDGER_LAST_COL).setValues([LEDGER_HEADERS])
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF');
  sh.getRange(1, LEDGER_COL.보도일시).setValue('보도날짜');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  hideLedgerColumns_(sh);

  const maxRows = 5000;
  if (sh.getMaxRows() < maxRows) sh.insertRowsAfter(sh.getMaxRows(), maxRows - sh.getMaxRows());
  sh.getRange(2, LEDGER_COL.보도일시, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd');

  applyValidation_(sh, LEDGER_COL.매체유형, maxRows, OPTS.매체유형);
  applyValidation_(sh, LEDGER_COL.매체영향력, maxRows, OPTS.매체영향력);
  applyValidation_(sh, LEDGER_COL.카테고리, maxRows, OPTS.카테고리);
  applyValidation_(sh, LEDGER_COL.논조자동, maxRows, OPTS.논조);
  applyValidation_(sh, LEDGER_COL.논조확정, maxRows, OPTS.논조);
  applyValidation_(sh, LEDGER_COL.위기등급, maxRows, OPTS.위기등급);
  applyValidation_(sh, LEDGER_COL.보도유형, maxRows, OPTS.보도유형);
  applyValidation_(sh, LEDGER_COL.노출도, maxRows, OPTS.노출도);
  applyValidation_(sh, LEDGER_COL.수집경로, maxRows, OPTS.수집경로);
  applyValidation_(sh, LEDGER_COL.상태, maxRows, OPTS.상태);
  applyValidation_(sh, LEDGER_COL.대응상태, maxRows, OPTS.대응상태);

  const statusRange = sh.getRange(2, LEDGER_COL.상태, maxRows - 1, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('미확인').setBackground('#FFF3CD').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('확정').setBackground('#D4EDDA').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('제외').setBackground('#E2E3E5').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('오탐').setBackground('#E2E3E5').setRanges([statusRange]).build()
  ];
  const crisisRange = sh.getRange(2, LEDGER_COL.위기등급, maxRows - 1, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('심각').setBackground('#F8D7DA').setFontColor('#721C24').setBold(true).setRanges([crisisRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('경계').setBackground('#FDE2C8').setRanges([crisisRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('주의').setBackground('#FFF3CD').setRanges([crisisRange]).build()
  );
  sh.setConditionalFormatRules(rules);

  sh.getRange(1, LEDGER_COL.기사ID).setNote('예: 20260820-0031 (수집일시+일련번호, 시스템 자동 생성)');
  sh.getRange(1, LEDGER_COL.논조확정).setNote('통계 산출에는 반드시 이 컬럼(확정값)만 사용합니다. 논조(자동)은 정확도 측정용으로 보존합니다.');
  sh.getRange(1, LEDGER_COL.진료과).setNote('의료진 컬럼에서 찾은 사람의 00_설정!의료진 사전 상 소속(복수 시 쉼표로 나열).');
  sh.getRange(1, LEDGER_COL.의료진).setNote('00_설정!의료진 사전에 등록된 이름이 제목·요약에서 발견되면 자동으로 채워집니다. 등록 전에는 비어 있습니다.');

  if (numRows > 0) {
    const newData = oldData.map(oldRow => {
      const row = new Array(LEDGER_LAST_COL).fill('');
      Object.keys(OLD_COL).forEach(key => { row[LEDGER_COL[key] - 1] = oldRow[OLD_COL[key] - 1]; });
      row[LEDGER_COL.제목 - 1] = titleHyperlinkFormula_(row[LEDGER_COL.제목 - 1], row[LEDGER_COL.URL - 1]);
      return row;
    });
    sh.getRange(2, 1, newData.length, LEDGER_LAST_COL).setValues(newData);
    sortLedgerByReportDate_(sh);
  }

  alertSafe_('복구 및 컬럼 재구성 완료 — 백업(' + files[0].getName() + ')에서 ' + numRows + '건을 복원했습니다.');
}

/**
 * 1회성(사용자 요청, 2026-10-01). 기자명·진료과센터·논조자동 컬럼을 화면에서 숨기고,
 * 09_수집이력 시트는 "삭제" 대신 숨김 처리한다 — 이 시트를 지우면 9월 16일 이전 삭제 후
 * 겪었던 "옛 기사가 되살아나 재알림되는" 사고(F-201)를 막는 영구 dedup 이력이 사라져
 * 같은 문제가 재발할 수 있기 때문이다. 데이터는 그대로 두고 탭만 안 보이게 한다.
 */
function hideExtraColumnsAndHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSh = ss.getSheetByName(SHEETS.LEDGER);
  if (ledgerSh) {
    ledgerSh.showColumns(1, LEDGER_LAST_COL); // 재적용 전 전체 보이기로 리셋해 숨김 범위 꼬임 방지
    hideLedgerColumns_(ledgerSh);
  }
  const historySh = ss.getSheetByName(SHEETS.HISTORY);
  if (historySh) historySh.hideSheet();
  alertSafe_('기자명·진료과센터·논조자동 컬럼을 숨겼고, 09_수집이력 탭도 숨겼습니다(삭제는 하지 않았습니다 — 중복 재수집 방지용 데이터라 보존이 필요합니다).');
}

/**
 * 1회성(사용자 요청, 2026-10-01). 09_수집이력 시트를 완전히 삭제한다.
 * 주의: 이후 01_모니터링현황에서 오래된 행을 지우면, 그 기사가 네이버 검색에 다시 잡혔을 때
 * "새 기사"로 오인되어 재적재·재알림될 수 있다(F-201) — 사용자가 이 트레이드오프를 알고 요청함.
 */
function deleteHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySh = ss.getSheetByName(SHEETS.HISTORY);
  if (!historySh) { alertSafe_('09_수집이력 시트가 이미 없습니다.'); return; }
  ss.deleteSheet(historySh);
  alertSafe_('09_수집이력 시트를 삭제했습니다.');
}

function getOrCreateBackupFolder_() {
  const parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName('모니터링현황백업');
  return it.hasNext() ? it.next() : parent.createFolder('모니터링현황백업');
}

function pruneOldBackups_(folder, keepWeeks) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(keepWeeks).forEach(f => f.setTrashed(true));
}

// ── F-304 연도별 아카이브 ─────────────────────────────────────────────

/** 연말에 실행: 올해분 모니터링 현황 데이터를 별도 아카이브 스프레드시트로 분리 보관한다(원본은 삭제하지 않음). */
function archiveYear() {
  const ui = SpreadsheetApp.getUi();
  const tz = Session.getScriptTimeZone();
  const defaultYear = Utilities.formatDate(new Date(), tz, 'yyyy');
  const resp = ui.prompt('아카이브할 연도', '예: ' + defaultYear, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const year = Number(resp.getResponseText().trim());
  if (!year) { ui.alert('연도를 확인해 주세요.'); return; }

  const start = new Date(year, 0, 1);
  const end = new Date(year + 1, 0, 1);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('모니터링 현황에 데이터가 없습니다.'); return; }
  const data = sh.getRange(2, 1, lastRow - 1, LEDGER_LAST_COL).getValues();
  const rows = data.filter(r => {
    const pub = r[LEDGER_COL.보도일시 - 1];
    return pub instanceof Date && pub >= start && pub < end;
  });
  if (!rows.length) { ui.alert(year + '년 데이터가 없습니다.'); return; }

  const folder = getOrCreateBackupFolder_();
  const archiveSs = SpreadsheetApp.create('01_모니터링현황_아카이브_' + year);
  DriveApp.getFileById(archiveSs.getId()).moveTo(folder);
  const archiveSh = archiveSs.getSheets()[0];
  archiveSh.setName('01_모니터링현황_' + year);
  archiveSh.getRange(1, 1, 1, LEDGER_LAST_COL).setValues([LEDGER_HEADERS]);
  archiveSh.getRange(2, 1, rows.length, LEDGER_LAST_COL).setValues(rows);

  ui.alert(year + '년 데이터 ' + rows.length + '건을 아카이브했습니다: ' + archiveSs.getUrl() +
    '\n\n원본 01_모니터링현황에서 해당 행을 삭제할지는 담당자가 직접 확인 후 결정해 주세요(자동 삭제하지 않음).');
}

// 고려대 구로병원 전임·임상·진료교원 명단(2026-07 기준, ★전임,임상,진료교원 명단_매월 업데이트 엑셀에서 추출) — seedMedicalStaffDictionary()에서 사용
const MEDICAL_STAFF_SEED_ = [
  ['전인숙', '마취통증의학과'],
  ['정수경', '마취통증의학과'],
  ['김희주', '마취통증의학과'],
  ['박종광', '마취통증의학과'],
  ['오석경', '마취통증의학과'],
  ['임병건', '마취통증의학과'],
  ['김영성', '마취통증의학과'],
  ['이충훈', '마취통증의학과'],
  ['김효성', '마취통증의학과'],
  ['이소영', '마취통증의학과'],
  ['박지은', '마취통증의학과'],
  ['정혜인', '마취통증의학과'],
  ['나소진', '마취통증의학과'],
  ['구은혜', '마취통증의학과'],
  ['권영준', '마취통증의학과'],
  ['문봉기', '마취통증의학과'],
  ['송지현', '마취통증의학과'],
  ['공명훈', '마취통증의학과'],
  ['이일옥', '마취통증의학과'],
  ['김재원', '마취통증의학과'],
  ['이민규', '마취통증의학과'],
  ['김백희', '병리과'],
  ['김정열', '병리과'],
  ['전태성', '병리과'],
  ['김애리', '병리과'],
  ['전이경', '병리과'],
  ['민재용', '병리과'],
  ['권정아', '진단검사의학과'],
  ['윤수영', '진단검사의학과'],
  ['임채승', '진단검사의학과'],
  ['윤정', '진단검사의학과'],
  ['김하늬', '진단검사의학과'],
  ['김보람', '진단검사의학과'],
  ['백만종', '흉부외과'],
  ['김현구', '흉부외과'],
  ['김영수', '심장혈관흉부외과'],
  ['이준희', '심장혈관흉부외과'],
  ['구병모', '심장혈관흉부외과'],
  ['김안나', '피부과'],
  ['전지현', '피부과'],
  ['백유상', '피부과'],
  ['김고은', '피부과'],
  ['신혜선', '영상의학과'],
  ['최재웅', '영상의학과'],
  ['김혜정', '영상의학과'],
  ['이종미', '영상의학과'],
  ['신나리', '영상의학과'],
  ['송명규', '영상의학과'],
  ['박지은', '영상의학과'],
  ['김수진', '영상의학과'],
  ['송초록', '영상의학과'],
  ['김경아', '영상의학과'],
  ['이기열', '영상의학과'],
  ['서태석', '영상의학과'],
  ['정혜나', '영상의학과'],
  ['김정우', '영상의학과'],
  ['박선영', '영상의학과'],
  ['위재연', '영상의학과'],
  ['이주희', '영상의학과'],
  ['이창희', '영상의학과'],
  ['유인선', '영상의학과'],
  ['서상일', '영상의학과'],
  ['우옥희', '영상의학과'],
  ['용환석', '영상의학과'],
  ['송미진', '영상의학과'],
  ['강우영', '영상의학과'],
  ['박빛나', '영상의학과'],
  ['추지영', '영상의학과'],
  ['강은영', '영상의학과'],
  ['설혜영', '영상의학과'],
  ['김명규', '영상의학과'],
  ['심지석', '치과'],
  ['이의석', '치과'],
  ['이수영', '치과'],
  ['이정열', '치과'],
  ['신주희', '치과'],
  ['김예진', '치과'],
  ['정석기', '치과'],
  ['이지은', '치과'],
  ['이한나', '치과'],
  ['임호경', '치과'],
  ['김무경', '치과'],
  ['조석우', '응급의학과'],
  ['조영덕', '응급의학과'],
  ['최성혁', '응급의학과'],
  ['윤영훈', '응급의학과'],
  ['김정윤', '응급의학과'],
  ['박성준', '응급의학과'],
  ['백대현', '응급의학과'],
  ['장진경', '응급의학과'],
  ['노예천', '응급의학과'],
  ['이지영', '응급의학과'],
  ['박광훈', '응급의학과'],
  ['김선미', '가정의학과'],
  ['남가은', '가정의학과'],
  ['최윤선', '가정의학과'],
  ['이유정', '가정의학과'],
  ['박현진', '가정의학과'],
  ['이도희', '가정의학과'],
  ['노현영', '가정의학과'],
  ['이지은', '가정의학과'],
  ['조신제', '외과'],
  ['남명지', '외과'],
  ['이성훈', '외과'],
  ['문두건', '비뇨의학과'],
  ['박홍석', '비뇨의학과'],
  ['오미미', '비뇨의학과'],
  ['김종욱', '비뇨의학과'],
  ['안순태', '비뇨의학과'],
  ['조선범', '비뇨의학과'],
  ['김민경', '건강증진센터'],
  ['이명현', '건강증진센터'],
  ['임영미', '건강증진센터'],
  ['이경훈', '건강증진센터'],
  ['박일순', '건강증진센터'],
  ['안성영', '건강증진센터'],
  ['유경호', '건강증진센터'],
  ['이기현', '내과'],
  ['어재선', '핵의학과'],
  ['이은성', '핵의학과'],
  ['이환희', '핵의학과'],
  ['윤원기', '신경외과'],
  ['김은상', '신경외과'],
  ['김주한', '신경외과'],
  ['김종현', '신경외과'],
  ['이승훈', '신경외과'],
  ['조현준', '신경외과'],
  ['권택현', '신경외과'],
  ['권우근', '신경외과'],
  ['최승진', '신경외과'],
  ['노해원', '신경외과'],
  ['변준호', '신경외과'],
  ['함창화', '신경외과'],
  ['오경미', '신경과'],
  ['고성범', '신경과'],
  ['김지현', '신경과'],
  ['이정윤', '신경과'],
  ['신지혜', '신경과'],
  ['김치경', '신경과'],
  ['이건주', '신경과'],
  ['송민지', '신경과'],
  ['강성훈', '신경과'],
  ['정연학', '신경과'],
  ['이혜림', '신경과'],
  ['이규봉', '신경과'],
  ['홍진화', '산부인과'],
  ['신정호', '산부인과'],
  ['김용진', '산부인과'],
  ['설현주', '산부인과'],
  ['오민정', '산부인과'],
  ['이재관', '산부인과'],
  ['오수민', '산부인과'],
  ['조금준', '산부인과'],
  ['정소현', '산부인과'],
  ['소경아', '산부인과'],
  ['우주현', '산부인과'],
  ['김태현', '산부인과'],
  ['송정민', '산부인과'],
  ['우정수', '이비인후과'],
  ['조재구', '이비인후과'],
  ['박일호', '이비인후과'],
  ['송재준', '이비인후과'],
  ['김영찬', '이비인후과'],
  ['채성원', '이비인후과'],
  ['문지원', '이비인후과'],
  ['박재만', '이비인후과'],
  ['서영우', '안과'],
  ['최광언', '안과'],
  ['백세현', '안과'],
  ['송종석', '안과'],
  ['김우진', '안과'],
  ['강형주', '안과'],
  ['최미현', '안과'],
  ['김승헌', '안과'],
  ['윤수민', '안과'],
  ['김종현', '안과'],
  ['문준규', '정형외과'],
  ['조재우', '정형외과'],
  ['장안성', '정형외과'],
  ['배지훈', '정형외과'],
  ['서승우', '정형외과'],
  ['오종건', '정형외과'],
  ['강성현', '정형외과'],
  ['김학준', '정형외과'],
  ['김상민', '정형외과'],
  ['남윤진', '정형외과'],
  ['최원석', '정형외과'],
  ['이정일', '정형외과'],
  ['박영환', '정형외과'],
  ['현충수', '정형외과'],
  ['임수빈', '정형외과'],
  ['김홍진', '정형외과'],
  ['양승도', '소아청소년과'],
  ['정희라', '소아청소년과'],
  ['김소희', '소아청소년과'],
  ['남효경', '소아청소년과'],
  ['조소윤', '소아청소년과'],
  ['박유진', '소아청소년과'],
  ['하기수', '소아청소년과'],
  ['이은상', '소아청소년과'],
  ['윤상현', '소아청소년과'],
  ['김동현', '소아청소년과'],
  ['송대진', '소아청소년과'],
  ['이경재', '소아청소년과'],
  ['윤윤선', '소아청소년과'],
  ['최의경', '소아청소년과'],
  ['신승현', '소아청소년과'],
  ['양동화', '소아청소년과'],
  ['이치원', '소아청소년과'],
  ['최다민', '소아청소년과'],
  ['최종훈', '소아청소년과'],
  ['한승규', '성형외과'],
  ['동은상', '성형외과'],
  ['정성호', '성형외과'],
  ['남궁식', '성형외과'],
  ['문경철', '성형외과'],
  ['이규일', '성형외과'],
  ['정혜원', '임상약리학과'],
  ['이문수', '정신건강의학과'],
  ['한창수', '정신건강의학과'],
  ['정현강', '정신건강의학과'],
  ['이승훈', '정신건강의학과'],
  ['최원석', '정신건강의학과'],
  ['장문영', '정신건강의학과'],
  ['윤준식', '재활의학과'],
  ['양승남', '재활의학과'],
  ['강석', '재활의학과'],
  ['전소연', '재활의학과'],
  ['양대식', '방사선종양학과'],
  ['이경화', '방사선종양학과'],
  ['김하경', '방사선종양학과'],
  ['김명준', '응급중환자외상외과'],
  ['이상목', '응급중환자외상외과'],
  ['조준민', '응급중환자외상외과'],
  ['이진영', '응급중환자외상외과'],
  ['이재원', '응급중환자외상외과'],
  ['김남렬', '응급중환자외상외과'],
  ['최낙준', '응급중환자외상외과'],
  ['노태욱', '응급중환자외상외과'],
  ['고지울', '응급중환자외상외과'],
  ['김태식', '응급중환자외상외과'],
  ['조민정', '응급중환자외상외과'],
  ['허윤정', '응급중환자외상외과'],
  ['한재민', '응급중환자외상외과'],
  ['우상욱', '유방내분비외과'],
  ['김우영', '유방내분비외과'],
  ['정재현', '유방내분비외과'],
  ['장미영', '유방내분비외과'],
  ['김용엽', '유방내분비외과'],
  ['김지혜', '유방내분비외과'],
  ['김예령', '유방내분비외과'],
  ['민병욱', '대장항문외과'],
  ['강상희', '대장항문외과'],
  ['봉준우', '대장항문외과'],
  ['이선일', '대장항문외과'],
  ['주연욱', '대장항문외과'],
  ['진성기', '대장항문외과'],
  ['김종한', '위장관외과(상부)'],
  ['서원준', '위장관외과(상부)'],
  ['김완배', '간담췌외과'],
  ['최새별', '간담췌외과'],
  ['김완준', '간담췌외과'],
  ['임수연', '간담췌외과'],
  ['한승욱', '간담췌외과'],
  ['나영현', '소아외과'],
  ['박평재', '이식혈관외과'],
  ['김효기', '이식혈관외과'],
  ['권태원', '이식혈관외과'],
  ['최철웅', '순환기내과'],
  ['노승영', '순환기내과'],
  ['이선기', '순환기내과'],
  ['나승운', '순환기내과'],
  ['김진원', '순환기내과'],
  ['최자연', '순환기내과'],
  ['박수형', '순환기내과'],
  ['박은진', '순환기내과'],
  ['이대인', '순환기내과'],
  ['김응주', '순환기내과'],
  ['나진오', '순환기내과'],
  ['이지은', '순환기내과'],
  ['강동오', '순환기내과'],
  ['추원상', '순환기내과'],
  ['곽서연', '순환기내과'],
  ['이경연', '순환기내과'],
  ['이권풍', '순환기내과'],
  ['류혜진', '내분비내과'],
  ['최경묵', '내분비내과'],
  ['장아름', '내분비내과'],
  ['송의연', '내분비내과'],
  ['장수연', '내분비내과'],
  ['박민정', '내분비내과'],
  ['백세현', '내분비내과'],
  ['조문경', '내분비내과'],
  ['김효정', '소화기내과'],
  ['김재선', '소화기내과'],
  ['주문경', '소화기내과'],
  ['김지훈', '소화기내과'],
  ['이영선', '소화기내과'],
  ['연종은', '소화기내과'],
  ['김원식', '소화기내과'],
  ['이범재', '소화기내과'],
  ['김승한', '소화기내과'],
  ['최은호', '소화기내과'],
  ['유양재', '소화기내과'],
  ['최성지', '소화기내과'],
  ['박경선', '혈액종양내과'],
  ['김대식', '혈액종양내과'],
  ['강은주', '혈액종양내과'],
  ['장수영', '혈액종양내과'],
  ['권혜미', '혈액종양내과'],
  ['최철원', '혈액종양내과'],
  ['박인혜', '혈액종양내과'],
  ['서재홍', '혈액종양내과'],
  ['전민지', '혈액종양내과'],
  ['오상철', '혈액종양내과'],
  ['유은상', '혈액종양내과'],
  ['남궁윤', '혈액종양내과'],
  ['이경민', '혈액종양내과'],
  ['윤진구', '감염내과'],
  ['정희진', '감염내과'],
  ['송준영', '감염내과'],
  ['최민주', '감염내과'],
  ['노지윤', '감염내과'],
  ['남엘리엘', '감염내과'],
  ['김우주', '감염내과'],
  ['고강지', '신장내과'],
  ['권영주', '신장내과'],
  ['조은정', '신장내과'],
  ['김지은', '신장내과'],
  ['김효진', '신장내과'],
  ['송승민', '신장내과'],
  ['강민우', '신장내과'],
  ['심재겸', '호흡기내과'],
  ['박소윤', '호흡기내과'],
  ['허규영', '호흡기내과'],
  ['이승룡', '호흡기내과'],
  ['민경훈', '호흡기내과'],
  ['오지연', '호흡기내과'],
  ['장성원', '호흡기내과'],
  ['최주환', '호흡기내과'],
  ['김상혁', '호흡기내과'],
  ['심재정', '호흡기내과'],
  ['최인아', '류마티스내과'],
  ['송관규', '류마티스내과'],
  ['장성혜', '류마티스내과']
];
