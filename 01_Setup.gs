/**
 * 01_Setup.gs
 * 시트 최초 생성/초기화 및 스프레드시트 커스텀 메뉴.
 * 담당자는 새 구글 스프레드시트를 만들고 이 프로젝트의 .gs 파일을 모두 붙여넣은 뒤,
 * 메뉴의 [모니터링 시스템 → 1. 초기 설정 실행]을 한 번만 실행하면 된다.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('모니터링 시스템')
    .addItem('1. 초기 설정 실행 (최초 1회)', 'initializeSystem')
    .addItem('2. 트리거 설치 (최초 1회)', 'installTriggers')
    .addSeparator()
    .addItem('지금 즉시 수집 실행', 'collectAll')
    .addItem('지면·방송 보도 수동 등록', 'addManualArticle')
    .addItem('대시보드 새로고침', 'refreshDashboard')
    .addItem('주간 통계 재계산', 'buildWeeklyStats')
    .addItem('월간 통계 재계산', 'buildMonthlyStats')
    .addItem('연간 통계 재계산', 'buildYearlyStats')
    .addItem('기간 임의 조회', 'runCustomRangeQuery')
    .addSeparator()
    .addItem('선택한 기사에 수동 알림 발송', 'sendManualAlertForSelectedRow')
    .addItem('선택한 기사를 오탐으로 표시', 'markSelectedRowFalsePositive')
    .addItem('주간 리포트 메일 발송(수동)', 'sendWeeklyReportEmail')
    .addItem('월간 리포트 메일 발송(수동)', 'sendMonthlyReportEmail')
    .addSeparator()
    .addItem('보도자료 채택 매칭 실행', 'matchPressReleases')
    .addSeparator()
    .addItem('매체 정보 재계산(유지보수)', 'renormalizeLedgerMedia')
    .addItem('네이버뉴스 매체명 재확인(유지보수)', 'backfillNaverOriginalMedia')
    .addItem('의료진 태그 재계산(유지보수)', 'backfillMedicalStaff')
    .addSeparator()
    .addItem('연간 아카이브 실행', 'archiveYear')
    .addItem('모니터링 현황 백업 실행', 'backupLedger')
    .addToUi();
}

/**
 * 최초 1회 실행: 8개 시트를 생성하고 헤더·드롭다운·서식·기본 설정값을 채운다.
 * 이미 존재하는 시트는 건드리지 않는다(재실행해도 안전).
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupConfigSheet_(ss);
  setupLedgerSheet_(ss);
  setupPressSheet_(ss);
  setupSimpleSheet_(ss, SHEETS.ALERT_LOG, ALERT_LOG_HEADERS);
  setupSimpleSheet_(ss, SHEETS.LOG, LOG_HEADERS);
  setupSimpleSheet_(ss, SHEETS.HISTORY, HISTORY_HEADERS);
  setupStatsPlaceholder_(ss, SHEETS.WEEKLY, '주간통계는 [모니터링 시스템 → 주간 통계 재계산]으로 생성됩니다.');
  setupStatsPlaceholder_(ss, SHEETS.MONTHLY, '월간통계는 [모니터링 시스템 → 월간 통계 재계산]으로 생성됩니다.');
  setupStatsPlaceholder_(ss, SHEETS.YEARLY, '연간통계는 [모니터링 시스템 → 연간 통계 재계산]으로 생성됩니다.');
  setupStatsPlaceholder_(ss, SHEETS.DASHBOARD, '대시보드는 [모니터링 시스템 → 대시보드 새로고침]으로 생성됩니다.');

  // 시트 탭 순서를 명세서 2.3항 순서로 정렬
  const order = [SHEETS.CONFIG, SHEETS.LEDGER, SHEETS.PRESS, SHEETS.WEEKLY,
    SHEETS.MONTHLY, SHEETS.YEARLY, SHEETS.DASHBOARD, SHEETS.ALERT_LOG, SHEETS.LOG, SHEETS.HISTORY];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });

  // 기본 시트(시트1 등) 정리 — 위 9개 외 다른 시트가 있으면 안내만 하고 삭제하지 않음
  SpreadsheetApp.getUi().alert('초기 설정 완료',
    '00_설정 시트에서 병원명·키워드·매체 목록·위기 키워드·수신자·API 키 위치를 확인/입력한 뒤,\n' +
    '[프로젝트 설정 → 스크립트 속성]에 NAVER_CLIENT_ID 등 API 키를 등록하고,\n' +
    '[모니터링 시스템 → 2. 트리거 설치]를 실행하세요.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupConfigSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.CONFIG);
  if (sh) return; // 이미 있으면 스킵 (재실행 안전)
  sh = ss.insertSheet(SHEETS.CONFIG);
  sh.getRange('A1').setValue('■ 00_설정 — 검색 키워드 / 매체 마스터 / 위기 키워드 사전 / 수신자 / RSS / 임계값')
    .setFontWeight('bold').setFontSize(12);

  writeBlockHeader_(sh, CFG_BLOCK.키워드, '검색 키워드 (3.3항)');
  writeBlockHeader_(sh, CFG_BLOCK.제외어, '제외어 (오수집 필터, F-204)');
  writeBlockHeader_(sh, CFG_BLOCK.매체마스터, '매체 마스터 (F-203)');
  writeBlockHeader_(sh, CFG_BLOCK.위기키워드, '위기 키워드 사전 (6.1항)');
  writeBlockHeader_(sh, CFG_BLOCK.수신자, '알림 수신자 (6.4항, F-605)');
  writeBlockHeader_(sh, CFG_BLOCK.RSS피드, 'RSS 피드 목록 (F-102)');
  writeBlockHeader_(sh, CFG_BLOCK.임계값, '임계값·운영 파라미터 (6.3항 등)');
  writeBlockHeader_(sh, CFG_BLOCK.논조키워드, '논조 판정 키워드 사전 (F-205)');
  writeBlockHeader_(sh, CFG_BLOCK.카테고리키워드, '카테고리 분류 키워드 (F-206)');
  writeBlockHeader_(sh, CFG_BLOCK.진료과키워드, '진료과·센터 태깅 키워드 (F-207)');
  writeBlockHeader_(sh, CFG_BLOCK.오탐예외, '오탐 예외 키워드 (F-512, 자동 축적됨)');
  writeBlockHeader_(sh, CFG_BLOCK.의료진사전, '의료진 사전 — 기사 매칭용 (이름·소속 직접 입력)');

  // 샘플/기본 데이터 채우기
  seedBlock_(sh, CFG_BLOCK.키워드, [
    ['기관명', '고려대 구로병원', 'Y'], ['기관명', '고대구로병원', 'Y'],
    ['기관명', '고려대학교 구로병원', 'Y'], ['기관명', '구로병원', 'Y'],
    ['상위기관', '고려대학교의료원', 'Y'], ['상위기관', '고대의료원', 'Y']
  ]);
  seedBlock_(sh, CFG_BLOCK.제외어, [['구로구청', 'Y'], ['구로디지털단지', 'Y']]);
  seedBlock_(sh, CFG_BLOCK.매체마스터, [
    ['연합뉴스', '연합뉴스', '통신사', '1급'],
    ['청년의사', '청년의사', '의학전문지', '3급'],
    ['데일리메디', '데일리메디', '의학전문지', '3급'],
    ['메디칼타임즈', '메디칼타임즈', '의학전문지', '3급'],
    ['메디게이트뉴스', '메디게이트뉴스', '의학전문지', '3급']
  ]);
  seedBlock_(sh, CFG_BLOCK.위기키워드, [
    ['의료사고·분쟁', '의료사고', 3], ['의료사고·분쟁', '의료과실', 3], ['의료사고·분쟁', '오진', 3],
    ['의료사고·분쟁', '수술 부작용', 3], ['의료사고·분쟁', '사망', 3], ['의료사고·분쟁', '손해배상', 3],
    ['의료사고·분쟁', '의료소송', 3], ['의료사고·분쟁', '의료분쟁조정', 3],
    ['수사·행정처분', '압수수색', 3], ['수사·행정처분', '고소', 3], ['수사·행정처분', '고발', 3],
    ['수사·행정처분', '기소', 3], ['수사·행정처분', '검찰', 3], ['수사·행정처분', '경찰 수사', 3],
    ['수사·행정처분', '행정처분', 3], ['수사·행정처분', '과징금', 3], ['수사·행정처분', '업무정지', 3],
    ['수사·행정처분', '실사', 3],
    ['감염·안전', '집단감염', 3], ['감염·안전', '원내감염', 3], ['감염·안전', '슈퍼박테리아', 3],
    ['감염·안전', '격리', 3], ['감염·안전', '방역 실패', 3], ['감염·안전', '화재', 3],
    ['감염·안전', '정전', 3], ['감염·안전', '낙상', 3],
    ['부정·비리', '리베이트', 3], ['부정·비리', '횡령', 3], ['부정·비리', '배임', 3],
    ['부정·비리', '진료비 부당청구', 3], ['부정·비리', '허위청구', 3], ['부정·비리', '채용비리', 3],
    ['부정·비리', '특혜', 3],
    ['인물 리스크', '갑질', 3], ['인물 리스크', '성희롱', 3], ['인물 리스크', '폭언', 3],
    ['인물 리스크', '폭행', 3], ['인물 리스크', '음주', 3], ['인물 리스크', '논문 조작', 3],
    ['인물 리스크', '연구부정', 3], ['인물 리스크', '징계', 3],
    ['노무·조직', '파업', 2], ['노무·조직', '노조', 2], ['노무·조직', '태업', 2],
    ['노무·조직', '집단행동', 2], ['노무·조직', '임단협 결렬', 2], ['노무·조직', '진료 차질', 2],
    ['노무·조직', '전공의', 2],
    ['평판·불만', '불만', 2], ['평판·불만', '항의', 2], ['평판·불만', '민원', 2],
    ['평판·불만', '환자 거부', 2], ['평판·불만', '대기시간', 2], ['평판·불만', '불친절', 2],
    ['평판·불만', '폭리', 2], ['평판·불만', '청구서 논란', 2],
    ['의혹 표현', '의혹', 1], ['의혹 표현', '논란', 1], ['의혹 표현', '파문', 1],
    ['의혹 표현', '충격', 1], ['의혹 표현', '물의', 1], ['의혹 표현', '도마', 1],
    ['의혹 표현', '뭇매', 1], ['의혹 표현', '해명', 1]
  ]);
  seedBlock_(sh, CFG_BLOCK.RSS피드, [
    ['청년의사', 'https://www.docdocdoc.co.kr/rssIndex.html', 'Y'],
    ['구글알리미(고려대구로병원)', '여기에 구글 알리미에서 발급받은 RSS URL을 붙여넣으세요', 'N']
  ]);
  seedBlock_(sh, CFG_BLOCK.임계값, DEFAULT_THRESHOLDS);
  seedBlock_(sh, CFG_BLOCK.논조키워드, [
    ['긍정', '우수'], ['긍정', '성과'], ['긍정', '혁신'], ['긍정', '최초'], ['긍정', '선정'],
    ['긍정', '수상'], ['긍정', '인증'], ['긍정', '완치'], ['긍정', '성공적'], ['긍정', '기부'],
    ['긍정', '봉사'], ['긍정', '감사'], ['긍정', '협약'], ['긍정', '지정'], ['긍정', '개원'],
    ['부정', '사망'], ['부정', '의혹'], ['부정', '논란'], ['부정', '과실'], ['부정', '소송'],
    ['부정', '부작용'], ['부정', '피해'], ['부정', '항의'], ['부정', '파업'], ['부정', '적자'],
    ['부정', '불만'], ['부정', '사고'], ['부정', '감염'], ['부정', '해명']
  ]);
  seedBlock_(sh, CFG_BLOCK.카테고리키워드, [
    ['진료·의료기술', '수술'], ['진료·의료기술', '시술'], ['진료·의료기술', '진료'], ['진료·의료기술', '치료'],
    ['진료·의료기술', '이식'], ['진료·의료기술', '로봇수술'], ['진료·의료기술', '신약'], ['진료·의료기술', '검진'],
    ['연구·학술', '연구'], ['연구·학술', '논문'], ['연구·학술', '학회'], ['연구·학술', '임상시험'], ['연구·학술', '특허'],
    ['경영·정책', '병상'], ['경영·정책', '평가'], ['경영·정책', '인증'], ['경영·정책', '투자'], ['경영·정책', '수가'], ['경영·정책', '협약'],
    ['인물·인사', '임명'], ['인물·인사', '취임'], ['인물·인사', '교수'], ['인물·인사', '원장'], ['인물·인사', '인터뷰'],
    ['사회공헌', '기부'], ['사회공헌', '봉사'], ['사회공헌', '후원'], ['사회공헌', '나눔'], ['사회공헌', '무료진료'],
    ['사건·사고', '사망'], ['사건·사고', '화재'], ['사건·사고', '감염'], ['사건·사고', '사고'], ['사건·사고', '소송'], ['사건·사고', '파업']
  ]);
  seedBlock_(sh, CFG_BLOCK.진료과키워드, [
    ['권역응급의료센터', '응급실'], ['권역응급의료센터', '응급의학과'], ['권역응급의료센터', '권역응급의료센터'],
    ['심혈관센터', '심혈관센터'], ['심혈관센터', '심장혈관흉부외과'], ['심혈관센터', '순환기내과'],
    ['로봇수술센터', '로봇수술'], ['로봇수술센터', '다빈치'],
    ['정형외과', '정형외과'], ['신경외과', '신경외과'], ['산부인과', '산부인과'],
    ['소아청소년과', '소아청소년과'], ['소아청소년과', '소아과']
  ]);

  // 열 너비 (기본 시트는 26열뿐이므로 필요한 만큼 열을 먼저 확보한다 — 의료진사전이 AT열까지 사용)
  if (sh.getMaxColumns() < 47) sh.insertColumnsAfter(sh.getMaxColumns(), 47 - sh.getMaxColumns());
  sh.setColumnWidths(1, 47, 100);
  sh.setFrozenRows(2);

  // 안내 노트
  sh.getRange('A1').setNote(
    '이 시트는 코드 수정 없이 운영 파라미터를 바꾸는 곳입니다(F-601).\n' +
    'API 키·발송 계정 비밀정보는 절대 이 시트에 적지 마세요 — ' +
    'Apps Script 편집기의 [프로젝트 설정 → 스크립트 속성]에 저장합니다.');
}

function writeBlockHeader_(sh, block, title) {
  const startRange = sh.getRange(block.header);
  const row = startRange.getRow();
  const col = startRange.getColumn();
  sh.getRange(row, col).setValue(title).setFontWeight('bold').setBackground('#FFF8E1');
  sh.getRange(row + 1, col, 1, block.cols.length).setValues([block.cols])
    .setFontWeight('bold').setBackground('#F1F3F4');
}

function seedBlock_(sh, block, rows) {
  if (!rows.length) return;
  const startRange = sh.getRange(block.dataStart);
  sh.getRange(startRange.getRow(), startRange.getColumn(), rows.length, block.cols.length)
    .setValues(rows);
}

function setupLedgerSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.LEDGER);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.LEDGER);
  sh.getRange(1, 1, 1, LEDGER_LAST_COL).setValues([LEDGER_HEADERS])
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF');
  sh.getRange(1, LEDGER_COL.보도일시).setValue('보도날짜'); // 내부 키 이름은 보도일시로 유지, 화면 표시만 "보도날짜"
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  hideLedgerColumns_(sh); // 사용자 요청: 화면에는 불필요, 내부 로직(위기판정·알림·통계·매칭 등)용으로는 유지

  const maxRows = 5000;
  if (sh.getMaxRows() < maxRows) sh.insertRowsAfter(sh.getMaxRows(), maxRows - sh.getMaxRows());

  // 사용자 요청(2026-10): 보도일시는 시간 없이 날짜만 표시
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

  // 상태 컬럼 조건부 서식 (미확인=노랑, 확정=녹색, 제외/오탐=회색)
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
}

/**
 * 사용자 요청: 화면에는 핵심 항목만 보이도록 나머지 컬럼을 숨긴다.
 * 데이터 자체는 지우지 않으므로 위기판정·알림 중복방지·통계(상태=확정 집계)·보도자료 매칭 로직은 그대로 동작한다.
 * 필요하면 시트에서 언제든 숨김을 해제할 수 있다.
 */
function hideLedgerColumns_(sh) {
  const hide = [
    LEDGER_COL.기사ID, LEDGER_COL.수집일시, LEDGER_COL.매체영향력, LEDGER_COL.요약,
    LEDGER_COL.논조확정, LEDGER_COL.이슈그룹ID, LEDGER_COL.보도자료ID, LEDGER_COL.노출도,
    LEDGER_COL.수집경로, LEDGER_COL.상태, LEDGER_COL.대응상태, LEDGER_COL.조치내역,
    LEDGER_COL.담당자, LEDGER_COL.비고,
    LEDGER_COL.URL, LEDGER_COL.매체유형, LEDGER_COL.검색키워드, LEDGER_COL.카테고리, LEDGER_COL.보도유형,
    LEDGER_COL.기자명, LEDGER_COL.진료과센터, LEDGER_COL.논조자동
  ].sort((a, b) => a - b);

  let runStart = hide[0], runLen = 1;
  for (let i = 1; i <= hide.length; i++) {
    if (i < hide.length && hide[i] === runStart + runLen) { runLen++; continue; }
    sh.hideColumns(runStart, runLen);
    if (i < hide.length) { runStart = hide[i]; runLen = 1; }
  }
}

/**
 * 1회성 등록용(사용자 제공 명단). 00_설정!매체마스터에 이미 있는 항목은 건드리지 않고,
 * 아래 목록을 기존 데이터 다음 행부터 이어서 추가한다(중복으로 다시 실행해도 계속 뒤에 쌓이므로 한 번만 실행).
 */
function seedAdditionalMediaMaster() {
  const rows = [
    ['siminilbo.co.kr', '시민일보', '일간&경제지', '3급'],
    ['biospectator.com', '바이오스펙테이터', '의학전문지', '3급'],
    ['khan.co.kr', '경향신문', '일간&경제지', '1급'],
    ['rapportian.com', '라포르시안', '의학전문지', '3급'],
    ['donga.com', '동아일보', '일간&경제지', '1급'],
    ['biz.heraldcorp.com', '헤럴드경제', '일간&경제지', '2급'],
    ['cnbizm.com', '문화경제', '일간&경제지', '3급'],
    ['sisaon.co.kr', '시사오늘(시사ON)', '일간&경제지', '3급'],
    ['weekly.hankooki.com', '주간한국', '일간&경제지', '3급'],
    ['news1.kr', '뉴스1', '통신사', '1급'],
    ['mdtoday.co.kr', '메디컬투데이', '의학전문지', '3급'],
    ['mkhealth.co.kr', '매경헬스', '의학전문지', '3급'],
    ['the-stock.kr', '더스탁', '일간&경제지', '3급'],
    ['newspim.com', '뉴스핌', '일간&경제지', '2급'],
    ['asiatime.co.kr', '아시아타임즈', '일간&경제지', '3급'],
    ['kukinews.com', '쿠키뉴스', '일간&경제지', '3급'],
    ['doctorstimes.com', '의사신문', '의학전문지', '3급'],
    ['m-i.kr', '매일일보', '일간&경제지', '3급'],
    ['apsk.co.kr', '메디팜스투데이', '의학전문지', '3급'],
    ['bokuennews.com', '보건뉴스', '의학전문지', '3급'],
    ['hankyung.com', '한국경제', '일간&경제지', '1급'],
    ['veritas-a.com', '베리타스알파', '일간&경제지', '3급'],
    ['shinailbo.co.kr', '신아일보', '일간&경제지', '3급'],
    ['munhwa.com', '문화일보', '일간&경제지', '2급'],
    ['medicalworldnews.co.kr', '메디컬월드뉴스', '의학전문지', '3급'],
    ['sportschosun.com', '스포츠조선', '일간&경제지', '2급'],
    ['kpanews.co.kr', '약사공론', '의학전문지', '3급'],
    ['joongang.co.kr', '중앙일보', '일간&경제지', '1급'],
    ['newsmp.com', '의약뉴스', '의학전문지', '3급'],
    ['topstarnews.net', '톱스타뉴스', '일간&경제지', '3급'],
    ['news.cpbc.co.kr', '가톨릭평화방송 뉴스', '일간&경제지', '3급'],
    ['lawissue.co.kr', '로이슈', '일간&경제지', '3급'],
    ['hellodd.com', '헬로디디', '일간&경제지', '3급'],
    ['imaeil.com', '매일신문', '일간&경제지', '2급'],
    ['asiatoday.co.kr', '아시아투데이', '일간&경제지', '2급'],
    ['seoul.co.kr', '서울신문', '일간&경제지', '2급'],
    ['catchnews.kr', '캐치뉴스', '일간&경제지', '3급'],
    ['koreahealthlog.com', '코리아헬스로그', '의학전문지', '3급'],
    ['joongangenews.com', '중앙이코노미뉴스', '일간&경제지', '3급'],
    ['newsway.co.kr', '뉴스웨이', '일간&경제지', '3급'],
    ['theviewers.co.kr', '더뷰어스', '일간&경제지', '3급'],
    ['mt.co.kr', '머니투데이', '일간&경제지', '2급'],
    ['thefirstmedia.net', '더퍼스트미디어', '일간&경제지', '3급'],
    ['businessplus.kr', '비즈니스플러스', '일간&경제지', '3급'],
    ['beyondpost.co.kr', '비욘드포스트', '일간&경제지', '3급'],
    ['dailian.co.kr', '데일리안', '일간&경제지', '2급'],
    ['fntimes.com', '한국금융신문', '일간&경제지', '3급'],
    ['sidae.com', '시대', '일간&경제지', '3급'],
    ['ftoday.co.kr', '파이낸셜투데이', '일간&경제지', '3급'],
    ['insight.co.kr', '인사이트', '일간&경제지', '3급'],
    ['kyosu.net', '교수신문', '일간&경제지', '3급'],
    ['edaily.co.kr', '이데일리', '일간&경제지', '2급'],
    ['viva100.com', '브릿지경제', '일간&경제지', '3급'],
    ['asiae.co.kr', '아시아경제', '일간&경제지', '2급'],
    ['e-science.co.kr', '이코노미사이언스', '일간&경제지', '3급'],
    ['ohmynews.com', '오마이뉴스', '일간&경제지', '2급'],
    ['medical-tribune.co.kr', '메디칼트리뷴', '의학전문지', '3급'],
    ['medipana.com', '메디파나뉴스', '의학전문지', '3급'],
    ['hansbiz.co.kr', '한스경제', '일간&경제지', '3급'],
    ['etnews.com', '전자신문', '일간&경제지', '2급'],
    ['sentv.co.kr', '서울경제TV', '일간&경제지', '3급'],
    ['health.chosun.com', '헬스조선', '의학전문지', '2급'],
    ['bosa.co.kr', '보건뉴스', '의학전문지', '3급'],
    ['thevaluenews.co.kr', '더밸류뉴스', '일간&경제지', '3급'],
    ['newsfreezone.co.kr', '뉴스프리존', '일간&경제지', '3급'],
    ['pointdaily.co.kr', '포인트데일리', '일간&경제지', '3급'],
    ['financialpost.co.kr', '파이낸셜포스트', '일간&경제지', '3급'],
    ['fnnews.com', '파이낸셜뉴스', '일간&경제지', '2급'],
    ['g-enews.com', '글로벌이코노믹', '일간&경제지', '3급'],
    ['getnews.co.kr', '글로벌경제신문', '일간&경제지', '3급'],
    ['businesskorea.co.kr', 'BusinessKorea', '일간&경제지', '3급'],
    ['bizwork.co.kr', '비즈워크', '일간&경제지', '3급'],
    ['daily.hankooki.com', '데일리한국', '일간&경제지', '3급'],
    ['monews.co.kr', '메디칼업저버', '의학전문지', '3급'],
    ['mydaily.co.kr', '마이데일리', '일간&경제지', '3급'],
    ['newsis.com', '뉴시스', '통신사', '1급'],
    ['mediapen.com', '미디어펜', '일간&경제지', '3급'],
    ['econovill.com', '이코노믹리뷰', '일간&경제지', '3급'],
    ['it.chosun.com', 'IT조선', '일간&경제지', '2급'],
    ['ikld.kr', '국토일보', '일간&경제지', '3급'],
    ['medisobizanews.com', '메디소비자뉴스', '의학전문지', '3급'],
    ['whosaeng.com', '후생신보', '의학전문지', '3급'],
    ['press9.kr', '프레스나인', '일간&경제지', '3급'],
    ['k-health.com', '헬스경향', '의학전문지', '3급'],
    ['thefairnews.co.kr', '더페어', '일간&경제지', '3급'],
    ['pinpointnews.co.kr', '핀포인트뉴스', '일간&경제지', '3급']
  ];
  const sh = getConfigSheet_();
  const block = CFG_BLOCK.매체마스터;
  const startRange = sh.getRange(block.dataStart);
  const existing = readConfigBlock_('매체마스터');
  const existingDomains = new Set(existing.map(r => String(r['수집매체명']).trim().toLowerCase()));
  const toAdd = rows.filter(r => !existingDomains.has(r[0].toLowerCase()));
  if (!toAdd.length) { alertSafe_('추가할 신규 매체가 없습니다(이미 모두 등록됨).'); return; }
  const targetRow = startRange.getRow() + existing.length;
  sh.getRange(targetRow, startRange.getColumn(), toAdd.length, 4).setValues(toAdd);
  alertSafe_('매체 마스터에 ' + toAdd.length + '건을 추가했습니다(기존 ' + (rows.length - toAdd.length) + '건은 중복이라 건너뜀).');
}

function setupPressSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.PRESS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.PRESS);
  sh.getRange(1, 1, 1, PRESS_HEADERS.length).setValues([PRESS_HEADERS])
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function setupSimpleSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (sh) return;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function setupStatsPlaceholder_(ss, name, note) {
  let sh = ss.getSheetByName(name);
  if (sh) return;
  sh = ss.insertSheet(name);
  sh.getRange('A1').setValue(note).setFontStyle('italic').setFontColor('#888888');
}

/**
 * 1회성 마이그레이션. 이미 만들어진 01_모니터링현황 시트(구 스키마)에 의료진 컬럼을 추가하고,
 * 요청한 컬럼 숨김·보도일시 시간표시 서식을 적용한다. 스크립트 편집기에서 함수를 선택해 한 번만 실행한다.
 */
function migrateLedgerSchemaV2() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEDGER);
  if (!sh) { SpreadsheetApp.getUi().alert('01_모니터링현황 시트를 찾을 수 없습니다.'); return; }

  if (sh.getMaxColumns() < LEDGER_LAST_COL) {
    sh.insertColumnsAfter(sh.getMaxColumns(), LEDGER_LAST_COL - sh.getMaxColumns());
  }
  const headerCell = sh.getRange(1, LEDGER_COL.의료진);
  if (!headerCell.getValue()) {
    headerCell.setValue('의료진').setFontWeight('bold').setBackground('#4A6FA5').setFontColor('#FFFFFF');
  }
  headerCell.setNote('00_설정!의료진 사전에 등록된 이름이 제목·요약에서 발견되면 자동으로 채워집니다. 등록 전에는 비어 있습니다.');

  hideLedgerColumns_(sh);

  const maxRows = sh.getMaxRows();
  if (maxRows > 1) {
    sh.getRange(2, LEDGER_COL.보도일시, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  SpreadsheetApp.getUi().alert('마이그레이션 완료', '의료진 컬럼 추가, 열 숨김, 보도일시 시간표시 서식을 적용했습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function applyValidation_(sh, col, maxRows, list) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
  sh.getRange(2, col, maxRows - 1, 1).setDataValidation(rule);
}
