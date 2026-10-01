/**
 * 00_Config.gs
 * 언론보도 실시간 모니터링 및 통계 자동화 시스템 — 전역 설정
 * 고려대학교 구로병원 홍보팀
 *
 * 이 파일은 시트명, 컬럼 정의, 00_설정 시트의 각 블록 위치, 임계값 기본값을
 * 한 곳에서 관리한다. 실제 API 키·발송 계정 등 민감정보는 이 파일이 아니라
 * [Apps Script 편집기 → 프로젝트 설정 → 스크립트 속성]에 저장한다.
 */

const SHEETS = {
  CONFIG: '00_설정',
  LEDGER: '01_모니터링현황',
  PRESS: '02_보도자료',
  WEEKLY: '03_주간통계',
  MONTHLY: '04_월간통계',
  YEARLY: '05_연간통계',
  DASHBOARD: '06_대시보드',
  ALERT_LOG: '07_알림이력',
  LOG: '08_로그',
  HISTORY: '09_수집이력'
};

/**
 * 01_모니터링현황 컬럼 정의 (제5장). 1-based index.
 * 사용자 요청(2026-10)으로 화면 맨 앞쪽 컬럼 순서를 보도일시(보도날짜로 표시)·매체명·의료진·진료과·제목
 * 순으로 재배치했다. 나머지(기자명·진료과센터·논조자동·위기등급)는 그 뒤에, URL과 매체유형/검색키워드/
 * 카테고리/보도유형(사용자 요청으로 숨김 처리)은 더 뒤에 둔다. 키 이름 자체는 기존 코드 호환을 위해 그대로 둔다.
 */
const LEDGER_COL = {
  보도일시: 1,
  매체명: 2,
  의료진: 3,
  진료과: 4,
  제목: 5,
  기자명: 6,
  진료과센터: 7,
  논조자동: 8,
  위기등급: 9,
  URL: 10,
  매체유형: 11,
  검색키워드: 12,
  카테고리: 13,
  보도유형: 14,
  기사ID: 15,
  수집일시: 16,
  매체영향력: 17,
  요약: 18,
  논조확정: 19,
  이슈그룹ID: 20,
  보도자료ID: 21,
  노출도: 22,
  수집경로: 23,
  상태: 24,
  대응상태: 25,
  조치내역: 26,
  담당자: 27,
  비고: 28
};
const LEDGER_HEADERS = Object.keys(LEDGER_COL);
const LEDGER_LAST_COL = LEDGER_HEADERS.length;

const PRESS_HEADERS = [
  '보도자료ID', '배포일시', '제목', '주제·진료과', '배포 매체 수',
  '기사화 건수', '채택 매체', '채택률', '담당자'
];

const ALERT_LOG_HEADERS = [
  '알림ID', '발송일시', '기사ID', '이슈그룹ID', '위기등급', '위기점수',
  '수신자', '채널', '발송결과', '메시지본문'
];

const LOG_HEADERS = ['실행일시', '작업', '호출건수', '신규저장건수', '오류내용', '소요시간(ms)'];

/**
 * 사용자 요청(2026-09-30): 01_모니터링현황에서 옛날 기사를 삭제해도, 네이버 검색 결과에 그 기사가
 * 다시 잡히면 "처음 보는 새 기사"로 오인해 재적재·재알림되는 사고가 있었다(F-201 중복판정이
 * 화면에 보이는 01_모니터링현황 데이터만 기준으로 동작했기 때문). 이를 막기 위해 한 번이라도
 * 적재된 기사의 dedupKey를 01_모니터링현황과 별개로, 절대 삭제하지 않는 이 시트에 영구 보관한다.
 */
const HISTORY_HEADERS = ['기사ID', 'dedupKey', '보도일시', '수집일시'];

// 상태·선택값 드롭다운
const OPTS = {
  매체유형: ['통신사', '의학전문지', '일간&경제지'],
  매체영향력: ['1급', '2급', '3급', '커뮤니티'],
  카테고리: ['진료·의료기술', '연구·학술', '경영·정책', '인물·인사', '사회공헌', '사건·사고', '기타'],
  논조: ['긍정', '중립', '부정'],
  위기등급: ['해당없음', '관심', '주의', '경계', '심각'],
  보도유형: ['보도자료 기반', '인터뷰·기고', '기획·단독', '인용·코멘트', '자체취재'],
  노출도: ['포털메인', '네이버뉴스', '일반', '지면', '방송'],
  수집경로: ['API', 'RSS', '커뮤니티', '수동'],
  상태: ['미확인', '확정', '제외', '오탐'],
  대응상태: ['해당없음', '확인', '대응중', '정정요청', '종결']
};

/**
 * 네이버 뉴스 미러 링크(n.news.naver.com/mnews/article/<코드>/...)의 <코드>를 실제 언론사 도메인으로 되돌리는 사전.
 * 원장 URL 칸에는 네이버 미러 주소만 저장되므로, 매체 정보를 나중에 다시 계산할 때(renormalizeLedgerMedia)
 * 이 코드가 없으면 원 매체를 알아낼 방법이 없어 "네이버뉴스"로만 표시된다. 잘못 분류된 매체를 발견하면
 * 기사 URL의 article/뒤 숫자(언론사 코드)를 확인해 여기에 추가해 주면 이후 재계산부터 정확히 잡힌다.
 */
const NAVER_PRESS_CODE_MAP = {
  '277': 'asiae.co.kr' // 아시아경제 (2026-09 사용자 확인)
};

/**
 * F-203 매체 정규화 보조 사전. 00_설정!매체마스터에 등록되지 않은 도메인이 나왔을 때
 * URL 도메인을 그대로 매체명으로 쓰지 않도록 자주 나오는 언론사를 기본 매핑해둔다.
 * 00_설정!매체마스터에 같은 수집매체명을 등록하면 그 값이 항상 우선한다(이 사전은 최후 보완용).
 */
const DOMAIN_MEDIA_FALLBACK = {
  // 통신사
  'yna.co.kr': { name: '연합뉴스', type: '통신사', influence: '1급' },
  'news1.kr': { name: '뉴스1', type: '통신사', influence: '1급' },
  'newsis.com': { name: '뉴시스', type: '통신사', influence: '1급' },
  // 일간지·경제지
  'chosun.com': { name: '조선일보', type: '일간&경제지', influence: '1급' },
  'donga.com': { name: '동아일보', type: '일간&경제지', influence: '1급' },
  'joongang.co.kr': { name: '중앙일보', type: '일간&경제지', influence: '1급' },
  'hani.co.kr': { name: '한겨레', type: '일간&경제지', influence: '1급' },
  'khan.co.kr': { name: '경향신문', type: '일간&경제지', influence: '1급' },
  'hankookilbo.com': { name: '한국일보', type: '일간&경제지', influence: '2급' },
  'seoul.co.kr': { name: '서울신문', type: '일간&경제지', influence: '2급' },
  'mk.co.kr': { name: '매일경제', type: '일간&경제지', influence: '1급' },
  'hankyung.com': { name: '한국경제', type: '일간&경제지', influence: '1급' },
  'mt.co.kr': { name: '머니투데이', type: '일간&경제지', influence: '2급' },
  'fnnews.com': { name: '파이낸셜뉴스', type: '일간&경제지', influence: '2급' },
  'edaily.co.kr': { name: '이데일리', type: '일간&경제지', influence: '2급' },
  'etnews.com': { name: '전자신문', type: '일간&경제지', influence: '2급' },
  'asiae.co.kr': { name: '아시아경제', type: '일간&경제지', influence: '2급' },
  'view.asiae.co.kr': { name: '아시아경제', type: '일간&경제지', influence: '2급' },
  'newsway.co.kr': { name: '뉴스웨이', type: '일간&경제지', influence: '3급' },
  'wikitree.co.kr': { name: '위키트리', type: '일간&경제지', influence: '3급' },
  'sentv.co.kr': { name: '서울경제TV', type: '일간&경제지', influence: '3급' },
  'lawissue.co.kr': { name: '로이슈', type: '일간&경제지', influence: '3급' },
  'siminilbo.co.kr': { name: '시민일보', type: '일간&경제지', influence: '3급' },
  'thefirstmedia.net': { name: '더퍼스트미디어', type: '일간&경제지', influence: '3급' },
  'n.news.naver.com': { name: '네이버뉴스', type: '일간&경제지', influence: '3급' },
  // 의학전문지
  'docdocdoc.co.kr': { name: '청년의사', type: '의학전문지', influence: '3급' },
  'dailymedi.com': { name: '데일리메디', type: '의학전문지', influence: '3급' },
  'medicaltimes.com': { name: '메디칼타임즈', type: '의학전문지', influence: '3급' },
  'medigatenews.com': { name: '메디게이트뉴스', type: '의학전문지', influence: '3급' },
  'doctorsnews.co.kr': { name: '닥터스뉴스', type: '의학전문지', influence: '3급' },
  'medicalworldnews.co.kr': { name: '메디컬월드뉴스', type: '의학전문지', influence: '3급' },
  'health.chosun.com': { name: '헬스조선', type: '의학전문지', influence: '2급' },
  'k-health.com': { name: '헬스경향', type: '의학전문지', influence: '3급' },
  'jhealthmedia.joins.com': { name: '헬스미디어', type: '의학전문지', influence: '3급' },
  'medicopharma.co.kr': { name: '메디코파마', type: '의학전문지', influence: '3급' },
  'bokuennews.com': { name: '보건뉴스', type: '의학전문지', influence: '3급' },
  // 사용자 요청(2026-10): 01_모니터링현황에 영문 도메인 그대로 남아있던 매체를 한글 매체명으로 매핑.
  // dt.co.kr은 사용자가 보낸 리스트엔 "디지털타임즈"로 적혀 있었으나 실제 매체명(STATS_MEDIA_WHITELIST에도
  // 등록된 표기)은 "디지털타임스"라 통일했다.
  'woman.donga.com': { name: '여성동아', type: '일간&경제지', influence: '2급' },
  'sedaily.com': { name: '서울경제', type: '일간&경제지', influence: '1급' },
  'digitalchosun.dizzo.com': { name: '디지털조선일보', type: '일간&경제지', influence: '2급' },
  'sports.donga.com': { name: '스포츠동아', type: '일간&경제지', influence: '2급' },
  'sportsseoul.com': { name: '스포츠서울', type: '일간&경제지', influence: '2급' },
  'news.tvchosun.com': { name: 'TV조선', type: '일간&경제지', influence: '1급' },
  'sportsworldi.com': { name: '스포츠월드', type: '일간&경제지', influence: '2급' },
  'news.sbs.co.kr': { name: 'SBS', type: '일간&경제지', influence: '1급' },
  'news.kbs.co.kr': { name: 'KBS', type: '일간&경제지', influence: '1급' },
  'biz.chosun.com': { name: '조선비즈', type: '일간&경제지', influence: '1급' },
  'mbn.co.kr': { name: 'MBN', type: '일간&경제지', influence: '1급' },
  'ytn.co.kr': { name: 'YTN', type: '일간&경제지', influence: '1급' },
  'ichannela.com': { name: '채널A', type: '일간&경제지', influence: '1급' },
  'segye.com': { name: '세계일보', type: '일간&경제지', influence: '1급' },
  'dt.co.kr': { name: '디지털타임스', type: '일간&경제지', influence: '1급' },
  'yonhapnewstv.co.kr': { name: '연합뉴스TV', type: '통신사', influence: '1급' }
};

/**
 * 사용자 요청(2026-10)으로 "보도통계 집계 매체" 목록(출처: 보도통계 집계 매체_2025-02-14, 7개 유형 51개 매체)을
 * 표준매체명 기준으로 등록한다. 08_Stats.gs의 주간·월간·연간·기간조회 통계는 이 목록에 있는 매체 보도만
 * 집계한다 — 수집·모니터링현황 자체는 그대로 전체 매체를 담되, "공식 보도실적" 집계만 이 매체로 한정한다.
 */
const STATS_MEDIA_WHITELIST = [
  // 종합일간지(12)
  '조선일보', '중앙일보', '동아일보', '경향신문', '국민일보', '문화일보', '서울신문', '세계일보',
  '한겨레', '한국일보', '내일신문', '아시아투데이',
  // 통신사(3)
  '뉴스1', '뉴시스', '연합뉴스',
  // 경제지(18)
  '매일경제', '머니투데이', '서울경제', '아주경제', '아시아경제', '이데일리', '파이낸셜뉴스', '한국경제',
  '헤럴드경제', '디지털타임스', '전자신문', '이투데이', '글로벌이코노믹', '이뉴스투데이', '파이낸셜투데이',
  '뉴스토마토', '프라임경제', '한스경제',
  // 스포츠지(6)
  '스포츠조선', '스포츠동아', '스포츠경향', '스포츠월드', '일간스포츠', '스포츠한국',
  // 섹션지(7)
  '헬스조선', '헬스중앙', '헬스경향', '매경헬스', '쿠키뉴스', '동아사이언스', '조선비즈',
  // 월간지(2)
  '신동아', '여성동아',
  // 영자지(3)
  '코리아중앙데일리', '코리아타임즈', '코리아헤럴드'
];

/**
 * 사용자 요청(2026-09)으로 모니터링 대상에서 완전히 제외한 매체(표준매체명 기준).
 * 이 목록에 있는 매체는 앞으로 수집되어도 01_모니터링현황에 아예 저장되지 않는다(processAndAppend_ 참고).
 * 더 제외할 매체가 생기면 이 배열에 표준매체명을 추가하면 된다(00_설정!매체마스터에 등록된 이름과 동일해야 함).
 */
const EXCLUDED_MEDIA = [
  '시민일보', '바이오스펙테이터', '문화경제', '더스탁', '베리타스알파', '신아일보', '캐치뉴스',
  '헬로디디', '더뷰어스', '더퍼스트미디어', '비즈니스플러스', '비욘드포스트', '한국금융신문', '시대',
  '교수신문', '더밸류뉴스', '비즈워크', '미디어펜', '국토일보', '더페어', '핀포인트뉴스',
  // 2026-10 사용자 요청: 아유경제(areyou.co.kr)는 아직 매체마스터에 정식 매핑이 없어 원문 도메인이
  // 그대로 표준매체명으로 쓰이므로, 혹시 모를 정식 매핑까지 대비해 두 값을 모두 등록해둔다.
  '아유경제', 'areyou.co.kr',
  // 2026-10 사용자 요청: 아래 매체는 모니터링 대상에서 제외(기존 기사 삭제 + 매체마스터에서도 제거).
  '코리아헬스로그', '뉴스핌', '약사공론', '의약뉴스', '톱스타뉴스', '가톨릭평화방송 뉴스', '매일신문',
  '뉴스웨이', '의사신문', '라포르시안', '인사이트', '아시아투데이', '브릿지경제', '메디컬월드뉴스',
  '이코노미사이언스', '메디칼트리뷴', '메디파나뉴스', '데일리메디', 'BusinessKorea', '중앙이코노미뉴스',
  '아시아타임즈', '마이데일리', '후생신보', '프레스나인', '메디소비자뉴스', '청년의사', '보건뉴스'
];

/**
 * 00_설정 시트 내 각 블록의 시작 셀(A1 표기). initializeSystem()이 이 좌표에 맞춰
 * 헤더를 생성하므로, 여기서 좌표를 바꾸면 01_Setup.gs의 레이아웃도 함께 바꿔야 한다.
 */
const CFG_BLOCK = {
  키워드:      { header: 'A2', dataStart: 'A3', cols: ['구분', '키워드', '사용여부'] },
  제외어:      { header: 'E2', dataStart: 'E3', cols: ['제외어', '사용여부'] },
  매체마스터:  { header: 'H2', dataStart: 'H3', cols: ['수집매체명', '표준매체명', '매체유형', '매체영향력'] },
  위기키워드:  { header: 'M2', dataStart: 'M3', cols: ['유형', '키워드', '가중치'] },
  수신자:      { header: 'R2', dataStart: 'R3', cols: ['이름', '소속', '연락처', '수신등급', '수신동의'] },
  RSS피드:     { header: 'X2', dataStart: 'X3', cols: ['피드명', 'URL', '사용여부'] },
  임계값:      { header: 'AB2', dataStart: 'AB3', cols: ['항목', '값', '설명'] },
  논조키워드:  { header: 'AF2', dataStart: 'AF3', cols: ['구분(긍정/부정)', '키워드'] },
  카테고리키워드: { header: 'AI2', dataStart: 'AI3', cols: ['카테고리', '키워드'] },
  진료과키워드: { header: 'AL2', dataStart: 'AL3', cols: ['진료과·센터', '키워드'] },
  오탐예외:    { header: 'AO2', dataStart: 'AO3', cols: ['키워드', '등록일', '예시제목'] },
  의료진사전:  { header: 'AS2', dataStart: 'AS3', cols: ['이름', '소속'] }
};

// 임계값 기본값 — 00_설정!임계값 블록에 최초 1회 기록되며, 이후 운영 중 담당자가 값만 수정한다.
const DEFAULT_THRESHOLDS = [
  ['수집주기_기본_분', 15, 'F-101: 평일 주간 기본 수집 주기(분)'],
  ['수집주기_야간주말_분', 45, 'F-105: 야간·주말 완화 주기(분)'],
  ['수집주기_위기_분', 5, 'F-105: 위기 모드 시 수집 주기(분)'],
  ['이슈그룹_유사도_임계', 0.62, 'F-202: 제목 유사도 임계값(0~1)'],
  ['확산가산_3건이상', 3, 'F-503/6.3: 1시간 내 3건 이상 시 가산점(사실상 급증 감지·등급 상향 역할)'],
  ['확산가산_5건이상', 6, '6.3: 1시간 내 5건 이상 가산점'],
  ['확산가산_커뮤니티전이', 3, '6.3: 커뮤니티→뉴스 전이 가산점'],
  ['점수_심각', 18, '6.3: 심각 등급 임계 점수'],
  ['점수_경계', 10, '6.3: 경계 등급 임계 점수'],
  ['점수_주의', 5, '6.3: 주의 등급 임계 점수'],
  ['점수_관심', 2, '6.3: 관심 등급 임계 점수'],
  ['알림_쿨다운_시간', 2, 'F-507: 동일 이슈그룹 동일 등급 재발송 억제(시간)'],
  ['야간_시작', '22:00', 'F-508: 야간 정책 시작 시각'],
  ['야간_종료', '07:00', 'F-508: 야간 정책 종료 시각'],
  ['API_한도_경고_비율', 0.8, 'F-107: 호출 한도 대비 경고 비율'],
  ['네이버_API_월간한도', 775000, 'F-107: NAVER API HUB 검색 API 통합 월 한도(2026-08 기준, 변동될 수 있으니 콘솔에서 재확인)'],
  ['수집실패_경고_연속횟수', 3, 'F-603: 연속 실패 경고 기준'],
  ['수집중단_경고_시간', 6, 'F-603: 신규 수집 없음 경고 기준(시간)'],
  ['커뮤니티수집_사용', 'N', 'F-103: Y로 바꾸면 카페·블로그 수집을 포함합니다'],
  ['위기모드_해제_시간', 12, '위기 모드 자동 해제 기준: 최근 N시간 내 경계/심각·미종결 건이 없으면 평시 주기로 복귀']
];

/** 스크립트 속성(민감정보)에서 값을 읽는다. 없으면 '' 반환. */
function getSecret_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

/** 00_설정 시트 객체를 반환 (없으면 예외). */
function getConfigSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  if (!sh) throw new Error('00_설정 시트가 없습니다. initializeSystem()을 먼저 실행하세요.');
  return sh;
}

/**
 * CFG_BLOCK 정의를 이용해 한 블록의 데이터 행을 [{col:val,...}, ...] 형태로 읽는다.
 * 첫 컬럼이 빈 값이면 그 지점에서 읽기를 멈춘다.
 */
function readConfigBlock_(blockName) {
  const block = CFG_BLOCK[blockName];
  if (!block) throw new Error('알 수 없는 설정 블록: ' + blockName);
  const sh = getConfigSheet_();
  const startRange = sh.getRange(block.dataStart);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();
  const numCols = block.cols.length;
  const maxRows = sh.getMaxRows() - startRow + 1;
  if (maxRows <= 0) return [];
  const values = sh.getRange(startRow, startCol, maxRows, numCols).getValues();
  const out = [];
  for (const row of values) {
    if (row[0] === '' || row[0] === null) break;
    const obj = {};
    block.cols.forEach((c, i) => obj[c] = row[i]);
    out.push(obj);
  }
  return out;
}

/** 지정 블록의 첫 빈 행에 rowValues(배열)를 추가한다. F-512 오탐 예외 축적 등에 사용. */
function appendToConfigBlock_(blockName, rowValues) {
  const block = CFG_BLOCK[blockName];
  const sh = getConfigSheet_();
  const startRange = sh.getRange(block.dataStart);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();
  const numCols = block.cols.length;
  const colValues = sh.getRange(startRow, startCol, sh.getMaxRows() - startRow + 1, 1).getValues();
  let targetRow = startRow;
  for (let i = 0; i < colValues.length; i++) {
    if (colValues[i][0] === '' || colValues[i][0] === null) { targetRow = startRow + i; break; }
    targetRow = startRow + i + 1;
  }
  sh.getRange(targetRow, startCol, 1, numCols).setValues([rowValues]);
}

/** 임계값 블록을 {항목: 값} 맵으로 변환해 반환 (자주 쓰므로 캐시). */
function getThresholds_() {
  const rows = readConfigBlock_('임계값');
  const map = {};
  rows.forEach(r => map[r['항목']] = r['값']);
  return map;
}
