/**
 * 03_CollectNaver.gs — 네이버 검색 API 기반 수집 (F-101 뉴스, F-103 커뮤니티·SNS)
 *
 * 2026-06-25 NAVER API HUB가 출시되면서 신규 발급은 기존 개발자센터(openapi.naver.com)가 아니라
 * 네이버클라우드플랫폼 콘솔에서 이뤄진다. 도메인·요청경로·인증헤더가 모두 바뀌었으므로
 * 기본값은 HUB 방식으로 두고, 레거시 키를 아직 쓰는 경우를 위해 전환 스위치를 남겨둔다.
 *
 * 필요한 스크립트 속성(프로젝트 설정 → 스크립트 속성):
 *   [HUB 방식 — 기본, 신규 발급은 모두 이 방식]
 *     NAVER_HUB_KEY_ID      → X-NCP-APIGW-API-KEY-ID
 *     NAVER_HUB_KEY         → X-NCP-APIGW-API-KEY
 *   [LEGACY 방식 — 2027-06-30까지만 지원되는 기존 개발자센터 키를 쓸 경우]
 *     NAVER_API_MODE = LEGACY
 *     NAVER_CLIENT_ID / NAVER_CLIENT_SECRET
 *   [공통, 선택] NAVER_API_BASE — 도메인이 안내와 다르면 이 값으로 덮어쓸 수 있다.
 *
 * 주의: HUB 쪽 API는 2026-06-25 출시된 지 얼마 되지 않았다. 아래 엔드포인트·응답 필드는
 * 공개된 이관 가이드를 기준으로 작성했으나, 실제 계정으로 최초 1회 테스트 호출을 해서
 * items[].title/link/description/pubDate 필드명이 그대로인지 반드시 확인하고 쓰기 바란다.
 */

const NAVER_ENDPOINTS = {
  hub: {
    base: 'https://naverapihub.apigw.ntruss.com',
    news: '/search/v1/news',
    cafearticle: '/search/v1/cafearticle',
    blog: '/search/v1/blog'
  },
  legacy: {
    base: 'https://openapi.naver.com',
    news: '/v1/search/news.json',
    cafearticle: '/v1/search/cafearticle.json',
    blog: '/v1/search/blog.json'
  }
};

/**
 * 네이버 뉴스 검색 (F-101). 00_설정!키워드 블록의 '기관명'/'상위기관'/'센터·사업'/'인물' 구분 키워드로 호출한다.
 * 반환: 정규화 이전의 raw article 배열
 */
function collectNaverNews_() {
  return collectNaverByEndpoint_('news', ['기관명', '상위기관', '인물', '센터·사업']);
}

/** 네이버 카페·블로그 검색 (F-103). 조기 경보 목적이므로 기관명 키워드만 사용. */
function collectNaverCommunity_() {
  const cafe = collectNaverByEndpoint_('cafearticle', ['기관명'], '커뮤니티');
  const blog = collectNaverByEndpoint_('blog', ['기관명'], '커뮤니티');
  return cafe.concat(blog);
}

function isNaverHubMode_() {
  return (getSecret_('NAVER_API_MODE') || 'HUB').toUpperCase() !== 'LEGACY';
}

function getNaverAuth_() {
  if (isNaverHubMode_()) {
    const keyId = getSecret_('NAVER_HUB_KEY_ID');
    const key = getSecret_('NAVER_HUB_KEY');
    if (!keyId || !key) return null;
    return {
      base: getSecret_('NAVER_API_BASE') || NAVER_ENDPOINTS.hub.base,
      paths: NAVER_ENDPOINTS.hub,
      headers: { 'X-NCP-APIGW-API-KEY-ID': keyId, 'X-NCP-APIGW-API-KEY': key },
      missingMsg: 'NAVER_HUB_KEY_ID/NAVER_HUB_KEY 미설정 — 스크립트 속성을 확인하세요.'
    };
  }
  const clientId = getSecret_('NAVER_CLIENT_ID');
  const clientSecret = getSecret_('NAVER_CLIENT_SECRET');
  if (!clientId || !clientSecret) return null;
  return {
    base: getSecret_('NAVER_API_BASE') || NAVER_ENDPOINTS.legacy.base,
    paths: NAVER_ENDPOINTS.legacy,
    headers: { 'X-Naver-Client-Id': clientId, 'X-Naver-Client-Secret': clientSecret },
    missingMsg: 'NAVER_CLIENT_ID/NAVER_CLIENT_SECRET 미설정 — 스크립트 속성을 확인하세요.'
  };
}

function collectNaverByEndpoint_(endpointKey, keywordGroups, forceSourceType) {
  const auth = getNaverAuth_();
  if (!auth) {
    writeLog_('collectNaver:' + endpointKey, 0, 0,
      (isNaverHubMode_() ? 'NAVER_HUB_KEY_ID/NAVER_HUB_KEY' : 'NAVER_CLIENT_ID/NAVER_CLIENT_SECRET') + ' 미설정 — 스크립트 속성을 확인하세요.', 0);
    return [];
  }
  const path = auth.paths[endpointKey];

  const keywords = readConfigBlock_('키워드')
    .filter(k => (k['사용여부'] === 'Y' || k['사용여부'] === true) && keywordGroups.indexOf(k['구분']) >= 0)
    .map(k => k['키워드'])
    .filter(String);

  const results = [];
  let callCount = 0;
  keywords.forEach(kw => {
    const url = auth.base + path + '?query=' + encodeURIComponent(kw) + '&display=100&sort=date';
    callCount++;
    incrementDailyCounter_('NAVER_API_CALLS');   // 대시보드/로그용 일별 카운트
    incrementMonthlyCounter_('NAVER_API_CALLS'); // F-107: HUB는 월 단위 통합 한도(기본 775,000건)로 관리됨
    try {
      const resp = UrlFetchApp.fetch(url, { method: 'get', headers: auth.headers, muteHttpExceptions: true });
      if (resp.getResponseCode() === 429) {
        writeLog_('collectNaver:' + endpointKey, callCount, 0, 'HTTP 429 — 호출 한도 초과. 잠시 후 자동 완화됩니다.', 0);
        return;
      }
      if (resp.getResponseCode() !== 200) {
        writeLog_('collectNaver:' + endpointKey, callCount, 0,
          '키워드 "' + kw + '" 호출 실패 HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 200), 0);
        return;
      }
      const json = JSON.parse(resp.getContentText());
      (json.items || []).forEach(item => {
        const title = stripHtml_(item.title);
        results.push({
          title: title,
          link: item.link || item.originallink,
          description: stripHtml_(item.description),
          pubDate: item.pubDate ? new Date(item.pubDate) : new Date(),
          rawMedia: guessMediaFromLink_(item.originallink || item.link),
          sourceType: forceSourceType || 'news',
          collectMethod: forceSourceType ? '커뮤니티' : 'API',
          keyword: kw
        });
      });
    } catch (e) {
      writeLog_('collectNaver:' + endpointKey, callCount, 0, '키워드 "' + kw + '" 예외: ' + e.message, 0);
    }
  });
  return results;
}

/**
 * 특정 검색어로 네이버 뉴스 검색을 1회만 호출해 원시 결과를 반환 (재검색·매체 백필용).
 * collectNaverByEndpoint_와 별개로 두는 이유: 키워드 목록 순회가 아니라 단건 재조회이기 때문.
 */
function fetchNaverNewsOnce_(query, display) {
  const auth = getNaverAuth_();
  if (!auth) return { error: 'NO_AUTH', items: [] };
  const url = auth.base + auth.paths.news + '?query=' + encodeURIComponent(query) + '&display=' + (display || 10) + '&sort=date';
  incrementDailyCounter_('NAVER_API_CALLS');
  incrementMonthlyCounter_('NAVER_API_CALLS');
  try {
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: auth.headers, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { error: 'HTTP_' + resp.getResponseCode(), items: [] };
    const json = JSON.parse(resp.getContentText());
    return { error: '', items: json.items || [] };
  } catch (e) {
    return { error: e.message, items: [] };
  }
}

/**
 * 네이버 뉴스 링크의 도메인에서 매체명을 대략 추정 (F-203 정규화 전 1차 힌트).
 * n.news.naver.com 미러 링크는 원 매체를 알려주지 않으므로, article/<코드>/의 코드를
 * NAVER_PRESS_CODE_MAP에서 찾아 실제 언론사 도메인으로 되돌린다(등록 안 된 코드는 그대로 네이버 도메인 반환).
 */
function guessMediaFromLink_(link) {
  if (!link) return '';
  const m = String(link).match(/https?:\/\/(?:www\.|m\.)?([^\/]+)/);
  const host = m ? m[1] : '';
  if (host === 'n.news.naver.com' || host === 'news.naver.com') {
    const codeMatch = String(link).match(/\/article\/(\d+)\//);
    if (codeMatch && NAVER_PRESS_CODE_MAP[codeMatch[1]]) return NAVER_PRESS_CODE_MAP[codeMatch[1]];
  }
  return host;
}
