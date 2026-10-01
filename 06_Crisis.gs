/**
 * 06_Crisis.gs — 위기 키워드 감지·등급 판정 (F-501~503, 명세서 6.1~6.3항)
 *
 * 위기점수 = (키워드 가중치 합계) × (매체 영향력 가중치) + (확산 가산점)
 */

const MEDIA_INFLUENCE_WEIGHT = { '1급': 3, '2급': 2, '3급': 1, '커뮤니티': 1 };

/**
 * 기사 1건의 위기 등급을 산정한다.
 * @return {grade, score, matchedKeywords, reasonText}
 */
function evaluateCrisis_(text, media, issueGroupId, sourceType, context, batchGroupTouch, crisisDict, orgPersonKeywords, th, fpExceptions) {
  const matched = matchCrisisKeywords_(text, crisisDict, orgPersonKeywords, fpExceptions || []);
  const keywordSum = matched.reduce((s, m) => s + m.weight, 0);

  if (keywordSum === 0) {
    return { grade: '해당없음', score: 0, matchedKeywords: [], reasonText: '' };
  }

  const mediaWeight = MEDIA_INFLUENCE_WEIGHT[media.매체영향력] || 1;
  const spread = computeSpreadBonus_(issueGroupId, sourceType, context, batchGroupTouch, th);
  const score = keywordSum * mediaWeight + spread.bonus;
  const grade = scoreToGrade_(score, th);

  const reasonText = matched.map(m => m.keyword + '(' + m.weight + '점)').join(', ') +
    (spread.bonus ? ' / 확산가산 +' + spread.bonus + '(' + spread.reason + ')' : '');

  return { grade: grade, score: score, matchedKeywords: matched, reasonText: reasonText };
}

/**
 * 위기 키워드 사전과 대조한다. 가중치 1점 키워드는 기관명/인물명이 본문에 함께 있을 때만 인정한다(6.1 운영원칙).
 */
function matchCrisisKeywords_(text, crisisDict, orgPersonKeywords, fpExceptions) {
  const hasOrgOrPerson = orgPersonKeywords.some(k => k && text.indexOf(k) >= 0);
  const exceptionSet = new Set(fpExceptions || []);
  const out = [];
  crisisDict.forEach(r => {
    const kw = r['키워드'];
    const weight = Number(r['가중치']) || 0;
    if (!kw || text.indexOf(kw) < 0) return;
    if (exceptionSet.has(kw)) return; // F-512: 오탐으로 축적된 키워드는 판정에서 제외
    if (weight <= 1 && !hasOrgOrPerson) return; // 단독 발화 금지
    out.push({ keyword: kw, type: r['유형'], weight: weight });
  });
  return out;
}

/**
 * 확산 가산점(6.3):
 *  - 동일 이슈그룹 1시간 내 3건 이상: +3 / 5건 이상: +6
 *  - 커뮤니티에서 뉴스로 전이: +3
 */
function computeSpreadBonus_(issueGroupId, sourceType, context, batchGroupTouch, th) {
  const oneHourAgo = new Date(Date.now() - 3600 * 1000);
  let countInHour = 0;
  let hadCommunity = false;
  context.forEach(c => {
    if (c.issueGroupId !== issueGroupId) return;
    const t = c.collectedAt instanceof Date ? c.collectedAt : new Date(c.collectedAt || 0);
    if (t >= oneHourAgo) countInHour++;
    if (c.sourceType === '커뮤니티') hadCommunity = true;
  });
  countInHour += (batchGroupTouch[issueGroupId] || 0);

  const th5 = Number(th['확산가산_5건이상']) || 6;
  const th3 = Number(th['확산가산_3건이상']) || 3;
  const thCommunity = Number(th['확산가산_커뮤니티전이']) || 3;

  let bonus = 0;
  const reasons = [];
  if (countInHour >= 5) { bonus += th5; reasons.push('1시간 내 ' + countInHour + '건'); }
  else if (countInHour >= 3) { bonus += th3; reasons.push('1시간 내 ' + countInHour + '건'); }

  if (hadCommunity && sourceType !== '커뮤니티') {
    bonus += thCommunity;
    reasons.push('커뮤니티→뉴스 전이');
  }
  return { bonus: bonus, reason: reasons.join(', '), countInHour: countInHour };
}

function scoreToGrade_(score, th) {
  const 심각 = Number(th['점수_심각']) || 18;
  const 경계 = Number(th['점수_경계']) || 10;
  const 주의 = Number(th['점수_주의']) || 5;
  const 관심 = Number(th['점수_관심']) || 2;
  if (score >= 심각) return '심각';
  if (score >= 경계) return '경계';
  if (score >= 주의) return '주의';
  if (score >= 관심) return '관심';
  return '해당없음';
}
