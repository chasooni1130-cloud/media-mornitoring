/**
 * 04_CollectRSS.gs — RSS 피드 수집 (F-102: 의학전문지 RSS, 구글 알리미 RSS)
 * 피드 목록은 00_설정!RSS피드 블록에서 관리한다(사용여부=Y인 것만 호출).
 */

function collectRSS_() {
  const feeds = readConfigBlock_('RSS피드')
    .filter(f => (f['사용여부'] === 'Y' || f['사용여부'] === true) && f['URL']);

  const results = [];
  feeds.forEach(feed => {
    try {
      const resp = UrlFetchApp.fetch(feed['URL'], { muteHttpExceptions: true, followRedirects: true });
      if (resp.getResponseCode() !== 200) {
        writeLog_('collectRSS:' + feed['피드명'], 1, 0, 'HTTP ' + resp.getResponseCode(), 0);
        return;
      }
      const items = parseRssItems_(resp.getContentText());
      items.forEach(item => {
        results.push({
          title: stripHtml_(item.title),
          link: item.link,
          description: stripHtml_(item.description),
          pubDate: item.pubDate || new Date(),
          rawMedia: feed['피드명'],
          sourceType: 'news',
          collectMethod: 'RSS',
          keyword: 'RSS:' + feed['피드명']
        });
      });
    } catch (e) {
      writeLog_('collectRSS:' + feed['피드명'], 1, 0, '예외: ' + e.message, 0);
    }
  });
  return results;
}

/**
 * RSS 2.0 / Atom을 모두 대략 지원하는 경량 파서.
 * XmlService 대신 정규식을 쓰는 이유: 구글 알리미 등 일부 피드가 완전한 XML 스펙을 지키지 않아
 * XmlService.parse()가 예외를 던지는 사례가 있어, 견고성을 우선했다.
 */
function parseRssItems_(xmlText) {
  const items = [];
  const isAtom = /<feed[\s>]/i.test(xmlText) && !/<rss[\s>]/i.test(xmlText);
  const itemTag = isAtom ? 'entry' : 'item';
  const blocks = xmlText.split(new RegExp('<' + itemTag + '[\\s>]', 'i')).slice(1);

  blocks.forEach(block => {
    const closeIdx = block.search(new RegExp('</' + itemTag + '>', 'i'));
    const body = closeIdx >= 0 ? block.substring(0, closeIdx) : block;

    const title = extractTag_(body, 'title');
    let link = extractTag_(body, 'link');
    if (isAtom) {
      const hrefMatch = body.match(/<link[^>]*href=["']([^"']+)["'][^>]*\/?>/i);
      if (hrefMatch) link = hrefMatch[1];
    }
    const description = extractTag_(body, 'description') || extractTag_(body, 'summary') || extractTag_(body, 'content');
    const pubDateRaw = extractTag_(body, 'pubDate') || extractTag_(body, 'published') || extractTag_(body, 'updated');
    const pubDate = pubDateRaw ? new Date(pubDateRaw) : new Date();

    if (title || link) {
      items.push({ title: title, link: link, description: description, pubDate: isNaN(pubDate) ? new Date() : pubDate });
    }
  });
  return items;
}

function extractTag_(xml, tag) {
  const m = xml.match(new RegExp('<' + tag + '[^>]*>([\\s\\S]*?)</' + tag + '>', 'i'));
  if (!m) return '';
  return m[1].replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').trim();
}
