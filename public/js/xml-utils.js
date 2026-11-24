// XML の読み取り(KML・GPX で共用)
// 名前空間の接頭辞が付いていても読めるよう、要素名は localName で比べる。

// XML を解析する。壊れていれば null
export function parseXml(text) {
  let doc;
  try {
    doc = new DOMParser().parseFromString(text, 'application/xml');
  } catch {
    return null;
  }
  if (!doc || !doc.documentElement || doc.getElementsByTagName('parsererror').length > 0) return null;
  return doc;
}

function childElements(el) {
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1);
}

// 直下の子要素から探す(子孫の同名要素と取り違えないように)
export function firstChild(el, localName) {
  return childElements(el).find((c) => c.localName === localName) || null;
}

export function firstDescendant(el, localName) {
  return el.getElementsByTagNameNS('*', localName)[0] || null;
}

export function allDescendants(el, localName) {
  return Array.from(el.getElementsByTagNameNS('*', localName));
}

export function textOf(el) {
  return el ? el.textContent.trim() : '';
}
