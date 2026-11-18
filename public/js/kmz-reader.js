// KMZ の展開・KML の解析
// KMZ は KML(地点の一覧)と画像を ZIP にまとめたもの。主な入力は minoh-hiking の写真出力で、
// 構成は次のとおり:
//   doc.kml             … 写真1枚につき Placemark 1つ(名前・撮影日時・位置・画像への参照)
//   files/photo-001.jpg … 写真本体(doc.kml の description から <img src> で参照する)
// 他のアプリの KMZ にも備え、「位置(Point)と、KMZ 内の画像への参照(description の <img src>)を
// 持つ Placemark」を写真として読めるだけ読む。ZIP の展開は JSZip(CDN)に任せる(無圧縮・Deflate とも読める)。

// 利用者に理由を伝えるためのエラー(メッセージをそのまま画面に出す)
export class KmzError extends Error {
  constructor(message) {
    super(message);
    this.name = 'KmzError';
  }
}

// 画像形式の判定。拡張子より中身(先頭のバイト列)を優先する
const IMAGE_FORMATS = [
  { ext: 'jpg', type: 'image/jpeg', test: (b) => b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF },
  { ext: 'png', type: 'image/png', test: (b) => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47 },
  { ext: 'gif', type: 'image/gif', test: (b) => b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38 },
  {
    ext: 'webp',
    type: 'image/webp',
    test: (b) => b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 &&
      b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50
  }
];
const EXT_TYPES = {
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  webp: 'image/webp'
};

function detectImageFormat(bytes, path) {
  const byContent = IMAGE_FORMATS.find((f) => f.test(bytes));
  if (byContent) return { ext: byContent.ext, type: byContent.type };
  const ext = (path.split('.').pop() || '').toLowerCase();
  if (EXT_TYPES[ext]) return { ext: ext === 'jpeg' ? 'jpg' : ext, type: EXT_TYPES[ext] };
  return null;
}

// ===== XML の読み取り(名前空間の接頭辞が付いていても読めるよう localName で比べる) =====
function childElements(el) {
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1);
}

function firstChild(el, localName) {
  return childElements(el).find((c) => c.localName === localName) || null;
}

function firstDescendant(el, localName) {
  return el.getElementsByTagNameNS('*', localName)[0] || null;
}

function textOf(el) {
  return el ? el.textContent.trim() : '';
}

// KML の置き場所(baseDir)からの相対パスを、ZIP 内のパスに直す(./ と ../ を解決する)
function resolvePath(baseDir, src) {
  const parts = [];
  for (const seg of (baseDir + src).split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

// KMZ ファイル(File)を読み込み、アルバム情報を返す。
//   { fileName, kmzBytes, title, photos: [{ name, timeMs, lat, lng, bytes, ext, type, srcPath }] }
// photos は出力順(撮影日時の昇順。撮影日時の無い写真は KML の順で末尾)に並べる。
// kmzBytes は KMZ そのもののバイト列(読み込んだ後に元のファイルを移動・削除しても保存できるよう持っておく)。
export async function readKmz(file) {
  const JSZip = globalThis.JSZip;
  const kmzBytes = new Uint8Array(await file.arrayBuffer());
  let zip;
  try {
    zip = await JSZip.loadAsync(kmzBytes);
  } catch {
    throw new KmzError('KMZ(ZIP 形式)として読み込めませんでした。ファイルが壊れているか、KMZ ではない可能性があります。');
  }

  const entries = Object.values(zip.files).filter((f) => !f.dir);
  const byName = new Map(entries.map((e) => [e.name, e]));
  const byLowerName = new Map(entries.map((e) => [e.name.toLowerCase(), e]));

  const kmlEntry = byName.get('doc.kml') || entries.find((e) => /\.kml$/i.test(e.name));
  if (!kmlEntry) {
    throw new KmzError('KMZ の中に KML ファイル(doc.kml)が見つかりません。写真の情報を読み取れませんでした。');
  }

  const kmlText = new TextDecoder('utf-8').decode(await kmlEntry.async('uint8array'));
  const doc = parseXml(kmlText);
  if (!doc) {
    throw new KmzError('KMZ の中の KML ファイルが壊れていて、内容を読み取れませんでした。');
  }

  // 画像の参照は KML の置き場所からの相対パス
  const baseDir = kmlEntry.name.includes('/') ? kmlEntry.name.slice(0, kmlEntry.name.lastIndexOf('/') + 1) : '';
  const findEntry = (path) => byName.get(path) || byLowerName.get(path.toLowerCase()) || null;

  const photos = [];
  const placemarks = Array.from(doc.getElementsByTagNameNS('*', 'Placemark'));
  for (const pm of placemarks) {
    // 位置: <Point><coordinates>経度,緯度[,標高]</coordinates></Point>
    const point = firstDescendant(pm, 'Point');
    const coordText = point ? textOf(firstDescendant(point, 'coordinates')) : '';
    if (!coordText) continue;
    const [lng, lat] = coordText.split(/\s+/)[0].split(',').map(Number);
    if (!Number.isFinite(lat) || !Number.isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;

    // 画像: <description> 内の <img src>。KMZ の外(http: や data: など)を指すものは対象外
    const desc = firstDescendant(pm, 'description')?.textContent || '';
    const src = /<img[^>]+src\s*=\s*["']([^"']+)["']/i.exec(desc)?.[1]?.trim();
    if (!src || /^[a-z][a-z0-9+.-]*:/i.test(src) || src.startsWith('//')) continue;
    const rawPath = resolvePath(baseDir, src.replace(/[?#].*$/, ''));
    let decodedPath = rawPath;
    try { decodedPath = decodeURI(rawPath); } catch { /* そのまま使う */ }
    const entry = findEntry(rawPath) || findEntry(decodedPath);
    if (!entry) continue;

    // 写真は KMZ の中のバイト列をそのまま使う(再圧縮しない)
    const bytes = await entry.async('uint8array');
    const format = detectImageFormat(bytes, entry.name);
    if (!format) continue;

    // 撮影日時: <TimeStamp><when>(UTC の ISO 8601)
    const timeStamp = firstDescendant(pm, 'TimeStamp');
    const whenText = timeStamp ? textOf(firstDescendant(timeStamp, 'when')) : '';
    const ms = whenText ? Date.parse(whenText) : NaN;

    photos.push({
      // 名称は Placemark 直下の <name>(Document の <name> と取り違えないよう直下の子から探す)
      name: textOf(firstChild(pm, 'name')),
      timeMs: Number.isFinite(ms) ? ms : null,
      lat,
      lng,
      bytes,
      ext: format.ext,
      type: format.type,
      srcPath: entry.name,
      kmlIndex: photos.length
    });
  }

  if (photos.length === 0) {
    throw new KmzError(
      'KMZ の中に写真が見つかりませんでした。位置と、KMZ 内の画像への参照を持つ地点(Placemark)がありません。'
    );
  }

  return {
    fileName: file.name,
    kmzBytes,
    title: documentTitle(doc),
    photos: sortPhotos(photos)
  };
}

// XML を解析する。壊れていれば null
function parseXml(text) {
  let doc;
  try {
    doc = new DOMParser().parseFromString(text, 'application/xml');
  } catch {
    return null;
  }
  if (!doc || !doc.documentElement || doc.getElementsByTagName('parsererror').length > 0) return null;
  return doc;
}

// タイトル: Document 直下の <name>(無ければ最初の Folder 直下の <name>)
function documentTitle(doc) {
  for (const tag of ['Document', 'Folder']) {
    const el = doc.getElementsByTagNameNS('*', tag)[0];
    const name = el ? textOf(firstChild(el, 'name')) : '';
    if (name) return name;
  }
  return '';
}

// 撮影日時の昇順。撮影日時の無い写真は KML の順で末尾に置く
function sortPhotos(photos) {
  return [...photos].sort((a, b) => {
    const ta = a.timeMs;
    const tb = b.timeMs;
    if (ta !== null && tb !== null && ta !== tb) return ta - tb;
    if (ta === null && tb !== null) return 1;
    if (ta !== null && tb === null) return -1;
    return a.kmlIndex - b.kmlIndex;
  });
}
