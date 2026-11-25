// 読み込んだ KMZ・GPX を、それぞれ1つのファイルにまとめる(「保存する」で、KMZ・GPX が2つ以上のとき)。
// KMZ・GPX が1つなら、まとめずに元のファイルをそのまま保存する(app.js)。
// 元のファイルの情報をできるだけ落とさないよう、読み取った値から作り直すのではなく、
// 写真の地点(<Placemark>)・経路(<trkseg>)・地点(<wpt>)の要素を、元のファイルからそのまま書き写す。
// 入れるのは、画面と同じく重なりを除いた写真・経路(merge.js)。
// 構成は minoh-hiking の出力と同じにし、minoh-hiking や Trail Album で読み込み直せるようにする。
import { firstDescendant } from './xml-utils.js';
import { pad3 } from './naming.js';
import { APP_NAME } from './constants.js';

const KML_NS = 'http://www.opengis.net/kml/2.2';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8"?>\n';
// description の中の画像の参照(kmz-reader.js と同じ見つけ方。最初の <img src> だけ)
const IMG_SRC = /(<img[^>]+src\s*=\s*["'])([^"']+)(["'])/i;

// 写真の KMZ をまとめ、KMZ のバイト列(Uint8Array)を返す。
// photos は naming.js の planPhotos の結果(保存の順。no・ext・bytes・placemark を持つ)、title は KML の Document 名。
// 構成: doc.kml(写真1枚につき Placemark 1つ)と files/photo-001.jpg …(写真は元のバイト列のまま)
export async function buildMergedKmz(photos, title) {
  const doc = document.implementation.createDocument(KML_NS, 'kml', null);
  const documentEl = doc.createElementNS(KML_NS, 'Document');
  const nameEl = doc.createElementNS(KML_NS, 'name');
  nameEl.textContent = title;
  appendLine(documentEl, nameEl, '\n  ');

  const images = photos.map((photo) => {
    // 写真の名前は KMZ ごとに付いていたもので重なることがあるため、保存の番号で付け直し、参照も書き換える
    const path = `files/photo-${pad3(photo.no)}.${photo.ext}`;
    const placemark = doc.importNode(photo.placemark, true);
    pointImageTo(placemark, path);
    appendLine(documentEl, placemark, '\n  ');
    return { path, data: photo.bytes };
  });
  documentEl.append('\n');
  appendLine(doc.documentElement, documentEl, '\n');
  doc.documentElement.append('\n');
  declarePrefixes(doc.documentElement, photos.map((p) => p.placemark.ownerDocument.documentElement));

  const zip = new globalThis.JSZip();
  // doc.kml を先頭に置く(KMZ の慣例。最初の .kml を本体とみなすアプリがある)。
  // 写真はもともと圧縮されているので無圧縮、doc.kml だけ圧縮する。フォルダ(files/)の項目は作らない
  zip.file('doc.kml', XML_HEADER + new XMLSerializer().serializeToString(doc), { compression: 'DEFLATE' });
  for (const image of images) zip.file(image.path, image.data, { binary: true, createFolders: false });
  return zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
}

// Placemark の description の中の画像の参照を path に書き換える(CDATA にして書き戻す)
function pointImageTo(placemark, path) {
  const desc = firstDescendant(placemark, 'description');
  if (!desc) return;
  const text = desc.textContent.replace(IMG_SRC, (match, before, src, after) => `${before}${path}${after}`);
  const doc = placemark.ownerDocument;
  desc.replaceChildren(text.includes(']]>') ? doc.createTextNode(text) : doc.createCDATASection(text));
}

// 移動経路の GPX をまとめ、GPX のバイト列(Uint8Array)を返す。
// track は merge.js の buildTrack の結果(segments は、すべての GPX を通した記録の日時の順)、name はトラックの名前。
// 構成: トラック(<trk>)1本に、経路ごとの <trkseg>。地点(<wpt>)があれば、同じものを除いてトラックの前に置く。
// GPX の版(名前空間・version)は、いちばん早く記録した GPX に合わせる
export function buildMergedGpx(track, name) {
  const first = track.files[0].source;
  const ns = first.namespace;
  const doc = document.implementation.createDocument(ns, 'gpx', null);
  const root = doc.documentElement;
  root.setAttribute('version', first.version || '1.1');
  root.setAttribute('creator', APP_NAME);

  const serializer = new XMLSerializer();
  const seen = new Set();
  for (const file of track.files) {
    for (const wpt of file.source.waypoints) {
      const key = serializer.serializeToString(wpt);
      if (seen.has(key)) continue;
      seen.add(key);
      appendLine(root, doc.importNode(wpt, true), '\n  ');
    }
  }

  const trk = doc.createElementNS(ns, 'trk');
  const nameEl = doc.createElementNS(ns, 'name');
  nameEl.textContent = name;
  appendLine(trk, nameEl, '\n    ');
  for (const segment of track.segments) {
    let trkseg;
    if (segment.xmlNode) {
      trkseg = doc.importNode(segment.xmlNode, true);
    } else {
      // trkseg を持たない GPX: <trkpt> の並びを trkseg 1本に入れる
      trkseg = doc.createElementNS(ns, 'trkseg');
      for (const node of segment.xmlPointNodes) trkseg.append(doc.importNode(node, true));
    }
    appendLine(trk, trkseg, '\n    ');
  }
  trk.append('\n  ');
  appendLine(root, trk, '\n  ');
  root.append('\n');
  declarePrefixes(root, track.segments.map((s) => (s.xmlNode ?? s.xmlPointNodes[0]).ownerDocument.documentElement));
  return new TextEncoder().encode(XML_HEADER + serializer.serializeToString(doc) + '\n');
}

// 元のファイルの直下の要素で宣言されていた名前空間の接頭辞(GPX の拡張の gpxtpx:、KML の gx: など)を、
// まとめたファイルの直下でも宣言する(宣言しないと、書き写した要素ごとに宣言が付いて読みにくくなる)。
// 同じ接頭辞が別の名前空間を指していたら、先のものを使う(後のものは、書き写した要素で宣言される)
function declarePrefixes(root, sourceRoots) {
  for (const source of new Set(sourceRoots)) {
    for (const attr of Array.from(source.attributes)) {
      if (attr.prefix !== 'xmlns' || root.hasAttribute(attr.name)) continue;
      root.setAttributeNS(XMLNS_NS, attr.name, attr.value);
    }
  }
}

// 要素を、改行と字下げ(indent)の後に加える(まとめたファイルを、テキストとして開いても読みやすいように)
function appendLine(parent, child, indent) {
  parent.append(indent, child);
}
