// GPX の解析(移動経路の読み込み)
// GPX は移動経路(トラック)を記録した XML。主な入力は minoh-hiking の移動経路の出力で、構成は次のとおり:
//   <gpx><trk><name>ファイル名(拡張子なし)</name>
//     <trkseg><trkpt lat="…" lon="…"><time>UTC の ISO 8601</time></trkpt>…</trkseg>  … 記録1回ごとに1本
//   </trk></gpx>
// minoh-hiking の読み込みと同じく、trkseg 1本を1本の経路として扱う。trkseg を持たない GPX への保険として、
// その場合は文書順の trkpt 全体を1本の経路にする。<time> は任意。
import { parseXml, firstChild, allDescendants, textOf } from './xml-utils.js';

// 利用者に理由を伝えるためのエラー(メッセージをそのまま画面に出す)
export class GpxError extends Error {
  constructor(message) {
    super(message);
    this.name = 'GpxError';
  }
}

// 地球の半径[m]。minoh-hiking(Leaflet の distanceTo)と同じ値・同じ式にして、移動距離をそろえる
const EARTH_RADIUS_M = 6371000;

// 2地点間の距離[m](球面上の大円距離)
function distanceBetween(a, b) {
  const rad = Math.PI / 180;
  const sinDLat = Math.sin(((b.lat - a.lat) * rad) / 2);
  const sinDLng = Math.sin(((b.lng - a.lng) * rad) / 2);
  const h = sinDLat * sinDLat + Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * sinDLng * sinDLng;
  return EARTH_RADIUS_M * 2 * Math.atan2(Math.sqrt(h), Math.sqrt(1 - h));
}

// GPX ファイル(File)を読み込み、移動経路の情報を返す。
//   { fileName, title, segments: [{ points: [{ lat, lng, timeMs }], distanceM, startMs, endMs }] }
// segments は GPX の順。startMs・endMs は記録時刻の最初と最後(記録時刻が1つも無ければ null)。
// 複数の GPX をまとめた合計(記録地点数・移動距離など)は merge.js の buildTrack で求める。
export async function readGpx(file) {
  const doc = parseXml(await file.text());
  if (!doc || doc.documentElement.localName !== 'gpx') {
    throw new GpxError('GPX として読み込めませんでした。ファイルが壊れているか、GPX ではない可能性があります。');
  }

  let pointLists = allDescendants(doc, 'trkseg')
    .map((seg) => readPoints(allDescendants(seg, 'trkpt')))
    .filter((points) => points.length > 0);
  if (pointLists.length === 0) {
    const all = readPoints(allDescendants(doc, 'trkpt'));
    if (all.length > 0) pointLists = [all];
  }
  if (pointLists.length === 0) {
    throw new GpxError('GPX の中に移動経路の記録地点(trkpt)が見つかりませんでした。');
  }

  return {
    fileName: file.name,
    title: gpxTitle(doc),
    segments: pointLists.map(summarize)
  };
}

// <trkpt lat lon> の並びから記録地点を取り出す(位置の値が正しくない地点は除く)
function readPoints(nodes) {
  const points = [];
  for (const node of nodes) {
    const lat = parseFloat(node.getAttribute('lat'));
    const lng = parseFloat(node.getAttribute('lon'));
    if (!Number.isFinite(lat) || !Number.isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;
    const timeText = textOf(firstChild(node, 'time'));
    const ms = timeText ? Date.parse(timeText) : NaN;
    points.push({ lat, lng, timeMs: Number.isFinite(ms) ? ms : null });
  }
  return points;
}

// 1本の経路の移動距離[m]と、記録時刻の最初と最後
function summarize(points) {
  let distanceM = 0;
  for (let i = 1; i < points.length; i++) distanceM += distanceBetween(points[i - 1], points[i]);
  return { points, distanceM, ...timeRange(points.map((p) => p.timeMs)) };
}

// 時刻の最初と最後(null は除く)。記録地点が多くても使えるよう、Math.min(...配列) を使わずに数える
function timeRange(times) {
  let startMs = null;
  let endMs = null;
  for (const ms of times) {
    if (ms === null) continue;
    if (startMs === null || ms < startMs) startMs = ms;
    if (endMs === null || ms > endMs) endMs = ms;
  }
  return { startMs, endMs };
}

// 名前: 最初の trk 直下の <name>(無ければ metadata 直下の <name>)
function gpxTitle(doc) {
  for (const tag of ['trk', 'metadata']) {
    const el = allDescendants(doc, tag)[0];
    const name = el ? textOf(firstChild(el, 'name')) : '';
    if (name) return name;
  }
  return '';
}
