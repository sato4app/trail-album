// 地理院地図に、写真の撮影場所と、移動経路の記録地点をすべてマーカーで表示し、
// 記録地点は経路ごとに記録の順に線で結ぶ(Leaflet を CDN から読み込んで使う)。
// 記録地点は数千になることもあるため、画像のマーカーではなく、canvas に描く円のマーカー(circleMarker)にする。
// 地図の中心は、KMZ の写真の平均の地点と、GPX の経路の線の真ん中(歩いた距離で重みを付けた平均)の、真ん中にする。

const GSI_TILE_URL = 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png';
const GSI_ATTRIBUTION =
  '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>';
const INITIAL_CENTER = [34.853667, 135.472041]; // 箕面大滝
const INITIAL_ZOOM = 14;
const FIT_MAX_ZOOM = 17; // 地点が1つ・近いときに寄りすぎないように
const FIT_PADDING = 24;  // いちばん端の地点と、地図の縁との間(px)

// マーカーの色。写真は1色、経路は GPX ごとに色を変える(表の A, B … の印と同じ色)
export const PHOTO_COLOR = '#e8590c';
const TRACK_COLORS = ['#1c7ed6', '#7048e8', '#0c8599', '#d6336c', '#5c940d', '#495057'];
export const trackColor = (index) => TRACK_COLORS[index % TRACK_COLORS.length];

let map = null;
let markers = null;
let fitTarget = null; // 中心(KMZ と GPX を同じ重みにして決めた地点)と、地図に入れる範囲の角(北西・南東)。地点が無ければ null
let autoFit = false;  // 利用者が地図を動かしていなければ true(地図の大きさが変わったら、合わせ直す)
let fitting = false;  // fitToTarget で地図を動かしている間だけ true

// 地図を使えるか(Leaflet を CDN から読み込めたか)
export const canShowMap = () => !!globalThis.L;

// 地図のマーカーを置き直し、KMZ と GPX を同じ重みにして決めた地点を中心に、すべてのマーカーが入るように地図を合わせる。
// fit が false なら、地図の中心・縮尺は今のままにする(KMZ・GPX をクリアしたとき)。
// container は表示されていること(隠れていると地図の大きさを測れない)。
//   photos:  [{ lat, lng, label, name }](写真の撮影場所)
//   tracks:  [{ label, colorIndex, points: [{ lat, lng, timeLabel }] }](経路ごとの記録地点)
export function updateMap(container, { photos, tracks }, { fit = true } = {}) {
  const L = globalThis.L;
  if (!L) return;
  if (!map) {
    map = L.map(container, { preferCanvas: true }).setView(INITIAL_CENTER, INITIAL_ZOOM);
    L.tileLayer(GSI_TILE_URL, { attribution: GSI_ATTRIBUTION, minZoom: 5, maxZoom: 18 }).addTo(map);
    markers = L.featureGroup().addTo(map);
    // 利用者が地図を動かした(ドラッグ・拡大・縮小)後は、大きさが変わっても合わせ直さない
    map.on('movestart', () => {
      if (!fitting) autoFit = false;
    });
    // 地図の枠の大きさが変わったら(ウインドウの幅を変えた・KMZ と GPX の境界線を動かした)描き直し、
    // 利用者が動かしていなければ中心と縮尺を合わせ直す。動かした後は、中心を保つだけにする
    // (地図の画像の読み込みは、動かし終えてからまとめて行う)
    new ResizeObserver(() => {
      map.invalidateSize({ debounceMoveend: true });
      if (autoFit) fitToTarget();
    }).observe(container);
  }
  map.invalidateSize(); // 隠れていた間に大きさが変わっていても合うように
  markers.clearLayers();

  // 経路の線と記録地点(写真の下に描く。先に加えたものほど下になる)。
  // 線は経路(A-1, A-2 …)ごとに引き、別の経路どうしはつながない
  for (const track of tracks) {
    const color = trackColor(track.colorIndex);
    L.polyline(track.points.map((p) => [p.lat, p.lng]), { color, weight: 3, opacity: 0.6 })
      .bindTooltip(track.label, { sticky: true })
      .addTo(markers);
  }
  for (const track of tracks) {
    const color = trackColor(track.colorIndex);
    track.points.forEach((p, i) => {
      L.circleMarker([p.lat, p.lng], { radius: 3, weight: 1, color, fillColor: color, fillOpacity: 0.85 })
        .bindTooltip(`${track.label} の ${i + 1} 点目${p.timeLabel ? `  ${p.timeLabel}` : ''}`)
        .addTo(markers);
    });
  }
  // 写真の撮影場所
  for (const photo of photos) {
    L.circleMarker([photo.lat, photo.lng], { radius: 7, weight: 2, color: '#fff', fillColor: PHOTO_COLOR, fillOpacity: 1 })
      .bindTooltip(`${photo.label}  ${photo.name}`)
      .addTo(markers);
  }

  if (!fit) {
    // 中心・縮尺は今のまま。この後に地図の大きさが変わっても、合わせ直さずに中心を保つだけにする
    autoFit = false;
    return;
  }

  // 中心は、KMZ(写真)全体と GPX(経路)全体を同じ重みにして決める:
  // KMZ の代表の地点(写真の撮影場所の平均)と、GPX の代表の地点(経路の線の真ん中: trackCenter)の真ん中。
  // 片方しか無ければ、その代表の地点
  const trackPoints = tracks.flatMap((t) => t.points);
  const centers = [];
  if (photos.length > 0) centers.push(averagePoint(photos));
  if (trackPoints.length > 0) centers.push(trackCenter(tracks));
  if (centers.length === 0) {
    fitTarget = null;
    return;
  }
  const bounds = L.latLngBounds([...photos, ...trackPoints].map((p) => [p.lat, p.lng]));
  fitTarget = {
    center: L.latLng(averagePoint(centers)),
    corners: [bounds.getNorthWest(), bounds.getSouthEast()]
  };
  autoFit = true;
  fitToTarget();
}

// GPX の代表の地点: 経路の線の真ん中(歩いた距離で重みを付けた平均)。
// 記録地点は一定の時間ごとに記録されるため、点の数で平均すると、休憩した場所やゆっくり歩いた区間に寄ってしまう。
// そこで、隣り合う2点の間ごとに、その中点を間の長さで重みづけて平均する(別の経路どうしの間は数えない)。
// 長さは重みに使うだけなので、緯度による東西の縮みだけを補った簡単な計算にする(地球の丸みは無視してよい)。
// 線の長さが 0(どの経路も1点だけ・同じ場所だけ)なら、記録地点の平均にする
function trackCenter(tracks) {
  let sumLat = 0;
  let sumLng = 0;
  let sumLength = 0;
  for (const { points } of tracks) {
    for (let i = 1; i < points.length; i++) {
      const a = points[i - 1];
      const b = points[i];
      const midLat = (a.lat + b.lat) / 2;
      const length = Math.hypot(b.lat - a.lat, (b.lng - a.lng) * Math.cos((midLat * Math.PI) / 180));
      sumLat += midLat * length;
      sumLng += ((a.lng + b.lng) / 2) * length;
      sumLength += length;
    }
  }
  if (sumLength === 0) return averagePoint(tracks.flatMap((t) => t.points));
  return { lat: sumLat / sumLength, lng: sumLng / sumLength };
}

// 地点の緯度・経度の平均({ lat, lng })
function averagePoint(points) {
  return {
    lat: points.reduce((sum, p) => sum + p.lat, 0) / points.length,
    lng: points.reduce((sum, p) => sum + p.lng, 0) / points.length
  };
}

// 地図の中心を fitTarget の中心にし、その中心のまま、
// すべての地点が入るいちばん大きい縮尺にする(FIT_MAX_ZOOM まで)
function fitToTarget() {
  const size = map.getSize();
  if (!fitTarget || size.x === 0 || size.y === 0) return; // 隠れている間は測れない
  const { center, corners } = fitTarget;

  // いちばん端の地点(北西の角・南東の角)が、中心から地図の縁までに収まる縮尺を、寄った方から探す
  const halfX = size.x / 2 - FIT_PADDING;
  const halfY = size.y / 2 - FIT_PADDING;
  let zoom = FIT_MAX_ZOOM;
  for (; zoom > map.getMinZoom(); zoom--) {
    const c = map.project(center, zoom);
    const fits = corners.every((corner) => {
      const p = map.project(corner, zoom);
      return Math.abs(p.x - c.x) <= halfX && Math.abs(p.y - c.y) <= halfY;
    });
    if (fits) break;
  }
  // 縮尺が同じで、中心もずれていなければ動かさない(動かすと地図の画像をすべて読み直し、
  // 境界線を動かしている間などに地図がちらつく)
  const shift = map.latLngToContainerPoint(center).subtract(size.divideBy(2));
  if (zoom === map.getZoom() && Math.abs(shift.x) <= 1 && Math.abs(shift.y) <= 1) return;
  // すぐに動かす(動きを見せると、その間に利用者が動かしたのかを見分けられない)
  fitting = true;
  map.setView(center, zoom, { animate: false });
  fitting = false;
}
