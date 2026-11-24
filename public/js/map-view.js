// 地理院地図に、写真の撮影場所と、移動経路の記録地点をすべてマーカーで表示し、
// 記録地点は経路ごとに記録の順に線で結ぶ(Leaflet を CDN から読み込んで使う)。
// 記録地点は数千になることもあるため、画像のマーカーではなく、canvas に描く円のマーカー(circleMarker)にする。
// 地図は、読み込んだ地点すべての東西南北の端が収まるように合わせる(地点の数や偏りに左右されない)。
// 右下に、縮尺の目盛りとズームレベルを出す。

const GSI_TILE_URL = 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png';
const GSI_ATTRIBUTION =
  '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>';
const INITIAL_CENTER = [34.853667, 135.472041]; // 箕面大滝
const INITIAL_ZOOM = 14;
const FIT_MAX_ZOOM = 17; // 地点が1つ・近いときに寄りすぎないように
const FIT_PADDING = 24;  // いちばん端の地点と、地図の縁との間(px。マーカーの円が縁で切れないように)
// ズームレベルの刻み。地点の範囲が地図いっぱいに収まるよう、整数より細かくする
// (整数のままだと、範囲が地図の半分ほどにしかならないことがある)
const ZOOM_SNAP = 0.1;

// マーカーの色。写真は1色、経路は GPX ごとに色を変える(表の A, B … の印と同じ色)
export const PHOTO_COLOR = '#e8590c';
const TRACK_COLORS = ['#1c7ed6', '#7048e8', '#0c8599', '#d6336c', '#5c940d', '#495057'];
export const trackColor = (index) => TRACK_COLORS[index % TRACK_COLORS.length];

let map = null;
let markers = null;
let fitArea = null;  // 地図に収める範囲(読み込んだ地点すべての東西南北の端。LatLngBounds)。地点が無ければ null
let autoFit = false; // 利用者が地図を動かしていなければ true(地図の大きさが変わったら、合わせ直す)
let fitting = false; // fitToArea で地図を動かしている間だけ true

// 地図を使えるか(Leaflet を CDN から読み込めたか)
export const canShowMap = () => !!globalThis.L;

// 地図のマーカーを置き直し、すべてのマーカーが収まるように地図を合わせる(fitToArea)。
// fit が false なら、地図の中心・縮尺は今のままにする(KMZ・GPX をクリアしたとき)。
// container は表示されていること(隠れていると地図の大きさを測れない)。
//   photos:  [{ lat, lng, label, name }](写真の撮影場所)
//   tracks:  [{ label, colorIndex, points: [{ lat, lng, timeLabel }] }](経路ごとの記録地点)
export function updateMap(container, { photos, tracks }, { fit = true } = {}) {
  const L = globalThis.L;
  if (!L) return;
  if (!map) {
    map = L.map(container, { preferCanvas: true, zoomSnap: ZOOM_SNAP }).setView(INITIAL_CENTER, INITIAL_ZOOM);
    L.tileLayer(GSI_TILE_URL, { attribution: GSI_ATTRIBUTION, minZoom: 5, maxZoom: 18 }).addTo(map);
    // 右下に、上からズームレベル・縮尺の目盛り(m・km)・地図の出典を並べる
    // (右下の欄は、後から加えたものほど上に積まれる)
    L.control.scale({ position: 'bottomright', imperial: false }).addTo(map);
    addZoomLevel(L);
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
      if (autoFit) fitToArea();
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

  // 地図に収める範囲は、読み込んだ地点すべて(写真の撮影場所・記録地点)の東西南北の端で決める
  // (平均ではないので、地点の数や偏りに左右されない)
  const latlngs = [...photos, ...tracks.flatMap((t) => t.points)].map((p) => [p.lat, p.lng]);
  fitArea = latlngs.length > 0 ? L.latLngBounds(latlngs) : null;
  if (!fitArea) return;
  autoFit = true;
  fitToArea();
}

// 地図を fitArea に合わせる: 範囲の真ん中を地図の中心にし、範囲がすべて収まるいちばん大きいズームレベル
// (ZOOM_SNAP 刻み。FIT_MAX_ZOOM まで)にする。最も北・南の地点が地図の上端・下端に、または最も西・東の地点が
// 左端・右端に、縁から FIT_PADDING の内側で届く(地図は正方形なので、範囲の縦横比により、どちらか一方が届く)
function fitToArea() {
  const L = globalThis.L;
  const size = map.getSize();
  if (!fitArea || size.x === 0 || size.y === 0) return; // 隠れている間は測れない
  const padding = L.point(FIT_PADDING * 2, FIT_PADDING * 2); // 上下・左右の余白の合計
  const zoom = Math.min(FIT_MAX_ZOOM, map.getBoundsZoom(fitArea, false, padding));
  // 範囲の真ん中は、地図の上(メルカトル図法の平面)で求める(上下の余白がそろうように)
  const nw = map.project(fitArea.getNorthWest(), zoom);
  const se = map.project(fitArea.getSouthEast(), zoom);
  const center = map.unproject(nw.add(se).divideBy(2), zoom);

  // ズームレベルが同じで、中心もずれていなければ動かさない(動かすと地図の画像をすべて読み直し、
  // 境界線を動かしている間などに地図がちらつく)
  const shift = map.latLngToContainerPoint(center).subtract(size.divideBy(2));
  if (Math.abs(zoom - map.getZoom()) < 1e-6 && Math.abs(shift.x) <= 1 && Math.abs(shift.y) <= 1) return;
  // すぐに動かす(動きを見せると、その間に利用者が動かしたのかを見分けられない)
  fitting = true;
  map.setView(center, zoom, { animate: false });
  fitting = false;
}

// 地図の右下に、ズームレベル(「ズームレベル 14.3」。整数なら小数を付けない)を出す
function addZoomLevel(L) {
  const control = L.control({ position: 'bottomright' });
  control.onAdd = () => {
    const div = L.DomUtil.create('div', 'map-zoom-level');
    const update = () => {
      div.textContent = `ズームレベル ${Math.round(map.getZoom() * 10) / 10}`;
    };
    map.on('zoomend', update);
    update();
    return div;
  };
  control.addTo(map);
}
