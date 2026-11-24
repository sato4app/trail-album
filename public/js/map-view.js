// 地理院地図に、写真の撮影場所と、移動経路の記録地点をすべてマーカーで表示する
// (Leaflet を CDN から読み込んで使う)。記録地点は数千になることもあるため、
// 画像のマーカーではなく、canvas に描く円のマーカー(circleMarker)にする。

const GSI_TILE_URL = 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png';
const GSI_ATTRIBUTION =
  '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>';
const INITIAL_CENTER = [34.853667, 135.472041]; // 箕面大滝
const INITIAL_ZOOM = 14;
const FIT_MAX_ZOOM = 17; // 地点が1つ・近いときに寄りすぎないように

// マーカーの色。写真は1色、経路は GPX ごとに色を変える(表の A, B … の印と同じ色)
export const PHOTO_COLOR = '#e8590c';
const TRACK_COLORS = ['#1c7ed6', '#7048e8', '#0c8599', '#d6336c', '#5c940d', '#495057'];
export const trackColor = (index) => TRACK_COLORS[index % TRACK_COLORS.length];

let map = null;
let markers = null;

// 地図を使えるか(Leaflet を CDN から読み込めたか)
export const canShowMap = () => !!globalThis.L;

// 地図のマーカーを置き直し、すべてのマーカーが入るように地図を合わせる。
// container は表示されていること(隠れていると地図の大きさを測れない)。
//   photos:  [{ lat, lng, label, name }](写真の撮影場所)
//   tracks:  [{ label, colorIndex, points: [{ lat, lng, timeLabel }] }](経路ごとの記録地点)
export function updateMap(container, { photos, tracks }) {
  const L = globalThis.L;
  if (!L) return;
  if (!map) {
    map = L.map(container, { preferCanvas: true }).setView(INITIAL_CENTER, INITIAL_ZOOM);
    L.tileLayer(GSI_TILE_URL, { attribution: GSI_ATTRIBUTION, minZoom: 5, maxZoom: 18 }).addTo(map);
    markers = L.featureGroup().addTo(map);
    // 地図の枠の大きさが変わったら(画面の幅いっぱいに広げた・KMZ と GPX の境界線を動かした)、
    // 地図の中心を保ったまま描き直す(地図の画像の読み込みは、動かし終えてからまとめて行う)
    new ResizeObserver(() => map.invalidateSize({ debounceMoveend: true })).observe(container);
  }
  map.invalidateSize(); // 隠れていた間に大きさが変わっていても合うように
  markers.clearLayers();

  // 経路の記録地点(写真の下に描く)
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

  const bounds = markers.getBounds();
  if (bounds.isValid()) map.fitBounds(bounds, { padding: [24, 24], maxZoom: FIT_MAX_ZOOM });
}
