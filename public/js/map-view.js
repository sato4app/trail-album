// 地理院地図に、写真の撮影場所と、移動経路の記録地点をすべてマーカーで表示し、
// 記録地点は経路ごとに記録の順に線で結ぶ(Leaflet を CDN から読み込んで使う)。
// 記録地点は数千になることもあるため、画像のマーカーではなく、canvas に描く円のマーカー(circleMarker)にする。
// 地図は、読み込んだ地点すべての東西南北の端が収まるように合わせる(地点の数や偏りに左右されない)。
// 右下に、「＋」「−」ボタン(左上のものと同じ)・ズームレベルのボタン(押すとズームレベルを変えるスライドバーを開く)・
// 縮尺の目盛りを出す。右上には、地図を別ウインドウで開くリンクを出せる(Trail Album の画面の地図だけ。map-window.js)。
// 写真の撮影場所のマーカーを押すと、写真の縮小版を出し、縮小版を押すと元の写真を大きく表示する。
// マーカーを押したことは呼び出し元にも知らせる(写真の一覧の行に枠を付けるため)。逆に、写真の一覧から
// 選んだ写真の縮小版を、地図に出すこともできる(openPhotoPopup)。
// 画面の地図と、別ウインドウの地図(map.html)の両方で使う。
// 地図の設定値(ズームレベルの範囲・色・マーカーの大きさなど)は constants.js に置く。
import {
  GSI_TILE_URL,
  GSI_ATTRIBUTION,
  GSI_TILE_MAX_NATIVE_ZOOM,
  MAP_MIN_ZOOM,
  MAP_MAX_ZOOM,
  ZOOM_SNAP,
  ZOOM_SLIDER_AUTO_CLOSE_SEC,
  INITIAL_CENTER,
  INITIAL_ZOOM,
  FIT_MAX_ZOOM,
  FIT_PADDING,
  PHOTO_COLOR,
  TRACK_COLORS,
  PHOTO_MARKER_RADIUS,
  TRACK_POINT_RADIUS,
  TRACK_LINE_WEIGHT,
  TRACK_LINE_OPACITY,
  PHOTO_POPUP_THUMB_SIZE
} from './constants.js';

// GPX ごとの色(表の 1, 2 … の印と同じ色。足りなければ繰り返す)
export const trackColor = (index) => TRACK_COLORS[index % TRACK_COLORS.length];

let map = null;
let markers = null;
let fitArea = null;  // 地図に収める範囲(読み込んだ地点すべての東西南北の端。LatLngBounds)。地点が無ければ null
let autoFit = false; // 利用者が地図を動かしていなければ true(地図の大きさが変わったら、合わせ直す)
let fitting = false; // fitToArea で地図を動かしている間だけ true
let onOpenPhoto = null;         // 写真の縮小版を押したときに呼ぶ関数(updateMap の openPhoto)
let onSelectPhoto = null;       // 写真のマーカーを押したときに呼ぶ関数(updateMap の selectPhoto)
const photoMarkers = new Map(); // 写真の通し番号(no)ごとのマーカー(openPhotoPopup で使う)
const thumbUrls = new Map();    // 写真ごとの縮小版の object URL(マーカーを置き直すときに捨てる)

// 写真の縮小版の吹き出しを出すとき、吹き出しが地図の左と上の縁からこれだけ離れるまで地図を動かす([左, 上] px)。
// 右上の「別ウインドウで開く」や左上の「＋」「−」ボタンが、吹き出し(閉じる「×」など)に重ならないように
const POPUP_PAN_PADDING_TOP_LEFT = [54, 44];

// 地図を使えるか(Leaflet を CDN から読み込めたか)
export const canShowMap = () => !!globalThis.L;

// 地図のマーカーを置き直し、すべてのマーカーが収まるように地図を合わせる(fitToArea)。
// fit が false なら、地図の中心・縮尺は今のままにする(KMZ・GPX をクリアしたとき)。
// container は表示されていること(隠れていると地図の大きさを測れない)。
//   photos:  [{ lat, lng, no, name, timeLabel, thumb }](写真の撮影場所。no は写真の通し番号、timeLabel は撮影日時、
//            thumb は縮小版 { blob, width, height }。縮小版を作れなかった(表示できない形式の)写真は null)
//   tracks:  [{ label, colorIndex, points: [{ lat, lng, timeLabel }] }](経路ごとの記録地点。label は 2-1 などの経路の番号、
//            timeLabel は記録日時)。日時が無ければ timeLabel は ''
// selectPhoto は、写真のマーカーを押したときに呼ぶ関数、openPhoto は、そのとき出した縮小版を押したときに呼ぶ関数
// (どちらも photos の1つを渡す)。
// windowLink を渡すと、地図の右上に「別ウインドウで開く」のリンクを出す(地図を初めて作るときだけ見る)。
//   { href, target, open }: href・target はリンク先のページとウインドウの名前。open はリンクを押したときに呼ぶ関数で、
//   別ウインドウを開けなかったら false を返す(そのときは、リンクのとおりに開く)
export function updateMap(container, { photos, tracks },
  { fit = true, selectPhoto = null, openPhoto = null, windowLink = null } = {}) {
  const L = globalThis.L;
  if (!L) return;
  onSelectPhoto = selectPhoto;
  onOpenPhoto = openPhoto;
  if (!map) {
    const zoomRange = { minZoom: MAP_MIN_ZOOM, maxZoom: MAP_MAX_ZOOM };
    map = L.map(container, { preferCanvas: true, zoomSnap: ZOOM_SNAP, ...zoomRange })
      .setView(INITIAL_CENTER, INITIAL_ZOOM);
    L.tileLayer(GSI_TILE_URL, { attribution: GSI_ATTRIBUTION, ...zoomRange, maxNativeZoom: GSI_TILE_MAX_NATIVE_ZOOM })
      .addTo(map);
    // 右下に、上から「＋」「−」ボタン・ズームレベルのボタンとスライドバー・縮尺の目盛り(m・km)・地図の出典を並べる
    // (右下の欄は、後から加えたものほど上に積まれる。左上の「＋」「−」ボタンも Leaflet の既定のまま残す)
    L.control.scale({ position: 'bottomright', imperial: false }).addTo(map);
    addZoomControl(L);
    L.control.zoom({ position: 'bottomright', zoomInTitle: '拡大', zoomOutTitle: '縮小' }).addTo(map);
    if (windowLink) addWindowLink(L, windowLink);
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
  markers.clearLayers(); // 出していた写真の縮小版も閉じる
  photoMarkers.clear();
  for (const url of thumbUrls.values()) URL.revokeObjectURL(url);
  thumbUrls.clear();

  // 経路の線と記録地点(写真の下に描く。先に加えたものほど下になる)。
  // 線は経路(1-1, 1-2 …)ごとに引き、別の経路どうしはつながない。
  // マウスを乗せたときの表示は、線は「移動経路 2-1」、記録地点は「移動経路 2-1 の38点目」と記録日時の2行
  for (const track of tracks) {
    const color = trackColor(track.colorIndex);
    L.polyline(track.points.map((p) => [p.lat, p.lng]), { color, weight: TRACK_LINE_WEIGHT, opacity: TRACK_LINE_OPACITY })
      .bindTooltip(`移動経路 ${track.label}`, { sticky: true })
      .addTo(markers);
  }
  for (const track of tracks) {
    const color = trackColor(track.colorIndex);
    track.points.forEach((p, i) => {
      L.circleMarker([p.lat, p.lng], { radius: TRACK_POINT_RADIUS, weight: 1, color, fillColor: color, fillOpacity: 0.85 })
        .bindTooltip(() => tooltipLines([`移動経路 ${track.label} の${i + 1}点目`, p.timeLabel]))
        .addTo(markers);
    });
  }
  // 写真の撮影場所。マウスを乗せたときの表示は「写真No.13」・名称・撮影日時の3行と、押すと写真を出すことの案内。
  // 押すと、写真の縮小版を出し(photoPopup)、押したことを知らせる(selectPhoto)。
  // 縮小版を出している間は、同じ写真のマウスを乗せたときの表示は出さない
  for (const photo of photos) {
    const marker = L.circleMarker([photo.lat, photo.lng], {
      radius: PHOTO_MARKER_RADIUS, weight: 2, color: '#fff', fillColor: PHOTO_COLOR, fillOpacity: 1
    })
      .bindTooltip(() => tooltipLines([`写真No.${photo.no}`, photo.name, photo.timeLabel],
        photo.thumb ? '押すと写真を表示' : ''))
      .bindPopup(() => photoPopup(photo), {
        maxWidth: PHOTO_POPUP_THUMB_SIZE,
        className: 'map-photo-popup',
        autoPanPaddingTopLeft: POPUP_PAN_PADDING_TOP_LEFT
      })
      .addTo(markers);
    marker.on('popupopen', () => marker.closeTooltip());
    marker.on('tooltipopen', () => {
      if (marker.isPopupOpen()) marker.closeTooltip();
    });
    // 押したときだけ知らせる(openPhotoPopup で開いたときは知らせない。知らせた先から開き直されて、繰り返さないように)
    marker.on('click', () => onSelectPhoto?.(photo));
    photoMarkers.set(photo.no, marker);
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

// 写真の通し番号 no の、写真の縮小版の吹き出しを出す(写真の一覧で写真を選んだとき・別ウインドウで選んだとき)。
// 写真が地図の見えている範囲の外なら、吹き出しが見えるところまで地図を動かす(Leaflet の autoPan)。
// もう出していれば、そのままにする
export function openPhotoPopup(no) {
  const marker = photoMarkers.get(no);
  if (marker && !marker.isPopupOpen()) marker.openPopup();
}

// 凡例を el に出す。legend は [{ color, text, kind }](kind は 'photo'(写真の撮影場所)か 'track'(移動経路))。
// 印の形は style.css の .legend-item
export function renderLegend(el, legend) {
  el.replaceChildren(...legend.map(({ color, text, kind }) => {
    const item = document.createElement('span');
    item.className = `legend-item is-${kind}`;
    const dot = document.createElement('span');
    dot.className = 'legend-dot';
    dot.style.backgroundColor = color;
    dot.style.color = color; // 経路の線(style.css の .is-track)の色
    item.append(dot, text);
    return item;
  }));
}

// 地図の右上に「別ウインドウで開く」のリンクを出す(windowLink は updateMap と同じ)。
// Ctrl・Shift などを押しながらのクリックは、ブラウザに任せる(新しいタブで開くなど)
function addWindowLink(L, { href, target, open }) {
  const control = L.control({ position: 'topright' });
  control.onAdd = () => {
    const link = L.DomUtil.create('a', 'map-window-link');
    link.href = href;
    link.target = target;
    link.title = '地図を別ウインドウで開きます(この画面の地図は、そのまま残ります)';
    // 印は「枠から矢印が飛び出す」形(固定の文字なので innerHTML で書く)
    link.innerHTML =
      '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/></svg>' +
      '別ウインドウで開く';
    L.DomEvent.on(link, 'click', (e) => {
      if (e.button !== 0 || e.ctrlKey || e.shiftKey || e.metaKey || e.altKey) return;
      if (open()) e.preventDefault();
    });
    // リンクの操作で、地図がドラッグ・拡大されないように
    L.DomEvent.disableClickPropagation(link);
    return link;
  };
  control.addTo(map);
}

// マウスを乗せたときの表示の中身: lines を1行ずつ並べる(空の行は出さない)。hint があれば、最後に薄い字で添える。
// 名称は利用者が付けた文字なので、HTML として解釈させずに文字のまま出す。
// 記録地点は数千になることもあるため、表示するときに初めて作る(bindTooltip に関数で渡す)
function tooltipLines(lines, hint = '') {
  const div = document.createElement('div');
  for (const text of lines) {
    if (!text) continue;
    const line = document.createElement('div');
    line.textContent = text;
    div.append(line);
  }
  if (hint) {
    const line = document.createElement('div');
    line.className = 'map-tooltip-hint';
    line.textContent = hint;
    div.append(line);
  }
  return div;
}

// 写真のマーカーを押したときに出す中身: 写真の縮小版(長辺 PHOTO_POPUP_THUMB_SIZE まで)と、その下に
// 「写真No.13」・名称・撮影日時。縮小版を押すと、元の写真を大きく表示する(updateMap の openPhoto)。
// 縮小版を作れなかった写真は、表示できない形式であることを書く。
// 写真は数百枚になることもあるため、押したときに初めて作る(bindPopup に関数で渡す)
function photoPopup(photo) {
  const div = document.createElement('div');
  div.className = 'map-photo';
  if (photo.thumb) {
    // 画像の大きさを先に決めておく(読み込み前後で吹き出しの大きさ・位置が変わらないように)
    const { width, height } = photo.thumb;
    const scale = Math.min(1, PHOTO_POPUP_THUMB_SIZE / Math.max(width, height));
    const img = document.createElement('img');
    if (!thumbUrls.has(photo)) thumbUrls.set(photo, URL.createObjectURL(photo.thumb.blob));
    img.src = thumbUrls.get(photo);
    img.alt = '';
    img.width = Math.max(1, Math.round(width * scale));
    img.height = Math.max(1, Math.round(height * scale));
    if (onOpenPhoto) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'map-photo-thumb';
      button.title = '元の写真を大きく表示';
      button.setAttribute('aria-label', `${photo.name} を大きく表示`);
      button.append(img);
      button.addEventListener('click', () => onOpenPhoto(photo));
      div.append(button);
    } else {
      img.className = 'map-photo-thumb';
      div.append(img);
    }
  } else {
    const missing = document.createElement('div');
    missing.className = 'map-photo-missing';
    missing.textContent = 'この形式は表示できません';
    div.append(missing);
  }
  const info = tooltipLines([`写真No.${photo.no}`, photo.name, photo.timeLabel]);
  info.className = 'map-photo-info';
  div.append(info);
  return div;
}

// 地図を fitArea に合わせる:範囲の真ん中を地図の中心にし、範囲がすべて収まるいちばん大きいズームレベル
// (ZOOM_SNAP 刻み。FIT_MAX_ZOOM まで。MAP_MIN_ZOOM でも収まらなければ MAP_MIN_ZOOM)にする。最も北・南の地点が地図の上端・下端に、または最も西・東の地点が
// 左端・右端に、縁から FIT_PADDING の内側で届く(地図と範囲の縦横比により、どちらか一方が届く。
// 地図は正方形とは限らない: 画面に固定するとき・別ウインドウでは横長になる)
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

// 地図の右下に、ズームレベルのボタン(「Z=16.3」。整数なら小数を付けない)を出す。
// ボタンを押すと、その左にズームレベルを変えるスライドバーを開く(もう一度押すか Esc で閉じる)。
// 地図を拡大・縮小したとき(「＋」「−」ボタン・ホイール・ダブルクリックなど)も開き、こちらは
// 操作が無いまま ZOOM_SLIDER_AUTO_CLOSE_SEC 秒たったら閉じる。ボタンで開いたときは自動では閉じないが、
// その後にズームレベルを変えたら(拡大・縮小・スライドバー)、同じく秒数がたったら閉じる。
// ボタンは右端の同じ位置のままにする(開け閉めでボタンが動かないように)。
// スライドバーは MAP_MIN_ZOOM〜MAP_MAX_ZOOM を ZOOM_SNAP(0.1)刻みで動かし、地図の中心を保ったまま拡大・縮小する。
function addZoomControl(L) {
  // 表示する小数の桁数(ZOOM_SNAP が 0.1 なら1桁)
  const digits = (String(ZOOM_SNAP).split('.')[1] || '').length;
  const format = (zoom) => String(Number(zoom.toFixed(digits)));

  const control = L.control({ position: 'bottomright' });
  control.onAdd = () => {
    // 左にスライドバー、右にボタンを置く(見た目の並びと、Tab キーで移る順をそろえる)
    const div = L.DomUtil.create('div', 'map-zoom');
    const slider = L.DomUtil.create('input', 'map-zoom-slider', div);
    slider.type = 'range';
    slider.id = 'mapZoomSlider';
    slider.min = String(MAP_MIN_ZOOM);
    slider.max = String(MAP_MAX_ZOOM);
    slider.step = String(ZOOM_SNAP);
    slider.hidden = true;
    slider.setAttribute('aria-label', 'ズームレベル');
    const button = L.DomUtil.create('button', 'map-zoom-button', div);
    button.type = 'button';
    button.title = 'ズームレベル(押すと、ズームレベルを変えるスライドバーを開きます)';
    button.setAttribute('aria-expanded', 'false');
    button.setAttribute('aria-controls', 'mapZoomSlider');

    // ボタンの文字(スライドバーを動かしている間は、選んでいる値を先に出す)
    const showZoom = (zoom) => {
      const text = format(zoom);
      button.textContent = `Z=${text}`;
      slider.setAttribute('aria-valuetext', text);
    };

    // スライドバーで選んだズームレベル(地図が追いつくまで持っておく。追いついたら null)。
    // 地図は拡大・縮小の動きを見せながら変える(一気に変えると地図の画像をすべて読み直し、ちらつくため)。
    // 動いている間に選んだ値は Leaflet が受け付けないので、動き終えるたび(zoomend)に、選んだ値へ向けて動かし直す
    let target = null;
    let zoomingBySlider = false; // 今の拡大・縮小は、スライドバーで選んだ値へ向けたものか
    const follow = () => {
      if (target === null) return;
      if (Math.abs(target - map.getZoom()) < 1e-6) {
        target = null;
        return;
      }
      zoomingBySlider = true;
      map.setZoom(target);
    };

    // スライドバーをどう開いたか: null(閉じている)・'button'(ボタンで開き、まだズームレベルを変えていない。自動では閉じない)・
    // 'zoom'(ズームレベルを変えた。操作が無いまま ZOOM_SLIDER_AUTO_CLOSE_SEC 秒たったら閉じる)
    let openedBy = null;
    let closeTimer = 0;
    let pointerInside = false; // マウスがボタン・スライドバーの上にある(タッチなら触れている)間は閉じない
    let pressing = false;      // スライドバーを押さえている(ドラッグしている)間は閉じない

    const setOpen = (by) => {
      openedBy = by;
      const open = by !== null;
      slider.hidden = !open;
      button.setAttribute('aria-expanded', String(open));
      button.title = `ズームレベル(押すと、ズームレベルを変えるスライドバーを${open ? '閉じます' : '開きます'})`;
      restartCloseTimer();
    };
    // 自動で開いたときは、操作が無いまま ZOOM_SLIDER_AUTO_CLOSE_SEC 秒たったら閉じる(操作するたびに数え直す)
    const restartCloseTimer = () => {
      clearTimeout(closeTimer);
      if (openedBy !== 'zoom' || pointerInside || pressing) return;
      closeTimer = setTimeout(() => {
        // キーボードでスライドバーを操作していた人が、フォーカスを失わないように
        if (document.activeElement === slider) button.focus();
        setOpen(null);
      }, ZOOM_SLIDER_AUTO_CLOSE_SEC * 1000);
    };

    // 地図のズームレベルが変わったら(スライドバー・＋−ボタン・ホイール・地図を合わせたとき)、表示を合わせる。
    // 利用者が拡大・縮小したときは、スライドバーを開き、操作が無いまま ZOOM_SLIDER_AUTO_CLOSE_SEC 秒たったら閉じる
    // (ボタンで開いていたときも同じ)。地点に地図を合わせたとき(読み込んだ・地図の大きさが変わった)は開かない。
    // スライドバーで変えたときは、閉じるまでの秒数を数え直すだけにする(閉じた後に地図が追いついても開き直さない)
    map.on('zoomend', () => {
      const bySlider = zoomingBySlider;
      zoomingBySlider = false;
      follow();
      if (target === null) {
        showZoom(map.getZoom());
        slider.value = format(map.getZoom());
      }
      if (fitting) return;
      if (bySlider) restartCloseTimer();
      else setOpen('zoom'); // ボタンで開いていたときも、拡大・縮小したら自動で閉じる開き方に変える
    });
    L.DomEvent.on(slider, 'input', () => {
      target = Number(slider.value);
      showZoom(target);
      follow();
      setOpen('zoom'); // ボタンで開いていたときも、スライドバーで変えたら自動で閉じる開き方に変える
    });

    L.DomEvent.on(button, 'click', () => setOpen(openedBy === null ? 'button' : null));
    L.DomEvent.on(div, 'keydown', (e) => {
      L.DomEvent.stopPropagation(e); // 地図のキー操作(矢印キーで地図を動かすなど)に渡さない
      if (e.key === 'Escape' && openedBy !== null) {
        setOpen(null);
        button.focus();
      } else {
        restartCloseTimer();
      }
    });
    L.DomEvent.on(div, 'pointerenter', () => {
      pointerInside = true;
      restartCloseTimer();
    });
    L.DomEvent.on(div, 'pointerleave', () => {
      pointerInside = false;
      restartCloseTimer();
    });
    L.DomEvent.on(slider, 'pointerdown', () => {
      pressing = true;
      restartCloseTimer();
      const release = () => {
        pressing = false;
        document.removeEventListener('pointerup', release);
        document.removeEventListener('pointercancel', release);
        restartCloseTimer();
      };
      document.addEventListener('pointerup', release);
      document.addEventListener('pointercancel', release);
    });
    // ボタン・スライドバーの操作で、地図がドラッグ・拡大されないように
    L.DomEvent.disableClickPropagation(div);
    L.DomEvent.disableScrollPropagation(div);

    showZoom(map.getZoom());
    slider.value = format(map.getZoom());
    return div;
  };
  control.addTo(map);
}
