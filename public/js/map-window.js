// 地図を別ウインドウで開く(画面の地図の右上の「別ウインドウで開く」。画面の地図はそのまま残す)。
// 別ウインドウ(map.html・map-window-page.js)にも、画面の地図と同じ内容(写真の撮影場所・移動経路・凡例)を出し、
// 画面の地図を描き直すたび(KMZ・GPX を読み込んだ・クリアした)に同じ内容を送って、別ウインドウの地図も合わせる。
//   1. 元の画面が、別ウインドウを開く(openMapWindow)
//   2. 別ウインドウは、読み込まれると元の画面へ「準備ができた」と知らせる(receiveMapView)
//   3. 元の画面は、今の内容を送る。その後も、画面の地図を描き直すたびに送る(sendToMapWindow)
// 選んだ写真(写真の一覧の枠・地図の縮小版の吹き出し)も、両方の向きで知らせ合う:
//   - 元の画面で選んだ(一覧の写真・画面の地図のマーカーを押した) → 別ウインドウの地図にも吹き出しを出す(selectInMapWindow)
//   - 別ウインドウの地図のマーカーを押した → 元の画面で同じ写真を選ぶ(notifySelectToOpener → connectMapWindow の onSelect)
// やりとりは postMessage で、このサイトの中(同じオリジン)の、開いた・開かれたウインドウとの間に限る
// (写真の名称・位置などを外へ出さない)。
import { MAP_WINDOW_WIDTH, MAP_WINDOW_HEIGHT } from './constants.js';

// 別ウインドウのページと、ウインドウの名前(同じ名前のウインドウがあれば、そこに開き直す)
export const MAP_WINDOW_PAGE = 'map.html';
export const MAP_WINDOW_NAME = 'trail-album-map';

// やりとりするメッセージの種類
const READY = 'trail-album:map-ready';   // 別ウインドウ → 元の画面: 準備ができた
const VIEW = 'trail-album:map-view';     // 元の画面 → 別ウインドウ: 地図の内容
const SELECT = 'trail-album:map-select'; // 両方の向き: 写真を選んだ(no は写真の通し番号)

const isOurs = (e, type) => e.origin === location.origin && e.data?.type === type;

// ===== 元の画面の側 =====

let mapWindow = null; // 開いた別ウインドウ(開いていなければ null)
let lastView = null;  // 最後に送った内容(地点が無ければ null)

// 別ウインドウからの知らせを受け取れるようにする(元の画面を開いたときに1回呼ぶ)。
// 「準備ができた」(別ウインドウを開いた・別ウインドウを再読み込みした)なら、今の内容を送る(地図は地点に合わせ直す)。
// 「写真を選んだ」なら、onSelect(no) を呼ぶ
export function connectMapWindow({ onSelect }) {
  window.addEventListener('message', (e) => {
    if (isOurs(e, READY) && e.source) {
      mapWindow = e.source;
      post({ type: VIEW, view: lastView, fit: true });
    } else if (isOurs(e, SELECT) && e.source && e.source === mapWindow) {
      onSelect(e.data.no);
    }
  });
}

// 地図を別ウインドウで開く。もう開いていれば、そのウインドウを前に出す(地図はそのまま)。
// ウインドウの大きさは MAP_WINDOW_WIDTH × MAP_WINDOW_HEIGHT(画面より大きければ画面に合わせる)。
// 開けなかった(ブラウザにポップアップを止められた)ときは false を返す
export function openMapWindow() {
  if (mapWindow && !mapWindow.closed) {
    mapWindow.focus();
    return true;
  }
  const width = Math.min(MAP_WINDOW_WIDTH, screen.availWidth);
  const height = Math.min(MAP_WINDOW_HEIGHT, screen.availHeight);
  const win = window.open(MAP_WINDOW_PAGE, MAP_WINDOW_NAME, `popup,width=${width},height=${height}`);
  if (!win) return false;
  mapWindow = win; // 内容は、別ウインドウの準備ができてから送る(connectMapWindow)
  return true;
}

// 画面の地図と同じ内容を、別ウインドウ(開いていれば)へ送る。
// view は { photos, tracks, legend }(map-view.js の updateMap・renderLegend に渡すもの)か、地点が無ければ null。
// fit は updateMap と同じ(false なら、別ウインドウの地図の中心・縮尺も今のままにする)
export function sendToMapWindow(view, { fit = true } = {}) {
  lastView = view;
  post({ type: VIEW, view, fit });
}

// 元の画面で選んだ写真(通し番号 no)の縮小版の吹き出しを、別ウインドウ(開いていれば)の地図にも出す
export function selectInMapWindow(no) {
  post({ type: SELECT, no });
}

function post(message) {
  if (!mapWindow || mapWindow.closed) return;
  mapWindow.postMessage(message, location.origin);
}

// ===== 別ウインドウの側 =====

// 元の画面へ「準備ができた」と知らせ、地図の内容が送られてくるたびに onView(view, { fit }) を呼ぶ
// (view・fit は sendToMapWindow と同じ)。元の画面で写真を選んだら、onSelect(no) を呼ぶ。
// 元の画面から開かれていない(このページを直接開いた・新しいタブで開いた)ときは false を返す
export function receiveMapView({ onView, onSelect }) {
  const opener = window.opener;
  if (!opener) return false;
  window.addEventListener('message', (e) => {
    if (e.source !== window.opener) return;
    if (isOurs(e, VIEW)) onView(e.data.view, { fit: e.data.fit });
    else if (isOurs(e, SELECT)) onSelect(e.data.no);
  });
  opener.postMessage({ type: READY }, location.origin);
  return true;
}

// 別ウインドウの地図で選んだ写真(通し番号 no)を、元の画面に知らせる
export function notifySelectToOpener(no) {
  window.opener?.postMessage({ type: SELECT, no }, location.origin);
}
