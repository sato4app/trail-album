// 別ウインドウの地図のページ(map.html)のエントリポイント。
// Trail Album の画面の地図の「別ウインドウで開く」から開かれ、画面の地図と同じ内容(写真の撮影場所・移動経路・凡例)を、
// ウインドウいっぱいの地図に出す。内容は元の画面から送られてきて、元の画面で KMZ・GPX を読み込んだ・クリアしたときも
// 合わせて変わる(map-window.js)。地図の描き方は画面の地図と同じ(map-view.js)。
// 写真のマーカーを押して出した縮小版を押すと、画面と同じく、元の写真を大きく表示する(写真は元の画面から送られてくる)。
// 選んだ写真は元の画面と知らせ合う: ここでマーカーを押すと、元の画面の写真の一覧にも枠が付き、
// 元の画面で選んだ写真は、ここの地図にも縮小版の吹き出しを出す。
import { canShowMap, updateMap, renderLegend, openPhotoPopup } from './map-view.js';
import { receiveMapView, notifySelectToOpener } from './map-window.js';

const $ = (id) => document.getElementById(id);

init();

function init() {
  setupPhotoDialog();
  if (!canShowMap()) {
    showMessage('地図を表示できませんでした。インターネットに接続していることを確かめてから、ページを再読み込みしてください。', 'error');
    return;
  }
  const connected = receiveMapView({
    onView: (view, { fit }) => {
      if (!view) {
        showMessage('表示する地点がありません(元の画面で、読み込んだ KMZ・GPX がクリアされました)。', 'guide');
        return;
      }
      // 地図を先に出す(隠れていると地図の大きさを測れない)
      showMessage('');
      $('map').hidden = false;
      $('mapLegend').hidden = false;
      updateMap($('map'), view, { fit, selectPhoto: (p) => notifySelectToOpener(p.no), openPhoto });
      renderLegend($('mapLegend'), view.legend);
    },
    onSelect: openPhotoPopup
  });
  if (!connected) {
    showMessage('この地図は、Trail Album の画面で、地図の右上の「別ウインドウで開く」を押して開いてください。', 'guide');
  }
}

// ===== 元の写真の拡大表示(画面の「拡大表示」と同じ見た目) =====

function setupPhotoDialog() {
  const dialog = $('photoDialog');
  // 画像の外側(背景)を押したら閉じる
  dialog.addEventListener('click', (e) => {
    if (e.target === dialog) dialog.close();
  });
  dialog.addEventListener('close', () => {
    URL.revokeObjectURL($('dialogImage').src);
    $('dialogImage').removeAttribute('src');
  });
}

// photo は updateMap に渡した写真の1つ(blob は元の写真)
function openPhoto(photo) {
  $('dialogImage').src = URL.createObjectURL(photo.blob);
  $('dialogImage').alt = photo.name;
  $('dialogCaption').textContent = [`No.${photo.no}`, photo.name, photo.timeLabel].filter(Boolean).join('  ');
  $('photoDialog').showModal();
}

// 地図の代わりに案内を出す(text が空なら、案内を消す)。kind は .message の data-kind(style.css)
function showMessage(text, kind = 'guide') {
  const el = $('mapMessage');
  el.textContent = text;
  el.dataset.kind = kind;
  el.hidden = !text;
  if (!text) return;
  $('map').hidden = true;
  $('mapLegend').hidden = true;
}
