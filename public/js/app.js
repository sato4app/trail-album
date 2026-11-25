// Trail Album のエントリポイント(画面の制御)
// KMZ・GPX を読み込んでプレビューを表示し、写真の一覧の Excel と元の画質の写真(KMZ があるとき)、
// 読み込んだ KMZ・GPX ファイル(選んだとき)を、利用者が付けた名前のフォルダにまとめて、
// フォルダへ直接保存するか、ZIP でダウンロードする。処理はすべてブラウザの中で行う。
// KMZ・GPX はいくつでも読み込め、後から読み込んだものは今の内容に追加する(merge.js)。
// 保存される内容(フォルダ・ファイルの木)は、名前やチェックを変えるたびに画面の中で作り直す。
import { readKmz, KmzError } from './kmz-reader.js';
import { readGpx, GpxError } from './gpx-reader.js';
import { excludeLoadedPhotos, buildAlbum, addSegments, buildTrack } from './merge.js';
import { createThumbnail } from './thumbnail.js';
import { canCopyImage, copyImage, saveImage, flashButton } from './photo-actions.js';
import { canShowMap, updateMap, renderLegend, openPhotoPopup, trackColor } from './map-view.js';
import {
  MAP_WINDOW_PAGE,
  MAP_WINDOW_NAME,
  connectMapWindow,
  openMapWindow,
  sendToMapWindow,
  selectInMapWindow
} from './map-window.js';
import { setupColSplitter } from './col-splitter.js';
import {
  planPhotos,
  defaultSaveName,
  saveNameCandidates,
  checkSaveName,
  buildSavePlan,
  formatJstDateTime,
  formatJstDateTimeSeconds,
  formatJstRange
} from './naming.js';
import { buildWorkbookBlob } from './excel-writer.js';
import { buildMergedKmz, buildMergedGpx } from './merge-writer.js';
import { PHOTO_COLOR, PHOTOS_DIR, MAX_NAME_LENGTH } from './constants.js';
import {
  canSaveToFolder,
  loadSaveDir,
  currentSaveDir,
  pickSaveDir,
  clearSaveDir,
  prepareSaveDir,
  saveToFolder,
  buildZipBlob,
  downloadBlob
} from './saver.js';

const $ = (id) => document.getElementById(id);

const state = {
  kmzFiles: [],    // 読み込んだ KMZ(readKmz の結果。読み込んだ順。写真は読み込み済みのものと重ならない分だけ)
  album: null,     // kmzFiles をまとめたもの(merge.js の buildAlbum の結果)。KMZ が無ければ null
  segments: [],    // 読み込んだ GPX の経路(merge.js の addSegments で重なりを除いたもの。読み込んだ順)
  track: null,     // segments をまとめたもの(merge.js の buildTrack の結果)。GPX が無ければ null
  photoRows: [],   // 写真の一覧に出している写真(保存の順。No. は保存の通し番号 no。地図にも使う)
  urls: [],        // プレビューで作った画像の URL(写真の一覧を作り直すときに解放する)
  busy: false,     // 読み込み中・保存中
  save: {
    defaultName: '',  // 名前の既定値(読み込んだ内容から作る: naming.js の defaultSaveName)
    edited: false,    // 利用者が名前を既定値から変えたか(変えたら、KMZ・GPX を追加しても名前を書き換えない)
    includeKmz: true, // KMZ から作るもの(Excel・写真・KMZ)を保存するか(読み込んでいない間も、選んだ状態を覚えておく)
    includeGpx: true  // GPX を保存するか
  }
};

init();

function init() {
  // CDN のライブラリが読み込めていなければ、操作できないことを伝える
  if (!globalThis.JSZip || !globalThis.ExcelJS) {
    showStatus([{
      message: '必要なライブラリを読み込めませんでした。インターネットに接続していることを確かめてから、ページを再読み込みしてください。',
      kind: 'error'
    }]);
    $('btnPickFile').disabled = true;
    return;
  }

  $('btnSaveFolder').hidden = !canSaveToFolder;
  $('guideFolder').hidden = !canSaveToFolder;
  $('saveDirRow').hidden = !canSaveToFolder;
  if (!canSaveToFolder) $('saveHint').hidden = false;
  loadSaveDir().then(updateSaveDirLabel);

  $('btnPickFile').addEventListener('click', () => $('fileInput').click());
  $('fileInput').addEventListener('change', (e) => {
    const files = Array.from(e.target.files || []);
    e.target.value = ''; // 同じファイルを続けて選んでも change が起きるように
    loadFiles(files);
  });
  setupDragAndDrop();
  setupColSplitter($('previewCols'), $('colSplitter'), $('kmzCol'));
  connectMapWindow({ onSelect: (no) => selectPhoto(no, 'window') }); // 別ウインドウの地図でマーカーを押した
  $('btnClearAll').addEventListener('click', onClearAll);
  $('btnClearKmz').addEventListener('click', onClearKmz);
  $('btnClearGpx').addEventListener('click', onClearGpx);

  $('btnSaveFolder').addEventListener('click', onSaveFolder);
  $('btnPickSaveDir').addEventListener('click', onPickSaveDir);
  $('btnClearSaveDir').addEventListener('click', onClearSaveDir);
  $('btnSaveZip').addEventListener('click', onSaveZip);
  setupSaveForm();

  const dialog = $('photoDialog');
  // 画像の外側(背景)を押したら閉じる
  dialog.addEventListener('click', (e) => {
    if (e.target === dialog) dialog.close();
  });
  dialog.addEventListener('close', () => {
    $('dialogImage').removeAttribute('src');
  });

  // 保存の途中でページを閉じようとしたら確認する
  window.addEventListener('beforeunload', (e) => {
    if (state.busy) e.preventDefault();
  });
}

// ===== ドラッグ＆ドロップ =====
// ページのどこに落としても受け付ける(受け付け領域から外れて落とすと、
// ブラウザが KMZ をダウンロードしたり、GPX をページの代わりに開いたりしてしまうため)。
function setupDragAndDrop() {
  const zone = $('dropZone');
  let depth = 0;
  const hasFiles = (e) => Array.from(e.dataTransfer?.types || []).includes('Files');

  window.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    depth++;
    if (!state.busy) zone.classList.add('is-over');
  });
  window.addEventListener('dragleave', (e) => {
    if (!hasFiles(e)) return;
    depth = Math.max(0, depth - 1);
    if (depth === 0) zone.classList.remove('is-over');
  });
  window.addEventListener('dragover', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = state.busy ? 'none' : 'copy';
  });
  window.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    depth = 0;
    zone.classList.remove('is-over');
    if (state.busy) return;
    loadFiles(Array.from(e.dataTransfer.files));
  });
}

// ===== KMZ・GPX の読み込み =====
// 選んだ(置いた)ファイルを、KMZ・GPX ともすべて読み込み、今の内容に追加する。
// 同じ写真・経路は重ねて追加しない(merge.js)。読み込めなかったファイルがあっても、他のファイルは読み込む。
async function loadFiles(files) {
  if (files.length === 0 || state.busy) return;
  setBusy(true);
  const results = []; // 読み込みの結果(1ファイル1行)
  const progress = (message) => showStatus([...results, { message }]);
  const kmzCountBefore = state.kmzFiles.length;
  const segmentsBefore = state.segments;
  try {
    const { kmz, gpx, unknown } = await groupFiles(files);
    for (const file of kmz) results.push(await loadKmz(file, progress));
    for (const file of gpx) results.push(await loadGpx(file, progress));
    if (unknown.length > 0) {
      const names = unknown.map((f) => `「${f.name}」`).join('');
      results.push({ message: `${names}は KMZ, GPX ファイルではないため、読み込みませんでした。`, kind: 'error' });
    }
  } catch (err) {
    console.error(err);
    results.push({ message: `読み込みに失敗しました(${err.message || err})。`, kind: 'error' });
  } finally {
    // 表示は、追加があった側だけ最後に1回作り直す
    if (state.kmzFiles.length !== kmzCountBefore) showKmzPreview();
    if (state.segments !== segmentsBefore) showGpxPreview();
    showStatus(results);
    setBusy(false);
  }
}

// ファイルを KMZ・GPX・それ以外に分ける(並びは選んだ順のまま)
async function groupFiles(files) {
  const groups = { kmz: [], gpx: [], unknown: [] };
  for (const file of files) groups[(await fileKind(file)) || 'unknown'].push(file);
  return groups;
}

// ファイルの種類('kmz' | 'gpx' | null)。拡張子で決め、拡張子で分からないとき
// (拡張子が無い・.xml になっているなど)は中身の先頭で見分ける
async function fileKind(file) {
  if (/\.kmz$/i.test(file.name)) return 'kmz';
  if (/\.gpx$/i.test(file.name)) return 'gpx';
  let head;
  try {
    head = new TextDecoder().decode(await file.slice(0, 1024).arrayBuffer());
  } catch {
    return null;
  }
  if (/<gpx[\s>]/.test(head)) return 'gpx';
  // ZIP の印(PK)で始まり、最初に入っているファイルが KML なら KMZ
  // (Excel の .xlsx なども ZIP なので、ZIP というだけでは KMZ とみなさない)
  if (head.startsWith('PK') && /\.kml/i.test(head)) return 'kmz';
  return null;
}

// KMZ を読み込み、読み込み済みの写真と重ならない写真を追加する(表示は loadFiles が最後に作り直す)。
// 結果の1行({ message, kind })を返す。読み込めなかったときは、今の内容をそのまま残す
async function loadKmz(file, progress) {
  progress(`「${file.name}」を読み込み中…`);
  try {
    const kmz = await readKmz(file);
    const photos = excludeLoadedPhotos(state.kmzFiles, kmz.photos);
    const dupCount = kmz.photos.length - photos.length;
    if (photos.length === 0) {
      return { message: `「${file.name}」の写真 ${dupCount} 枚は、すべて読み込み済みです(追加した写真はありません)。`, kind: 'guide' };
    }
    for (let i = 0; i < photos.length; i++) {
      progress(`「${file.name}」のサムネイルを作成中… ${i + 1} / ${photos.length}`);
      const photo = photos[i];
      photo.blob = new Blob([photo.bytes], { type: photo.type });
      photo.thumb = await createThumbnail(photo.blob);
    }
    state.kmzFiles.push({ ...kmz, photos });
    const note = dupCount > 0 ? `(読み込み済みの写真と同じ ${dupCount} 枚は除きました)` : '';
    return { message: `「${file.name}」を読み込み、写真を ${photos.length} 枚追加しました。${note}`, kind: 'ok' };
  } catch (err) {
    console.error(err);
    const message = err instanceof KmzError ? err.message : `読み込みに失敗しました(${err.message || err})。`;
    return { message: `「${file.name}」: ${message}`, kind: 'error' };
  }
}

// GPX を読み込み、経路を追加する(表示は loadFiles が最後に作り直す)。
// 読み込み済みの経路と同じ経路は、記録地点の多いほうを残す(merge.js の addSegments)。
// 結果の1行({ message, kind })を返す。読み込めなかったときは、今の内容をそのまま残す
async function loadGpx(file, progress) {
  progress(`「${file.name}」を読み込み中…`);
  try {
    const { segments, added, replaced, skipped } = addSegments(state.segments, await readGpx(file));
    if (added.length + replaced.length === 0) {
      return { message: `「${file.name}」の経路 ${skipped.length} 本は、すべて読み込み済みです(追加した経路はありません)。`, kind: 'guide' };
    }
    state.segments = segments;
    const parts = [];
    if (added.length > 0) parts.push(`経路を ${added.length} 本(記録地点 ${countPoints(added)} 点)追加しました。`);
    if (replaced.length > 0) {
      parts.push(`読み込み済みの経路 ${replaced.length} 本を、記録地点が多いこのファイルの経路に置き換えました。`);
    }
    const note = skipped.length > 0 ? `(読み込み済みの経路と同じ ${skipped.length} 本は除きました)` : '';
    return { message: `「${file.name}」を読み込み、${parts.join('')}${note}`, kind: 'ok' };
  } catch (err) {
    console.error(err);
    const message = err instanceof GpxError ? err.message : `読み込みに失敗しました(${err.message || err})。`;
    return { message: `「${file.name}」: ${message}`, kind: 'error' };
  }
}

// 経路の記録地点の合計(1,234 のように桁区切りで)
function countPoints(segments) {
  return segments.reduce((sum, s) => sum + s.points.length, 0).toLocaleString('ja-JP');
}

// 記録地点数。経路が2本以上なら経路数も付ける(1,234 点(3 経路))
function formatPoints(segments) {
  const routes = segments.length > 1 ? `(${segments.length} 経路)` : '';
  return `${countPoints(segments)} 点${routes}`;
}

// 読み込んだ KMZ・GPX をすべて画面から外し、ページを開いたときの状態に戻す(ファイルは消えない)。
// 再読み込み(F5)と同じく、保存の設定(名前・保存するもの)も最初に戻す。保存先は、再読み込みしても残るものなので残す。
// 地図を別ウインドウで開いていれば、そちらには「表示する地点がありません」と出る
function onClearAll() {
  if (state.busy || (!state.album && !state.track)) return;
  state.kmzFiles = [];
  state.segments = [];
  state.track = null; // 先に KMZ 側を作り直すとき、地図に GPX の経路を描き直さないように
  showKmzPreview();
  showGpxPreview();
  state.save.includeKmz = true;
  state.save.includeGpx = true;
  showStatus([{ message: '読み込んだ KMZ・GPX をすべてクリアしました。', kind: 'ok' }]);
  $('btnPickFile').focus(); // 押したボタンが消えるので、次に押すボタンへ移す(キーボードで操作している人のため)
}

// 読み込んだ KMZ を画面から外す(ファイルは消えない)
function onClearKmz() {
  if (state.busy || state.kmzFiles.length === 0) return;
  state.kmzFiles = [];
  showKmzPreview({ fit: false }); // 地図の中心・縮尺は変えない
  showStatus([{ message: '読み込んだ KMZ をクリアしました。', kind: 'ok' }]);
}

// 読み込んだ GPX を画面から外す(ファイルは消えない)
function onClearGpx() {
  if (state.busy || state.segments.length === 0) return;
  state.segments = [];
  showGpxPreview({ fit: false }); // 地図の中心・縮尺は変えない
  showStatus([{ message: '読み込んだ GPX をクリアしました。', kind: 'ok' }]);
}

// 左側(KMZ)の表示を、読み込んだ KMZ(state.kmzFiles)から作り直す。fit は showMap と同じ
function showKmzPreview({ fit = true } = {}) {
  for (const url of state.urls) URL.revokeObjectURL(url);
  state.urls = [];
  const album = buildAlbum(state.kmzFiles);
  state.album = album;

  $('kmzEmpty').hidden = !!album;
  $('kmzPreview').hidden = !album;
  $('btnClearKmz').hidden = !album;
  updatePreviewSection();

  // 写真は保存の順(すべての KMZ を通した撮影日時の順)に並べ、No. は保存するファイル名・Excel と同じ通し番号にする
  state.photoRows = album ? planPhotos(album) : [];
  if (album) {
    // KMZ ごとの写真の枚数(読み込み済みの写真と同じものを除いた、一覧に出している枚数)
    $('sumFile').replaceChildren(...album.files.map((f, i) =>
      fileLine(fileNo(i), f.fileName, PHOTO_COLOR, [`${f.photos.length} 枚`])
    ));
    $('sumCount').textContent = `${album.photos.length} 枚`;
    $('sumCountItem').hidden = album.files.length < 2; // 1つならファイルの行と同じ数なので出さない
  }
  $('photoRows').replaceChildren(...state.photoRows.map(photoRow));
  showMap({ fit });
  refreshSaveSection();
}

// 右側(GPX)の表示を、読み込んだ経路(state.segments)から作り直す。fit は showMap と同じ
function showGpxPreview({ fit = true } = {}) {
  const track = buildTrack(state.segments);
  state.track = track;

  $('gpxEmpty').hidden = !!track;
  $('gpxPreview').hidden = !track;
  $('btnClearGpx').hidden = !track;
  updatePreviewSection();

  if (track) {
    // GPX ごとの記録地点数(経路数)・移動距離(読み込み済みの経路と同じものを除いた、内訳に出している経路の分)
    $('gpxFile').replaceChildren(...track.files.map((f, i) => fileLine(fileNo(i), f.source.fileName, trackColor(i), [
      formatPoints(f.segments),
      formatDistance(f.segments.reduce((sum, s) => sum + s.distanceM, 0))
    ])));
    $('gpxPoints').textContent = formatPoints(track.files.flatMap((f) => f.segments));
    $('gpxDistance').textContent = formatDistance(track.distanceM);
    const single = track.files.length < 2; // 1つならファイルの行と同じ数なので、合計は出さない
    $('gpxPointsItem').hidden = single;
    $('gpxDistanceItem').hidden = single;
  }
  // 経路ごとの内訳: GPX ごとに、ファイル名の行と、その経路(1-1, 1-2 …)の行を並べる
  $('segRows').replaceChildren(...(track?.files ?? []).flatMap((f, i) => [
    segmentGroupRow(fileNo(i), f.source.fileName, trackColor(i)),
    ...f.segments.map((segment, j) => segmentRow(segment, segmentNo(i, j)))
  ]));
  showMap({ fit });
  refreshSaveSection();
}

// KMZ・GPX のどちらかを読み込んでいるときだけ「内容を確かめる」と、上の「すべてクリア」を出す
function updatePreviewSection() {
  const hasData = !!state.album || !!state.track;
  $('previewSection').hidden = !hasData;
  $('btnClearAll').hidden = !hasData;
}

// KMZ・GPX の番号(それぞれ 1, 2, 3 …)。経路の番号は、GPX の番号と GPX の中での順番(1-1, 1-2, 2-1 …)
const fileNo = (index) => String(index + 1);
const segmentNo = (fileIndex, segmentIndex) => `${fileIndex + 1}-${segmentIndex + 1}`;

// ファイル名の行: 左に番号(色は地図のマーカーの色)とファイル名、
// 右にファイルごとの数(stats: 写真の枚数、記録地点数・移動距離)。数は行どうしで列をそろえる(style.css の .file-list)
function fileLine(no, name, color, stats) {
  const div = document.createElement('div');
  div.className = 'file-line';
  const nameEl = document.createElement('span');
  nameEl.className = 'file-name';
  nameEl.append(fileBadge(no, color), name);
  div.append(nameEl, ...stats.map((text) => {
    const span = document.createElement('span');
    span.className = 'file-stat';
    span.textContent = text;
    return span;
  }));
  return div;
}

function fileBadge(no, color) {
  const badge = document.createElement('span');
  badge.className = 'file-badge';
  badge.style.backgroundColor = color;
  badge.textContent = no;
  return badge;
}

// 8.4 km(minoh-hiking の統計表と同じく小数点以下1位)
function formatDistance(meters) {
  return `${(meters / 1000).toFixed(1)} km`;
}

// 経路ごとの内訳の、GPX のファイル名の行
function segmentGroupRow(no, name, color) {
  const tr = document.createElement('tr');
  tr.className = 'group-row';
  const th = document.createElement('th');
  th.colSpan = 4;
  th.append(fileBadge(no, color), name);
  tr.append(th);
  return tr;
}

// 経路ごとの内訳の、経路の行(No. は 1-1 など)
function segmentRow(segment, label) {
  const tr = document.createElement('tr');
  const hasTime = segment.startMs !== null;
  const timeCell = cell(hasTime ? formatJstRange(segment.startMs, segment.endMs) : '(記録時刻なし)', 'col-time');
  if (!hasTime) timeCell.classList.add('is-empty');
  tr.append(
    cell(label, 'col-no'),
    timeCell,
    cell(`${segment.points.length.toLocaleString('ja-JP')} 点`, 'col-num'),
    cell(formatDistance(segment.distanceM), 'col-num')
  );
  return tr;
}

// ===== 地図 =====
// 写真の撮影場所と、移動経路の記録地点をすべてマーカーで出す(KMZ・GPX のどちらかがあるときだけ)。
// fit が false なら、地図の中心・縮尺は今のままにする(KMZ・GPX をクリアしたとき)。
// 「内容を確かめる」が表示されてから呼ぶこと(隠れていると地図の大きさを測れない)。
// 地図を別ウインドウで開いていれば、そちらにも同じ内容を出す(map-window.js)
function showMap({ fit = true } = {}) {
  const hasData = state.photoRows.length > 0 || !!state.track;
  $('mapBlock').hidden = !hasData;
  // 地図を出している間は、画面の幅いっぱいに広げる(広がった分は GPX・地図の列に回る: style.css)
  document.body.classList.toggle('is-wide', hasData && canShowMap());
  if (!hasData) {
    sendToMapWindow(null); // 別ウインドウには、表示する地点が無いことを出す
    return;
  }
  if (!canShowMap()) {
    $('map').hidden = true;
    $('mapLegend').hidden = true;
    $('mapError').hidden = false;
    return;
  }

  const files = state.track?.files ?? [];
  // マーカーに出す日時は、KMZ・GPX に記録された日時(日本時間。秒まで)
  const timeLabel = (ms) => (Number.isFinite(ms) ? formatJstDateTimeSeconds(ms) : '');
  // 写真は、マーカーを押したときに出す縮小版(thumb)と、別ウインドウで元の写真を大きく表示するための元の写真(blob)も渡す
  const view = {
    photos: state.photoRows.map((p) => ({
      lat: p.lat, lng: p.lng, no: p.no, name: p.displayName, timeLabel: timeLabel(p.timeMs),
      thumb: p.thumb ? { blob: p.thumb.blob, width: p.thumb.width, height: p.thumb.height } : null,
      blob: p.blob
    })),
    tracks: files.flatMap((f, i) => f.segments.map((segment, j) => ({
      label: segmentNo(i, j),
      colorIndex: i,
      points: segment.points.map((p) => ({ lat: p.lat, lng: p.lng, timeLabel: timeLabel(p.timeMs) }))
    }))),
    // 凡例: 写真の撮影場所と、GPX(1, 2 …)ごとの記録地点の色
    legend: [
      ...(state.photoRows.length > 0
        ? [{ color: PHOTO_COLOR, text: `写真の撮影場所(${state.photoRows.length} 枚)`, kind: 'photo' }]
        : []),
      ...files.map((f, i) => ({ color: trackColor(i), text: `GPX ${fileNo(i)} の移動経路・記録地点`, kind: 'track' }))
    ]
  };
  // 地図の写真のマーカーを押したら、写真の一覧のその行に枠を付ける。出した縮小版を押したら、
  // 写真の一覧の「拡大表示」と同じく、元の写真を大きく表示する
  const selectOnMap = (p) => selectPhoto(p.no, 'map');
  const openPhotoOnMap = (p) => openPhoto(state.photoRows.find((row) => row.no === p.no));
  // 地図の右上に「別ウインドウで開く」を出す
  const windowLink = { href: MAP_WINDOW_PAGE, target: MAP_WINDOW_NAME, open: openMapWindow };
  updateMap($('map'), view, { fit, selectPhoto: selectOnMap, openPhoto: openPhotoOnMap, windowLink });
  renderLegend($('mapLegend'), view.legend);
  sendToMapWindow(view, { fit });
}

function objectUrl(blob) {
  const url = URL.createObjectURL(blob);
  state.urls.push(url);
  return url;
}

function cell(text, className) {
  const td = document.createElement('td');
  td.className = className;
  td.textContent = text;
  return td;
}

// 写真の一覧の1行: 1列目に No.(保存の通し番号)・名称・撮影日時・緯度, 経度を1行ずつ、
// 2列目に写真(押すと、地図でその写真の場所に縮小版を出す: selectPhoto)、3列目にボタン(拡大表示・写真をコピー・写真を保存)
function photoRow(photo) {
  const tr = document.createElement('tr');
  tr.dataset.no = String(photo.no);

  const infoCell = cell('', 'col-info');
  const line = (text, className) => {
    const div = document.createElement('div');
    div.className = className;
    div.textContent = text;
    infoCell.append(div);
    return div;
  };
  line(`No.${photo.no}`, 'info-no');
  const nameLine = line(photo.displayName, 'info-name');
  if (!photo.name) nameLine.classList.add('is-fallback');
  nameLine.title = `保存するファイル名: ${photo.fileName}`;
  const hasTime = Number.isFinite(photo.timeMs);
  const timeLine = line(hasTime ? formatJstDateTime(photo.timeMs) : '(撮影日時なし)', 'info-time');
  if (!hasTime) timeLine.classList.add('is-empty');
  line(`${photo.lat.toFixed(6)}, ${photo.lng.toFixed(6)}`, 'info-pos');

  const thumbCell = cell('', 'col-thumb');
  if (photo.thumb) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'thumb-button';
    button.title = '地図で、この写真の場所を表示';
    button.setAttribute('aria-label', `${photo.displayName} の場所を地図で表示`);
    const img = document.createElement('img');
    img.src = objectUrl(photo.thumb.blob);
    img.alt = '';
    img.width = photo.thumb.width;
    img.height = photo.thumb.height;
    button.append(img);
    button.addEventListener('click', () => selectPhoto(photo.no, 'list'));
    thumbCell.append(button);
  } else {
    const span = document.createElement('span');
    span.className = 'thumb-missing';
    span.textContent = 'この形式は表示できません';
    thumbCell.append(span);
  }

  // 表示できない形式(縮小版を作れなかった写真)は、拡大表示・コピーができない(保存はできる)
  const notViewable = photo.thumb ? '' : 'この形式は表示できません';
  const noCopy = canCopyImage
    ? notViewable
    : 'このブラウザでは写真をコピーできません(Chrome・Edge などで使えます)';
  const actions = document.createElement('div');
  actions.className = 'photo-actions';
  actions.append(
    actionButton(photo, '拡大表示', '元の写真を大きく表示します', notViewable, () => openPhoto(photo)),
    actionButton(photo, '写真をコピー', '元の大きさの写真をコピーします。そのまま貼り付けて使えます',
      noCopy, (button) => onCopyImage(button, photo.blob)),
    actionButton(photo, '写真を保存', `元の写真を「${photo.fileName}」としてダウンロードします`,
      '', () => saveImage(photo))
  );
  const actionsCell = cell('', 'col-actions');
  actionsCell.append(actions);

  tr.append(infoCell, thumbCell, actionsCell);
  return tr;
}

// 写真の一覧のボタン。disabledReason があれば押せなくし、理由をツールチップに出す
function actionButton(photo, label, title, disabledReason, onClick) {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'btn btn-mini';
  button.textContent = label;
  button.title = disabledReason || title;
  button.disabled = !!disabledReason;
  button.setAttribute('aria-label', `${label}: No.${photo.no} ${photo.displayName}`);
  button.addEventListener('click', () => onClick(button));
  return button;
}

// 画像をクリップボードにコピーし、結果をボタンの文字で知らせる。
// 押した操作の中でコピーを始める必要があるため、copyImage は待たずにすぐ呼ぶ
async function onCopyImage(button, blob) {
  try {
    await copyImage(blob);
    flashButton(button, 'コピーしました');
  } catch (err) {
    console.error(err);
    flashButton(button, 'コピー失敗', 2500);
  }
}

// 元の写真を大きく表示する
function openPhoto(photo) {
  photo.fullUrl ??= objectUrl(photo.blob);
  $('dialogImage').src = photo.fullUrl;
  $('dialogImage').alt = photo.displayName;
  const time = Number.isFinite(photo.timeMs) ? `  ${formatJstDateTime(photo.timeMs)}` : '';
  $('dialogCaption').textContent = `No.${photo.no}  ${photo.displayName}${time}`;
  $('photoDialog').showModal();
}

// ===== 写真を選ぶ(写真の一覧と地図を結ぶ) =====
// 写真の一覧のその写真(通し番号 no)の行に枠を付け(ほかの行の枠は外す)、地図(画面と、開いていれば別ウインドウ)に
// その写真の縮小版の吹き出しを出す。枠は、ほかの写真を選ぶか、KMZ を読み込み直す・クリアするまで残す
// (吹き出しを閉じても残す)。from はどこで選んだか:
//   'list'   … 写真の一覧の写真を押した。画面の地図が見えるようにスクロールする
//   'map'    … 画面の地図のマーカーを押した。スクロールはしない(地図が見えなくならないように)
//   'window' … 別ウインドウの地図のマーカーを押した。スクロールはしない
function selectPhoto(no, from) {
  for (const tr of $('photoRows').rows) tr.classList.toggle('is-selected', Number(tr.dataset.no) === no);
  if (from === 'list' && !$('mapBlock').hidden && canShowMap()) {
    const reduceMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
    $('map').scrollIntoView({ block: 'nearest', behavior: reduceMotion ? 'auto' : 'smooth' });
  }
  // マーカーを押した地図には、もう吹き出しが出ている(出ていれば、そのままにする)
  openPhotoPopup(no);
  selectInMapWindow(no);
}

// ===== 保存 =====
// 保存の設定(名前・保存するもの・保存先)と、保存される内容(フォルダ・ファイルの木)。
// 名前やチェックを変えるたびに、保存される内容を作り直す(renderSavePreview)。
// 保存ボタンを押したときは、画面に出している内容と同じ規則(naming.js の buildSavePlan)で保存する。

function setupSaveForm() {
  $('saveName').addEventListener('input', () => {
    state.save.edited = $('saveName').value !== state.save.defaultName;
    renderSavePreview();
  });
  $('btnResetName').addEventListener('click', () => setSaveName(state.save.defaultName));
  // 候補を押したら、欄の名前をその候補にする(そこから続けて直せる)
  $('nameCandidates').addEventListener('click', (e) => {
    const button = e.target.closest('button[data-name]');
    if (button && !button.disabled) setSaveName(button.dataset.name);
  });
  for (const key of ['includeKmz', 'includeGpx']) {
    $(key).addEventListener('change', () => {
      state.save[key] = $(key).checked;
      renderSavePreview();
    });
  }
}

function setSaveName(name) {
  $('saveName').value = name;
  state.save.edited = name !== state.save.defaultName;
  renderSavePreview();
}

// 読み込んだ内容が変わったら(読み込み・クリア)、名前の既定値と候補、写真のファイル名の一覧を作り直す。
// 利用者が名前を変えていなければ、名前も既定値に合わせる(変えていれば、そのまま残す)。
// KMZ・GPX をすべてクリアしたら「保存する」を隠し、名前を変えたことも忘れる
function refreshSaveSection() {
  const { album, track } = state;
  const hasData = !!album || !!track;
  $('outputSection').hidden = !hasData;
  showResult('');
  if (!hasData) {
    state.save.edited = false;
    return;
  }

  const nowMs = Date.now();
  state.save.defaultName = defaultSaveName(album, track, nowMs);
  if (!state.save.edited) $('saveName').value = state.save.defaultName;
  const label = document.createElement('span');
  label.className = 'name-candidates-label';
  label.textContent = '候補:';
  $('nameCandidates').replaceChildren(label, ...saveNameCandidates(album, track, nowMs).map((name) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn-chip';
    button.dataset.name = name;
    button.textContent = name;
    button.title = `名前を「${name}」にします`;
    return button;
  }));

  // 写真のファイル名の一覧(名前を変えても変わらないので、読み込んだ内容が変わったときだけ作る)
  const photos = state.photoRows;
  $('savePhotos').hidden = photos.length === 0;
  $('savePhotosSummary').textContent = `写真のファイル名を見る(${photos.length} 枚)`;
  $('savePhotoList').replaceChildren(...photos.map((p) => {
    const li = document.createElement('li');
    li.textContent = p.fileName;
    return li;
  }));
  renderSavePreview();
}

// 名前・チェック・保存先から、保存される内容(木)を作り直し、保存ボタンを押せるかを決める
function renderSavePreview() {
  const { album, track, busy } = state;
  if (!album && !track) return;
  const input = $('saveName');
  const check = checkSaveName(input.value);

  // 名前について知らせること(使えない文字は除き、長すぎれば切る。どちらも保存される内容の木に反映する)
  const notes = [];
  if (check.removed.length > 0) {
    const chars = [...new Set(check.removed.map((c) => (/[\u0000-\u001f\u007f]/.test(c) ? '制御文字' : `「${c}」`)))];
    notes.push(`${chars.join('')}は名前に使えないため除きます。`);
  }
  if (check.truncated) notes.push(`名前は ${MAX_NAME_LENGTH} 文字までにします(それより後は除きます)。`);
  if (!check.name) notes.push('名前を入力してください。');
  const note = $('saveNameNote');
  note.textContent = notes.join('');
  note.hidden = notes.length === 0;
  note.dataset.kind = check.name ? 'guide' : 'error';
  input.setAttribute('aria-invalid', String(!check.name));
  input.disabled = busy;
  $('btnResetName').disabled = busy || input.value === state.save.defaultName;
  for (const button of $('nameCandidates').querySelectorAll('button')) {
    button.disabled = busy;
    button.setAttribute('aria-pressed', String(button.dataset.name === input.value));
  }

  // 保存するもの: KMZ(Excel・写真・KMZ)と GPX。読み込んでいない種類は押せなくし、チェックも外して見せる
  setIncludeCheck('includeKmz', album ? album.files.length : 0, state.save.includeKmz, busy);
  setIncludeCheck('includeGpx', track ? track.files.length : 0, state.save.includeGpx, busy);

  // 保存される内容。名前が空のときは「（名前未入力）」と薄く出す(保存ボタンは押せない)
  const plan = buildSavePlan({
    album,
    track,
    name: check.name || '（名前未入力）',
    includeKmz: state.save.includeKmz,
    includeGpx: state.save.includeGpx
  });
  const noName = !check.name;
  const contents = [];
  if (plan.xlsxFileName) {
    contents.push(treeItem(plan.xlsxFileName, '写真の一覧(Excel)', 'file', [], noName));
    contents.push(treeItem(PHOTOS_DIR, `写真 ${plan.photos.length} 枚(元の画質のまま)`, 'folder'));
  }
  $('savePhotos').hidden = plan.photos.length === 0; // 写真を保存しないときは、写真のファイル名の一覧も出さない
  // KMZ・GPX は、1つなら元のファイルのまま、2つ以上なら1つにまとめたもの
  const sourceNote = (kind, count) => (count === 1
    ? `読み込んだ ${kind}(元のファイルのまま)`
    : `読み込んだ ${kind} ${count} ファイルを1つにまとめたもの`);
  if (plan.kmz) contents.push(treeItem(plan.kmz.fileName, sourceNote('KMZ', plan.kmz.count), 'file', [], noName));
  if (plan.gpx) contents.push(treeItem(plan.gpx.fileName, sourceNote('GPX', plan.gpx.count), 'file', [], noName));
  const zipName = `${plan.folderName}.zip`;
  if (canSaveToFolder) {
    const dir = currentSaveDir();
    $('saveTree').replaceChildren(treeItem(dir ? dir.name : '（保存先が未指定）', '保存先', 'folder', [
      treeItem(plan.folderName, '新しく作るフォルダ', 'folder', contents, noName)
    ], !dir));
    $('saveTreeNote').textContent =
      '「フォルダに保存」では、同じ名前のフォルダが既にあるときは上書きせず、「名前 (2)」のように番号を付けて作ります。' +
      `「ZIP でダウンロード」では、このフォルダの中身を「${zipName}」にまとめてダウンロードします。`;
  } else {
    $('saveTree').replaceChildren(treeItem('ダウンロード', 'ダウンロード先', 'folder', [
      treeItem(zipName, 'ZIP ファイル', 'zip', contents, noName)
    ]));
    $('saveTreeNote').textContent = '同じ名前のファイルが既にあるときは、ブラウザが番号を付けた名前にします。';
  }

  // 保存ボタンを押せないときは、理由をボタンの下に出す
  const loadedKinds = [album && 'KMZ', track && 'GPX'].filter(Boolean).join(' か ');
  const reason = noName
    ? '名前を入力すると保存できます。'
    : contents.length === 0
      ? `保存するものがありません。「保存するもの」の ${loadedKinds} にチェックを入れてください。`
      : '';
  $('saveBlocked').textContent = reason;
  $('saveBlocked').hidden = !reason;
  $('btnSaveFolder').disabled = busy || !!reason;
  $('btnSaveZip').disabled = busy || !!reason;
}

// 「保存するもの」のチェック。count はその種類の読み込んだファイルの数(0 なら押せなくし、チェックを外して見せる)
function setIncludeCheck(id, count, wanted, busy) {
  const box = $(id);
  box.checked = count > 0 && wanted;
  box.disabled = busy || count === 0;
  $(`${id}Count`).textContent = count > 0 ? `(${count} ファイル)` : '(読み込んでいません)';
}

// 木の1項目(名前・説明・中身)。kind は 'folder' | 'zip' | 'file'(アイコンの違い: style.css)。
// unset が true なら、まだ決まっていないもの(保存先が未指定・名前が空)として薄く出す
function treeItem(name, note, kind, children = [], unset = false) {
  const li = document.createElement('li');
  li.className = `is-${kind}`;
  if (unset) li.classList.add('is-unset');
  const nameEl = document.createElement('span');
  nameEl.className = 'tree-name';
  nameEl.textContent = name;
  li.append(nameEl);
  if (note) {
    const noteEl = document.createElement('span');
    noteEl.className = 'tree-note';
    noteEl.textContent = note;
    li.append(noteEl);
  }
  if (children.length > 0) {
    const ul = document.createElement('ul');
    ul.append(...children);
    li.append(ul);
  }
  return li;
}

// 保存先の表示(ブラウザからはフォルダの名前だけが分かり、場所(パス)は分からない)。
// 未指定なら「（フォルダ未指定）」を薄く表示し、「クリア」を押せないようにする。保存される内容の木の先頭も合わせる
function updateSaveDirLabel() {
  const dir = currentSaveDir();
  $('saveDirName').textContent = dir ? dir.name : '（フォルダ未指定）';
  $('saveDirName').classList.toggle('is-unset', !dir);
  $('btnClearSaveDir').disabled = state.busy || !dir;
  renderSavePreview();
}

async function onPickSaveDir() {
  if (state.busy) return;
  showResult('');
  try {
    const dir = await pickSaveDir();
    if (dir) showResult(`保存先を「${dir.name}」にしました。`, 'ok');
  } catch (err) {
    // まだ何も書き込んでいないので、失敗ではなく次の操作の案内として伝える
    showResult(err.message || '保存先を選べませんでした。もう一度「保存先を指定」を押してください。', 'guide');
  } finally {
    updateSaveDirLabel();
  }
}

async function onClearSaveDir() {
  if (state.busy || !currentSaveDir()) return;
  await clearSaveDir();
  updateSaveDirLabel();
  showResult('保存先をクリアしました。', 'ok');
}

async function onSaveFolder() {
  if (state.busy) return;
  const output = prepareOutput();
  if (!output) return;
  showResult('');
  // 保存先が未指定なら、失敗ではなく次の操作の案内として伝える
  if (!currentSaveDir()) {
    showResult('保存先が指定されていません。「保存先を指定」で、保存先のフォルダを指定してください。', 'guide');
    return;
  }

  // 書き込みの権限の確認は、押した直後に出す(Excel の作成より先に。間に時間のかかる処理を挟むと止められる)
  let parent;
  try {
    parent = await prepareSaveDir();
  } catch (err) {
    // 許可されなかった・保存先が見つからないなど。まだ何も書き込んでいないので案内として伝える
    showResult(err.message || '保存先を使えませんでした。「保存先を指定」で指定し直してください。', 'guide');
    return;
  } finally {
    updateSaveDirLabel();
  }

  const { folderName } = output.plan;
  await runOutput(output, async (entries, onFileDone) => {
    const name = await saveToFolder(parent, folderName, entries, onFileDone);
    const renamed = name !== folderName ? '同じ名前のフォルダが既にあったため、番号を付けた名前にしました。' : '';
    return `「${parent.name}」の中に「${name}」フォルダを作り、保存しました。${renamed}`;
  }, 'ファイルを保存中…');
}

async function onSaveZip() {
  if (state.busy) return;
  const output = prepareOutput();
  if (!output) return;
  showResult('');
  const { folderName, xlsxFileName } = output.plan;
  const zipName = `${folderName}.zip`;
  await runOutput(output, async (entries, onFileDone) => {
    const blob = await buildZipBlob(entries, onFileDone);
    downloadBlob(blob, zipName);
    const howToOpen = xlsxFileName
      ? `ZIP を展開してから「${xlsxFileName}」を開いてください(展開せずに開くと、Excel のリンクから写真を開けません)。`
      : '';
    return `「${zipName}」をダウンロードしました。${howToOpen}`;
  }, 'ZIP を作成中…');
}

// 保存する内容を、押した時点の名前・チェックで決める(画面の「保存される内容」と同じ規則)。
// 名前が空・保存するファイルが無いときは null(保存ボタンは押せなくしてあるが、念のため確かめる)
function prepareOutput() {
  const name = checkSaveName($('saveName').value).name;
  if (!name) return null;
  const plan = buildSavePlan({
    album: state.album,
    track: state.track,
    name,
    includeKmz: state.save.includeKmz,
    includeGpx: state.save.includeGpx
  });
  if (!plan.xlsxFileName && !plan.kmz && !plan.gpx) return null;
  return { album: state.album, track: state.track, nowMs: Date.now(), plan };
}

// 保存の共通処理: Excel を作り(KMZ を保存するとき)、KMZ・GPX をそろえ(2つ以上なら1つにまとめる)、
// 保存する内容(entries)を save に渡す。save(entries, onFileDone) は完了メッセージを返す。
// 進み具合は、保存したファイルの数で出す
async function runOutput(output, save, label) {
  setBusy(true);
  const { album, track, nowMs, plan } = output;
  try {
    const entries = [];
    if (plan.xlsxFileName) {
      showProgress('Excel を作成中…');
      entries.push({ path: plan.xlsxFileName, data: await buildWorkbookBlob(album, plan, nowMs) });
      entries.push(...plan.photos.map((p) => ({ path: `${PHOTOS_DIR}/${p.fileName}`, data: p.bytes })));
    }
    if (plan.kmz) {
      let data = album.files[0].kmzBytes;
      if (plan.kmz.count > 1) {
        showProgress('KMZ をまとめています…');
        data = await buildMergedKmz(plan.photos, plan.folderName);
      }
      entries.push({ path: plan.kmz.fileName, data });
    }
    if (plan.gpx) {
      const data = plan.gpx.count > 1 ? buildMergedGpx(track, plan.folderName) : track.files[0].source.bytes;
      entries.push({ path: plan.gpx.fileName, data });
    }

    const total = entries.length;
    let done = 0;
    showProgress(label, 0, total);
    const message = await save(entries, () => {
      done++;
      showProgress(label, done, total);
    });
    showResult(message, 'ok');
  } catch (err) {
    console.error(err);
    const partial = err?.partialFolder
      ? `「${err.partialFolder}」フォルダには一部のファイルだけが保存されています。`
      : '';
    showResult(`保存できませんでした(${err?.message || err})。${partial}`, 'error');
  } finally {
    hideProgress();
    setBusy(false);
  }
}

// ===== 表示の切り替え =====
function setBusy(busy) {
  state.busy = busy;
  for (const id of ['btnPickFile', 'btnClearAll', 'btnClearKmz', 'btnClearGpx', 'btnPickSaveDir']) {
    $(id).disabled = busy;
  }
  // 保存先の「クリア」と、保存の設定・保存ボタンは、ほかの条件でも押せなくするので別に扱う
  updateSaveDirLabel();
  $('dropZone').classList.toggle('is-disabled', busy);
  document.body.classList.toggle('is-busy', busy);
}

// 読み込みの結果を出す。lines は [{ message, kind }](1ファイル1行。kind は省くと 'info')。
// KMZ と GPX を一緒に読み込んだとき、片方だけ失敗しても、それぞれの結果が分かるよう行を分ける
function showStatus(lines) {
  $('loadStatus').replaceChildren(...lines.map(({ message, kind = 'info' }) => {
    const p = document.createElement('p');
    p.className = 'message';
    p.dataset.kind = kind;
    p.textContent = message;
    return p;
  }));
}

function showResult(message, kind = 'info') {
  const el = $('resultMessage');
  el.textContent = message;
  el.dataset.kind = kind;
  el.hidden = !message;
}

// total を省くと、件数を出さずに「処理中」の表示(進み具合の分からないバー)にする
function showProgress(label, done, total) {
  const bar = $('progressBar');
  $('progress').hidden = false;
  if (total === undefined) {
    bar.removeAttribute('value');
    $('progressText').textContent = label;
    return;
  }
  bar.max = Math.max(1, total);
  bar.value = done;
  $('progressText').textContent = `${label} ${done} / ${total}`;
}

function hideProgress() {
  $('progress').hidden = true;
}
