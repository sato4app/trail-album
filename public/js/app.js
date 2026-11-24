// Trail Album のエントリポイント(画面の制御)
// KMZ を読み込んでプレビューを表示し、写真の一覧の Excel(KMZ と同じ名前の .xlsx)と
// 元の画質の写真を、フォルダへ直接保存するか、ZIP でダウンロードする。処理はすべてブラウザの中で行う。
// KMZ・GPX はいくつでも読み込め、後から読み込んだものは今の内容に追加する(merge.js)。
// GPX(移動経路)は、今は内容の表示だけで、保存する内容には含めない。
import { readKmz, KmzError } from './kmz-reader.js';
import { readGpx, GpxError } from './gpx-reader.js';
import { excludeLoadedPhotos, buildAlbum, addSegments, buildTrack } from './merge.js';
import { createThumbnail } from './thumbnail.js';
import { canCopyImage, copyImage, saveImage, flashButton } from './photo-actions.js';
import { canShowMap, updateMap, PHOTO_COLOR, trackColor } from './map-view.js';
import { buildOutputPlan, formatJstDateTime, formatJstRange } from './naming.js';
import { buildWorkbookBlob, PHOTOS_DIR } from './excel-writer.js';
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
  track: null,     // segments をまとめたもの(merge.js の buildTrack の結果。表示だけに使う)
  photoRows: [],   // 写真の一覧に出している写真(No. の a-1 … を付けたもの。地図にも使う)
  urls: [],        // プレビューで作った画像の URL(写真の一覧を作り直すときに解放する)
  busy: false      // 読み込み中・保存中
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
  $('btnClearKmz').addEventListener('click', onClearKmz);
  $('btnClearGpx').addEventListener('click', onClearGpx);

  $('btnSaveFolder').addEventListener('click', onSaveFolder);
  $('btnPickSaveDir').addEventListener('click', onPickSaveDir);
  $('btnClearSaveDir').addEventListener('click', onClearSaveDir);
  $('btnSaveZip').addEventListener('click', onSaveZip);

  const dialog = $('photoDialog');
  // 画像の外側(背景)を押したら閉じる
  dialog.addEventListener('click', (e) => {
    if (e.target === dialog) dialog.close();
  });
  dialog.addEventListener('close', () => {
    $('dialogImage').removeAttribute('src');
  });

  // 保存の確認: 背景を押したらキャンセルとして閉じる
  const confirmDialog = $('confirmDialog');
  confirmDialog.addEventListener('click', (e) => {
    if (e.target === confirmDialog) confirmDialog.close('cancel');
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
  // 保存の確認を出している間は読み込まない(確認した内容と違うものを保存しないように)
  if (files.length === 0 || state.busy || $('confirmDialog').open) return;
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

// 読み込んだ KMZ を画面から外す(ファイルは消えない)
function onClearKmz() {
  if (state.busy || state.kmzFiles.length === 0) return;
  state.kmzFiles = [];
  showKmzPreview();
  showStatus([{ message: '読み込んだ KMZ をクリアしました。', kind: 'ok' }]);
}

// 読み込んだ GPX を画面から外す(ファイルは消えない)
function onClearGpx() {
  if (state.busy || state.segments.length === 0) return;
  state.segments = [];
  showGpxPreview();
  showStatus([{ message: '読み込んだ GPX をクリアしました。', kind: 'ok' }]);
}

// 左側(KMZ)の表示を、読み込んだ KMZ(state.kmzFiles)から作り直す
function showKmzPreview() {
  for (const url of state.urls) URL.revokeObjectURL(url);
  state.urls = [];
  const album = buildAlbum(state.kmzFiles);
  state.album = album;

  $('kmzEmpty').hidden = !!album;
  $('kmzPreview').hidden = !album;
  $('btnClearKmz').hidden = !album;
  $('outputSection').hidden = !album; // 保存するのは KMZ の内容なので、KMZ を読み込んでいるときだけ出す
  showResult('');
  updatePreviewSection();

  state.photoRows = album ? labelPhotos(album, buildOutputPlan(album, Date.now())) : [];
  if (album) {
    $('sumFile').replaceChildren(...album.files.map((f, i) => fileLine(fileLetter(i), f.fileName, PHOTO_COLOR)));
    $('sumCount').textContent = `${album.photos.length} 枚`;
  }
  $('photoRows').replaceChildren(...state.photoRows.map(photoRow));
  showMap();
}

// 写真を KMZ ごと(a, b …)・撮影日時の順に並べ、No.(a-1, a-2 …)を label に付ける。
// plan(buildOutputPlan の結果。保存するファイル名などを持つ)の photos は album.photos と同じ順で対応する
function labelPhotos(album, plan) {
  const groups = new Map(album.files.map((file, i) => [file, { letter: fileLetter(i), photos: [] }]));
  const fileOf = new Map(album.files.flatMap((file) => file.photos.map((photo) => [photo, file])));
  album.photos.forEach((photo, i) => {
    const group = groups.get(fileOf.get(photo));
    group.photos.push({ ...plan.photos[i], label: `${group.letter}-${group.photos.length + 1}` });
  });
  return [...groups.values()].flatMap((g) => g.photos);
}

// 右側(GPX)の表示を、読み込んだ経路(state.segments)から作り直す
function showGpxPreview() {
  const track = buildTrack(state.segments);
  state.track = track;

  $('gpxEmpty').hidden = !!track;
  $('gpxPreview').hidden = !track;
  $('btnClearGpx').hidden = !track;
  updatePreviewSection();

  if (track) {
    const routes = track.segmentCount > 1 ? `(${track.segmentCount} 経路)` : '';
    $('gpxFile').replaceChildren(
      ...track.files.map((f, i) => fileLine(gpxLetter(i), f.source.fileName, trackColor(i)))
    );
    $('gpxPoints').textContent = `${track.pointCount.toLocaleString('ja-JP')} 点${routes}`;
    $('gpxDistance').textContent = formatDistance(track.distanceM);
  }
  // 経路ごとの内訳: GPX ごとに、ファイル名の行と、その経路(A-1, A-2 …)の行を並べる
  $('segRows').replaceChildren(...(track?.files ?? []).flatMap((f, i) => [
    segmentGroupRow(gpxLetter(i), f.source.fileName, trackColor(i)),
    ...f.segments.map((segment, j) => segmentRow(segment, `${gpxLetter(i)}-${j + 1}`))
  ]));
  showMap();
}

// KMZ・GPX のどちらかを読み込んでいるときだけ「内容を確かめる」を出す
function updatePreviewSection() {
  $('previewSection').hidden = !state.album && !state.track;
}

// KMZ の記号: a, b, … z, aa, ab …(GPX は大文字: gpxLetter)
function fileLetter(index) {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(97 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

const gpxLetter = (index) => fileLetter(index).toUpperCase();

// ファイル名の行(左に a, b … / A, B … の記号。記号の色は地図のマーカーの色)
function fileLine(letter, name, color) {
  const div = document.createElement('div');
  div.className = 'file-line';
  div.append(fileBadge(letter, color), name);
  return div;
}

function fileBadge(letter, color) {
  const badge = document.createElement('span');
  badge.className = 'file-badge';
  badge.style.backgroundColor = color;
  badge.textContent = letter;
  return badge;
}

// 8.4 km(minoh-hiking の統計表と同じく小数点以下1位)
function formatDistance(meters) {
  return `${(meters / 1000).toFixed(1)} km`;
}

// 経路ごとの内訳の、GPX のファイル名の行
function segmentGroupRow(letter, name, color) {
  const tr = document.createElement('tr');
  tr.className = 'group-row';
  const th = document.createElement('th');
  th.colSpan = 4;
  th.append(fileBadge(letter, color), name);
  tr.append(th);
  return tr;
}

// 経路ごとの内訳の、経路の行(No. は A-1 など)
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
// 「内容を確かめる」が表示されてから呼ぶこと(隠れていると地図の大きさを測れない)
function showMap() {
  const hasData = state.photoRows.length > 0 || !!state.track;
  $('mapBlock').hidden = !hasData;
  if (!hasData) return;
  if (!canShowMap()) {
    $('map').hidden = true;
    $('mapLegend').hidden = true;
    $('mapError').hidden = false;
    return;
  }

  const files = state.track?.files ?? [];
  updateMap($('map'), {
    photos: state.photoRows.map((p) => ({ lat: p.lat, lng: p.lng, label: p.label, name: p.displayName })),
    tracks: files.flatMap((f, i) => f.segments.map((segment, j) => ({
      label: `${gpxLetter(i)}-${j + 1}`,
      colorIndex: i,
      points: segment.points.map((p) => ({
        lat: p.lat,
        lng: p.lng,
        timeLabel: p.timeMs !== null ? formatJstDateTime(p.timeMs) : ''
      }))
    })))
  });

  // 凡例: 写真の撮影場所と、GPX(A, B …)ごとの記録地点の色
  const legend = [];
  if (state.photoRows.length > 0) legend.push(legendItem(PHOTO_COLOR, `写真の撮影場所(${state.photoRows.length} 枚)`, 'is-photo'));
  files.forEach((f, i) => legend.push(legendItem(trackColor(i), `${gpxLetter(i)} の記録地点`)));
  $('mapLegend').replaceChildren(...legend);
}

function legendItem(color, text, className = '') {
  const span = document.createElement('span');
  span.className = `legend-item ${className}`.trim();
  const dot = document.createElement('span');
  dot.className = 'legend-dot';
  dot.style.backgroundColor = color;
  span.append(dot, text);
  return span;
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

// 写真の一覧の1行: 1列目に No.(a-1 など)・名称・撮影日時・緯度, 経度を1行ずつ、2列目に写真、
// 3列目にボタン(拡大表示・縮小版をコピー・画像をコピー・画像を保存)
function photoRow(photo) {
  const tr = document.createElement('tr');

  const infoCell = cell('', 'col-info');
  const line = (text, className) => {
    const div = document.createElement('div');
    div.className = className;
    div.textContent = text;
    infoCell.append(div);
    return div;
  };
  line(`No.${photo.label}`, 'info-no');
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
    button.title = '元の写真を大きく表示';
    button.setAttribute('aria-label', `${photo.displayName} を大きく表示`);
    const img = document.createElement('img');
    img.src = objectUrl(photo.thumb.blob);
    img.alt = '';
    img.width = photo.thumb.width;
    img.height = photo.thumb.height;
    button.append(img);
    button.addEventListener('click', () => openPhoto(photo));
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
    : 'このブラウザでは画像をコピーできません(Chrome・Edge などで使えます)';
  const actions = document.createElement('div');
  actions.className = 'photo-actions';
  actions.append(
    actionButton(photo, '拡大表示', '元の写真を大きく表示します', notViewable, () => openPhoto(photo)),
    actionButton(photo, '縮小版をコピー', '縮小版(長辺 約400px)の画像をコピーします。そのまま貼り付けて使えます',
      noCopy, (button) => onCopyImage(button, photo.thumb.blob)),
    actionButton(photo, '画像をコピー', '元の大きさの画像をコピーします。そのまま貼り付けて使えます',
      noCopy, (button) => onCopyImage(button, photo.blob)),
    actionButton(photo, '画像を保存', `元の写真を「${photo.fileName}」としてダウンロードします`,
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
  button.setAttribute('aria-label', `${label}: No.${photo.label} ${photo.displayName}`);
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
  $('dialogCaption').textContent = `No.${photo.label}  ${photo.displayName}${time}`;
  $('photoDialog').showModal();
}

// ===== 保存 =====
// 保存先の表示(ブラウザからはフォルダの名前だけが分かり、場所(パス)は分からない)。
// 未指定なら「（フォルダ未指定）」を薄く表示し、「クリア」を押せないようにする
function updateSaveDirLabel() {
  const dir = currentSaveDir();
  $('saveDirName').textContent = dir ? dir.name : '（フォルダ未指定）';
  $('saveDirName').classList.toggle('is-unset', !dir);
  $('btnClearSaveDir').disabled = state.busy || !dir;
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
  if (!state.album || state.busy) return;
  showResult('');
  // 保存先が未指定なら、失敗ではなく次の操作の案内として伝える
  const dir = currentSaveDir();
  if (!dir) {
    showResult('保存先が指定されていません。「保存先を指定」で、保存先のフォルダを指定してください。', 'guide');
    return;
  }

  const output = prepareOutput();
  const confirmed = await confirmSave({
    title: 'フォルダに保存します',
    lead: `保存先「${dir.name}」の中に、次のフォルダを作って保存します。`,
    parent: { name: dir.name, note: '保存先' },
    container: { name: output.plan.folderName, note: '新しく作るフォルダ', kind: 'folder' },
    output,
    note: '同じ名前のフォルダが既にあるときは、上書きせず「名前 (2)」のように番号を付けて作ります。',
    okLabel: '保存する'
  });
  if (!confirmed) return;

  // 書き込みの権限の確認は、「保存する」を押した直後に出す(Excel の作成より先に)
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

  await runOutput(output, async (plan, entries, onPhotoDone) => {
    const name = await saveToFolder(parent, plan.folderName, entries, (path) => {
      if (path.startsWith(`${PHOTOS_DIR}/`)) onPhotoDone();
    });
    const renamed = name !== plan.folderName ? '同じ名前のフォルダが既にあったため、番号を付けた名前にしました。' : '';
    return `「${parent.name}」の中に「${name}」フォルダを作り、保存しました。${renamed}`;
  }, '写真を保存中…');
}

async function onSaveZip() {
  if (!state.album || state.busy) return;
  showResult('');
  const output = prepareOutput();
  const zipName = `${output.plan.folderName}.zip`;
  const confirmed = await confirmSave({
    title: 'ZIP でダウンロードします',
    lead: '次の ZIP ファイルを「ダウンロード」フォルダにダウンロードします' +
      '(ブラウザの設定でダウンロード先を変えている場合は、その場所になります)。',
    parent: { name: 'ダウンロード', note: 'ダウンロード先' },
    container: { name: zipName, note: 'ZIP ファイル', kind: 'zip' },
    output,
    note: '同じ名前のファイルが既にあるときは、ブラウザが番号を付けた名前にします。' +
      'ZIP は展開してから Excel ファイルを開いてください。',
    okLabel: 'ダウンロードする'
  });
  if (!confirmed) return;

  await runOutput(output, async (plan, entries, onPhotoDone) => {
    const blob = await buildZipBlob(entries, (path) => {
      if (path.startsWith(`${PHOTOS_DIR}/`)) onPhotoDone();
    });
    downloadBlob(blob, zipName);
    return `「${zipName}」をダウンロードしました。ZIP を展開してから「${plan.xlsxFileName}」を開いてください` +
      '(展開せずに開くと、Excel のリンクから写真を開けません)。';
  }, 'ZIP を作成中…');
}

// 保存する内容を決める。確認の画面に出した名前のとおりに保存するよう、ここで決めたものを最後まで使う
function prepareOutput() {
  const nowMs = Date.now();
  return {
    album: state.album,
    nowMs,
    plan: buildOutputPlan(state.album, nowMs),
    includeKmz: $('includeKmz').checked
  };
}

// ===== 保存の確認 =====
// どこに何が保存されるかを、フォルダ・ファイルの木で見せて確かめる。
// 「保存する」(okLabel)なら true、キャンセル・Esc・背景を押したときは false を返す。
//   parent:    保存先 { name, note }
//   container: 保存先の中に作るもの { name, note, kind: 'folder' | 'zip' }
//   output:    prepareOutput の結果(中身のファイル名に使う)
function confirmSave({ title, lead, parent, container, output, note, okLabel }) {
  const { plan, includeKmz } = output;
  const contents = [
    treeItem(plan.xlsxFileName, '写真の一覧(Excel)', 'file'),
    treeItem(PHOTOS_DIR, `写真 ${plan.photos.length} 枚(元の画質のまま)`, 'folder')
  ];
  if (includeKmz) contents.push(...plan.kmzFileNames.map((name) => treeItem(name, '読み込んだ KMZ', 'file')));
  const tree = treeItem(parent.name, parent.note, 'folder', [
    treeItem(container.name, container.note, container.kind, contents)
  ]);

  $('confirmTitle').textContent = title;
  $('confirmLead').textContent = lead;
  $('confirmTree').replaceChildren(tree);
  $('confirmPhotosSummary').textContent = `写真のファイル名を見る(${plan.photos.length} 枚)`;
  $('confirmPhotoList').replaceChildren(...plan.photos.map((p) => {
    const li = document.createElement('li');
    li.textContent = p.fileName;
    return li;
  }));
  $('confirmPhotos').open = false;
  $('confirmNote').textContent = note;
  $('confirmOk').textContent = okLabel;

  const dialog = $('confirmDialog');
  dialog.returnValue = '';
  dialog.showModal();
  return new Promise((resolve) => {
    dialog.addEventListener('close', () => resolve(dialog.returnValue === 'ok'), { once: true });
  });
}

// 木の1項目(名前・説明・中身)。kind は 'folder' | 'zip' | 'file'(アイコンの違い: style.css)
function treeItem(name, note, kind, children = []) {
  const li = document.createElement('li');
  li.className = `is-${kind}`;
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

// 保存の共通処理: Excel を作り、保存する内容(entries)をそろえて save に渡す。
// output は prepareOutput の結果。save(plan, entries, onPhotoDone) は完了メッセージを返す。
async function runOutput(output, save, photoLabel) {
  setBusy(true);
  const { album, nowMs, plan, includeKmz } = output;
  const total = plan.photos.length;
  let done = 0;
  try {
    showProgress('Excel を作成中…');
    const xlsx = await buildWorkbookBlob(album, plan, nowMs);

    const entries = [
      { path: plan.xlsxFileName, data: xlsx },
      ...plan.photos.map((p) => ({ path: `${PHOTOS_DIR}/${p.fileName}`, data: p.bytes }))
    ];
    if (includeKmz) {
      album.files.forEach((file, i) => entries.push({ path: plan.kmzFileNames[i], data: file.kmzBytes }));
    }

    showProgress(photoLabel, 0, total);
    const message = await save(plan, entries, () => {
      done++;
      showProgress(photoLabel, done, total);
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
  for (const id of ['btnPickFile', 'btnClearKmz', 'btnClearGpx', 'btnSaveFolder', 'btnPickSaveDir', 'btnSaveZip', 'includeKmz']) {
    $(id).disabled = busy;
  }
  updateSaveDirLabel(); // 「クリア」は保存先が未指定のときも押せないので別に扱う
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
