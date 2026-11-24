// Trail Album のエントリポイント(画面の制御)
// KMZ を読み込んでプレビューを表示し、写真の一覧の Excel(KMZ と同じ名前の .xlsx)と
// 元の画質の写真を、フォルダへ直接保存するか、ZIP でダウンロードする。処理はすべてブラウザの中で行う。
// GPX(移動経路)も読み込めるが、今は内容の表示だけで、保存する内容には含めない。
import { readKmz, KmzError } from './kmz-reader.js';
import { readGpx, GpxError } from './gpx-reader.js';
import { createThumbnail } from './thumbnail.js';
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
  album: null,     // 読み込んだ KMZ(kmz-reader.js の readKmz の結果)
  track: null,     // 読み込んだ GPX(gpx-reader.js の readGpx の結果。表示だけに使う)
  urls: [],        // プレビューで作った画像の URL(次の KMZ を読み込んだときに解放する)
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
// 選んだ(置いた)ファイルから、KMZ と GPX をそれぞれ1つずつ読み込む。
// KMZ と GPX の表示は別々に置き換え、片方だけ読み込んだときは、もう片方の表示をそのまま残す。
async function loadFiles(files) {
  // 保存の確認を出している間は読み込まない(確認した内容と違うものを保存しないように)
  if (files.length === 0 || state.busy || $('confirmDialog').open) return;
  setBusy(true);
  const results = []; // 読み込みの結果(1ファイル1行)
  const progress = (message) => showStatus([...results, { message }]);
  try {
    const { kmz, gpx, unknown } = await sortFiles(files);
    if (kmz.length > 0) results.push(await loadKmz(kmz[0], kmz.length > 1, progress));
    if (gpx.length > 0) results.push(await loadGpx(gpx[0], gpx.length > 1, progress));
    if (unknown.length > 0) {
      const names = unknown.map((f) => `「${f.name}」`).join('');
      results.push({ message: `${names}は KMZ, GPX ファイルではないため、読み込みませんでした。`, kind: 'error' });
    }
  } catch (err) {
    console.error(err);
    results.push({ message: `読み込みに失敗しました(${err.message || err})。`, kind: 'error' });
  } finally {
    showStatus(results);
    setBusy(false);
  }
}

// ファイルを KMZ・GPX・それ以外に分ける(並びは選んだ順のまま)
async function sortFiles(files) {
  const sorted = { kmz: [], gpx: [], unknown: [] };
  for (const file of files) sorted[(await fileKind(file)) || 'unknown'].push(file);
  return sorted;
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

// KMZ を読み込み、左側(KMZ)の表示を置き換える。結果の1行({ message, kind })を返す。
// 読み込めなかったときは、前の表示をそのまま残す
async function loadKmz(file, hadMultiple, progress) {
  progress(`「${file.name}」を読み込み中…`);
  try {
    const album = await readKmz(file);
    const total = album.photos.length;
    for (let i = 0; i < total; i++) {
      progress(`「${file.name}」のサムネイルを作成中… ${i + 1} / ${total}`);
      const photo = album.photos[i];
      photo.blob = new Blob([photo.bytes], { type: photo.type });
      photo.thumb = await createThumbnail(photo.blob);
    }
    showKmzPreview(album);
    const note = hadMultiple ? '(KMZ ファイルが複数あったため、最初の1つだけを読み込みました)' : '';
    return { message: `「${file.name}」を読み込みました。写真は ${total} 枚です。${note}`, kind: 'ok' };
  } catch (err) {
    console.error(err);
    const message = err instanceof KmzError ? err.message : `読み込みに失敗しました(${err.message || err})。`;
    return { message: `「${file.name}」: ${message}`, kind: 'error' };
  }
}

// GPX を読み込み、右側(GPX)の表示を置き換える。結果の1行({ message, kind })を返す。
// 読み込めなかったときは、前の表示をそのまま残す
async function loadGpx(file, hadMultiple, progress) {
  progress(`「${file.name}」を読み込み中…`);
  try {
    const track = await readGpx(file);
    showGpxPreview(track);
    const note = hadMultiple ? '(GPX ファイルが複数あったため、最初の1つだけを読み込みました)' : '';
    return {
      message: `「${file.name}」を読み込みました。記録地点は ${pointsText(track)}です。${note}`,
      kind: 'ok'
    };
  } catch (err) {
    console.error(err);
    const message = err instanceof GpxError ? err.message : `読み込みに失敗しました(${err.message || err})。`;
    return { message: `「${file.name}」: ${message}`, kind: 'error' };
  }
}

// 左側(KMZ)の表示を置き換える
function showKmzPreview(album) {
  for (const url of state.urls) URL.revokeObjectURL(url);
  state.urls = [];
  state.album = album;

  const plan = buildOutputPlan(album, Date.now());
  $('sumFile').textContent = album.fileName;
  $('sumTitle').textContent = album.title || '(なし)';
  $('sumCount').textContent = `${plan.photos.length} 枚`;
  $('sumFolder').textContent = plan.folderName;

  $('photoRows').replaceChildren(...plan.photos.map(photoRow));
  $('kmzEmpty').hidden = true;
  $('kmzPreview').hidden = false;
  $('previewSection').hidden = false;
  $('outputSection').hidden = false; // 保存するのは KMZ の内容なので、KMZ を読み込んでから出す
  showResult('');
}

// 右側(GPX)の表示を置き換える
function showGpxPreview(track) {
  state.track = track;
  const hasTime = track.startMs !== null;
  $('gpxFile').textContent = track.fileName;
  $('gpxTitle').textContent = track.title || '(なし)';
  $('gpxPoints').textContent = pointsText(track);
  $('gpxDistance').textContent = formatDistance(track.distanceM);
  $('gpxTime').textContent = hasTime ? formatJstRange(track.startMs, track.endMs) : '(記録時刻なし)';
  $('gpxDuration').textContent = hasTime ? formatDuration(track.endMs - track.startMs) : '(なし)';

  $('segBlock').hidden = track.segments.length < 2;
  $('segRows').replaceChildren(...track.segments.map(segmentRow));
  $('gpxEmpty').hidden = true;
  $('gpxPreview').hidden = false;
  $('previewSection').hidden = false;
}

// 1,234 点(2 経路)。経路が1本なら経路の数は省く
function pointsText(track) {
  const routes = track.segments.length > 1 ? `(${track.segments.length} 経路)` : '';
  return `${track.pointCount.toLocaleString('ja-JP')} 点${routes}`;
}

// 8.4 km(minoh-hiking の統計表と同じく小数点以下1位)
function formatDistance(meters) {
  return `${(meters / 1000).toFixed(1)} km`;
}

// 2時間13分 / 45分
function formatDuration(ms) {
  const minutes = Math.round(ms / 60000);
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return h > 0 ? `${h}時間${m}分` : `${m}分`;
}

function segmentRow(segment, i) {
  const tr = document.createElement('tr');
  const hasTime = segment.startMs !== null;
  const timeCell = cell(hasTime ? formatJstRange(segment.startMs, segment.endMs) : '(なし)', 'col-time');
  if (!hasTime) timeCell.classList.add('is-empty');
  tr.append(
    cell(String(i + 1), 'col-no'),
    timeCell,
    cell(`${segment.points.length.toLocaleString('ja-JP')} 点`, 'col-num'),
    cell(formatDistance(segment.distanceM), 'col-num')
  );
  return tr;
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

function photoRow(photo) {
  const tr = document.createElement('tr');

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

  const nameCell = cell(photo.displayName, 'col-name');
  if (!photo.name) nameCell.classList.add('is-fallback');
  nameCell.dataset.no = String(photo.no); // スマートフォン表示で名称の前に No. を出す(style.css)
  nameCell.title = `保存するファイル名: ${photo.fileName}`;

  const time = Number.isFinite(photo.timeMs) ? formatJstDateTime(photo.timeMs) : '(なし)';
  const timeCell = cell(time, 'col-time');
  if (!Number.isFinite(photo.timeMs)) timeCell.classList.add('is-empty');

  tr.append(
    cell(String(photo.no), 'col-no'),
    thumbCell,
    nameCell,
    timeCell,
    cell(`${photo.lat.toFixed(6)}, ${photo.lng.toFixed(6)}`, 'col-pos')
  );
  return tr;
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
  if (includeKmz) contents.push(treeItem(plan.kmzFileName, '読み込んだ KMZ', 'file'));
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
    if (includeKmz) entries.push({ path: plan.kmzFileName, data: album.kmzBytes });

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
  for (const id of ['btnPickFile', 'btnSaveFolder', 'btnPickSaveDir', 'btnSaveZip', 'includeKmz']) {
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
