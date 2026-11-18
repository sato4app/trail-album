// Trail Album のエントリポイント(画面の制御)
// KMZ を読み込んでプレビューを表示し、写真の一覧の Excel(KMZ と同じ名前の .xlsx)と
// 元の画質の写真を、フォルダへ直接保存するか、ZIP でダウンロードする。処理はすべてブラウザの中で行う。
import { readKmz, KmzError } from './kmz-reader.js';
import { createThumbnail } from './thumbnail.js';
import { buildOutputPlan, formatJstDateTime } from './naming.js';
import { buildWorkbookBlob, PHOTOS_DIR } from './excel-writer.js';
import {
  canSaveToFolder,
  loadSaveDir,
  currentSaveDir,
  pickSaveDir,
  prepareSaveDir,
  saveToFolder,
  buildZipBlob,
  downloadBlob
} from './saver.js';

const $ = (id) => document.getElementById(id);

const state = {
  album: null,     // 読み込んだ KMZ(kmz-reader.js の readKmz の結果)
  urls: [],        // プレビューで作った画像の URL(次の KMZ を読み込んだときに解放する)
  busy: false      // 読み込み中・保存中
};

init();

function init() {
  // CDN のライブラリが読み込めていなければ、操作できないことを伝える
  if (!globalThis.JSZip || !globalThis.ExcelJS) {
    showStatus(
      '必要なライブラリを読み込めませんでした。インターネットに接続していることを確かめてから、ページを再読み込みしてください。',
      'error'
    );
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
    const file = e.target.files?.[0];
    e.target.value = ''; // 同じファイルを続けて選んでも change が起きるように
    if (file) loadFile(file);
  });
  setupDragAndDrop();

  $('btnSaveFolder').addEventListener('click', onSaveFolder);
  $('btnChangeSaveDir').addEventListener('click', onChangeSaveDir);
  $('btnSaveZip').addEventListener('click', onSaveZip);

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
// ブラウザが KMZ をダウンロードしてしまうため)。
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
    const files = Array.from(e.dataTransfer.files);
    if (files.length === 0) return;
    loadFile(files[0], files.length > 1);
  });
}

// ===== KMZ の読み込み =====
async function loadFile(file, hadMultiple = false) {
  if (state.busy) return;
  setBusy(true);
  showStatus(`「${file.name}」を読み込み中…`);
  try {
    const album = await readKmz(file);
    const total = album.photos.length;
    for (let i = 0; i < total; i++) {
      showStatus(`サムネイルを作成中… ${i + 1} / ${total}`);
      const photo = album.photos[i];
      photo.blob = new Blob([photo.bytes], { type: photo.type });
      photo.thumb = await createThumbnail(photo.blob);
    }
    replacePreview(album);
    const note = hadMultiple ? '(複数のファイルが置かれたため、最初の1つだけを読み込みました)' : '';
    showStatus(`「${file.name}」を読み込みました。写真は ${total} 枚です。${note}`, 'ok');
  } catch (err) {
    // 読み込めなかったときは、前のプレビューをそのまま残す
    const message = err instanceof KmzError ? err.message : `読み込みに失敗しました(${err.message || err})。`;
    showStatus(`「${file.name}」: ${message}`, 'error');
    console.error(err);
  } finally {
    setBusy(false);
  }
}

// 前のプレビューを置き換える
function replacePreview(album) {
  for (const url of state.urls) URL.revokeObjectURL(url);
  state.urls = [];
  state.album = album;

  const plan = buildOutputPlan(album, Date.now());
  $('sumFile').textContent = album.fileName;
  $('sumTitle').textContent = album.title || '(なし)';
  $('sumCount').textContent = `${plan.photos.length} 枚`;
  $('sumFolder').textContent = plan.folderName;

  $('photoRows').replaceChildren(...plan.photos.map(photoRow));
  $('previewSection').hidden = false;
  $('outputSection').hidden = false;
  showResult('');
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
// 保存先の表示(ブラウザからはフォルダの名前だけが分かり、場所(パス)は分からない)
function updateSaveDirLabel() {
  const dir = currentSaveDir();
  $('saveDirName').textContent = dir ? dir.name : '未設定(初めて保存するときに選びます)';
  $('saveDirName').classList.toggle('is-unset', !dir);
}

async function onChangeSaveDir() {
  if (state.busy) return;
  showResult('');
  try {
    const dir = await pickSaveDir();
    if (dir) showResult(`保存先を「${dir.name}」にしました。`, 'ok');
  } catch (err) {
    showResult(err.message || '保存先を選べませんでした。', 'error');
  } finally {
    updateSaveDirLabel();
  }
}

async function onSaveFolder() {
  if (!state.album || state.busy) return;
  showResult('');
  // 権限の確認や保存先を選ぶ画面は、ボタンを押した直後に出す(Excel の作成より先に)
  let parent;
  try {
    parent = await prepareSaveDir();
  } catch (err) {
    showResult(err.message || '保存先を使えませんでした。', 'error');
    return;
  } finally {
    updateSaveDirLabel();
  }
  if (!parent) return; // 保存先を選ぶ画面でキャンセルされた

  await runOutput(async (plan, entries, onPhotoDone) => {
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
  await runOutput(async (plan, entries, onPhotoDone) => {
    const blob = await buildZipBlob(entries, (path) => {
      if (path.startsWith(`${PHOTOS_DIR}/`)) onPhotoDone();
    });
    const zipName = `${plan.folderName}.zip`;
    downloadBlob(blob, zipName);
    return `「${zipName}」をダウンロードしました。ZIP を展開してから「${plan.xlsxFileName}」を開いてください` +
      '(展開せずに開くと、Excel のリンクから写真を開けません)。';
  }, 'ZIP を作成中…');
}

// 保存の共通処理: Excel を作り、保存する内容(entries)をそろえて save に渡す。
// save(plan, entries, onPhotoDone) は完了メッセージを返す。
async function runOutput(save, photoLabel) {
  setBusy(true);
  const nowMs = Date.now();
  const plan = buildOutputPlan(state.album, nowMs);
  const total = plan.photos.length;
  let done = 0;
  try {
    showProgress('Excel を作成中…');
    const xlsx = await buildWorkbookBlob(state.album, plan, nowMs);

    const entries = [
      { path: plan.xlsxFileName, data: xlsx },
      ...plan.photos.map((p) => ({ path: `${PHOTOS_DIR}/${p.fileName}`, data: p.bytes }))
    ];
    if ($('includeKmz').checked) entries.push({ path: plan.kmzFileName, data: state.album.kmzBytes });

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
  for (const id of ['btnPickFile', 'btnSaveFolder', 'btnChangeSaveDir', 'btnSaveZip', 'includeKmz']) {
    $(id).disabled = busy;
  }
  $('dropZone').classList.toggle('is-disabled', busy);
  document.body.classList.toggle('is-busy', busy);
}

function showStatus(message, kind = 'info') {
  const el = $('loadStatus');
  el.textContent = message;
  el.dataset.kind = kind;
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
