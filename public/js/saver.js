// 出力の保存: フォルダへ直接保存(File System Access API)と ZIP のダウンロード
// 保存する内容は entries([{ path: '写真一覧.xlsx' | 'photos/001_….jpg' | '….kmz', data: Blob | Uint8Array }])
// で受け取り、どちらの方法でも同じ構成で書き出す。

// フォルダへ直接保存できるブラウザか(Chrome / Edge。安全な接続 https / localhost のときだけ使える)
export const canSaveToFolder = typeof globalThis.showDirectoryPicker === 'function';

// ===== 前回選んだ保存先(IndexedDB) =====
// フォルダのハンドルは IndexedDB にそのまま保存できる。次回は保存先を選ぶ画面の初期位置にする
// (書き込みの権限は保存先を選ぶ画面で改めて確認される)。
const DB_NAME = 'trail-album';
const STORE = 'settings';
const LAST_DIR_KEY = 'lastSaveDir';

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbRequest(mode, fn) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const req = fn(db.transaction(STORE, mode).objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

// 前回選んだ保存先(起動時に loadLastSaveDir で読み込んでおく)
let lastSaveDir = null;

// 起動時に呼ぶ。保存先を選ぶ画面はボタンを押した直後に開く必要があるため、
// IndexedDB の読み出し(非同期)は前もって済ませておく
export async function loadLastSaveDir() {
  if (!canSaveToFolder) return;
  try {
    lastSaveDir = (await dbRequest('readonly', (s) => s.get(LAST_DIR_KEY))) || null;
  } catch {
    lastSaveDir = null;
  }
}

async function storeLastSaveDir(handle) {
  lastSaveDir = handle;
  try {
    if (handle) await dbRequest('readwrite', (s) => s.put(handle, LAST_DIR_KEY));
    else await dbRequest('readwrite', (s) => s.delete(LAST_DIR_KEY));
  } catch {
    // 覚えられなくても保存自体は続ける
  }
}

// ===== フォルダへ直接保存 =====
// 保存先(出力フォルダを作る親フォルダ)を選ばせる。ボタンを押した直後に呼ぶこと
// (間に時間のかかる処理を挟むと、利用者の操作によるものと見なされず止められる)。
// キャンセルされたら null。
export async function pickParentFolder() {
  const options = { mode: 'readwrite' };
  const startIn = lastSaveDir;
  if (startIn) options.startIn = startIn;
  try {
    const handle = await globalThis.showDirectoryPicker(options);
    await storeLastSaveDir(handle);
    return handle;
  } catch (err) {
    if (err?.name === 'AbortError') return null;
    if (startIn) {
      // 覚えていた保存先が消えた・移動したなどで開けないときは、覚えた保存先を忘れる。
      // 画面はボタンを押した直後にしか開けないため、ここで開き直さず、もう一度押してもらう
      await storeLastSaveDir(null);
      throw new Error('保存先を選ぶ画面を開けませんでした。もう一度「フォルダに保存」を押してください。');
    }
    throw err;
  }
}

// 親フォルダの中に出力フォルダを作り、entries を書き込む。実際に作ったフォルダ名を返す。
// 同じ名前のフォルダ(またはファイル)が既にあれば上書きせず、「名前 (2)」のように番号を付ける。
// onProgress(path) はファイルを1つ書き終えるたびに呼ぶ。
// 途中で失敗したときは、作ったフォルダ名を err.partialFolder に入れて投げ直す。
export async function saveToFolder(parentHandle, folderName, entries, onProgress) {
  const name = await findFreeName(parentHandle, folderName);
  const root = await parentHandle.getDirectoryHandle(name, { create: true });
  try {
    for (const entry of entries) {
      const segments = entry.path.split('/');
      const fileName = segments.pop();
      let dir = root;
      for (const seg of segments) dir = await dir.getDirectoryHandle(seg, { create: true });
      const fileHandle = await dir.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      try {
        await writable.write(entry.data);
        await writable.close();
      } catch (err) {
        await writable.abort().catch(() => {});
        throw err;
      }
      onProgress?.(entry.path);
    }
  } catch (err) {
    err.partialFolder = name;
    throw err;
  }
  return name;
}

async function findFreeName(parentHandle, baseName) {
  for (let n = 1; n < 1000; n++) {
    const name = n === 1 ? baseName : `${baseName} (${n})`;
    if (!(await existsIn(parentHandle, name))) return name;
  }
  throw new Error('同じ名前のフォルダが多すぎて、新しいフォルダを作れませんでした。');
}

async function existsIn(parentHandle, name) {
  try {
    await parentHandle.getDirectoryHandle(name);
    return true;
  } catch (err) {
    if (err?.name === 'NotFoundError') return false;
    if (err?.name === 'TypeMismatchError') return true; // 同じ名前のファイルがある
    throw err;
  }
}

// ===== ZIP でダウンロード =====
// entries を ZIP の直下に置く。Windows の「すべて展開」は ZIP 名のフォルダを作って展開するため、
// ZIP 名を出力フォルダ名にしておけば、展開すると出力フォルダと同じ構成になる。
// ファイル名は UTF-8 で格納する(JSZip は日本語の名前に UTF-8 の印を付ける)。
// 写真・xlsx・KMZ はもともと圧縮されているので、無圧縮で格納する。
// onProgress(path) は ZIP の作成でファイルを1つ処理するたびに呼ぶ。
export async function buildZipBlob(entries, onProgress) {
  const zip = new globalThis.JSZip();
  for (const entry of entries) {
    // Blob(写真一覧.xlsx)はバイト列にしてから入れる(どの環境の JSZip でも確実に読めるように)
    const data = entry.data instanceof Blob ? new Uint8Array(await entry.data.arrayBuffer()) : entry.data;
    zip.file(entry.path, data, { binary: true });
  }
  let current = null;
  return zip.generateAsync({ type: 'blob', compression: 'STORE', mimeType: 'application/zip' }, (meta) => {
    // フォルダ(JSZip が自動で作る「photos/」)は数えない
    if (meta.currentFile && !meta.currentFile.endsWith('/') && meta.currentFile !== current) {
      current = meta.currentFile;
      onProgress?.(current);
    }
  });
}

export function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // ダウンロードが始まるまで少し待ってから解放する
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
