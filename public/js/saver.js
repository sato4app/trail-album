// 出力の保存: フォルダへ直接保存(File System Access API)と ZIP のダウンロード
// 保存する内容は entries([{ path: '….xlsx' | 'photos/001_….jpg' | '….kmz', data: Blob | Uint8Array }])
// で受け取り、どちらの方法でも同じ構成で書き出す。

// フォルダへ直接保存できるブラウザか(Chrome / Edge。安全な接続 https / localhost のときだけ使える)
export const canSaveToFolder = typeof globalThis.showDirectoryPicker === 'function';

// ===== 保存先(IndexedDB) =====
// 保存先(出力フォルダを作る親フォルダ)は、最初に1回選べば覚えておき、次からは選ぶ画面を出さずに使う。
// フォルダのハンドルは IndexedDB にそのまま保存できる。
// 書き込みの権限はブラウザが管理する。同じ訪問のうち(このサイトのタブを全部閉じるまで)は、
// 一度許可されれば確認なしで書き込める。次の訪問では requestPermission で改めて確認が出るが、
// Chrome 122 以降はその確認で「毎回許可」を選べば、以後は確認なしで書き込める。
const DB_NAME = 'trail-album';
const STORE = 'settings';
const SAVE_DIR_KEY = 'lastSaveDir'; // 以前の版(前回の保存先を選ぶ画面の初期位置に使っていた)と同じキー
const RW = { mode: 'readwrite' };

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

// 保存先(起動時に loadSaveDir で読み込んでおく)。未設定なら null
let saveDir = null;

// 起動時に呼ぶ。権限の確認や保存先を選ぶ画面はボタンを押した直後に出す必要があるため、
// IndexedDB の読み出し(非同期)は前もって済ませておく。覚えている保存先(無ければ null)を返す。
export async function loadSaveDir() {
  if (!canSaveToFolder) return null;
  try {
    saveDir = (await dbRequest('readonly', (s) => s.get(SAVE_DIR_KEY))) || null;
  } catch {
    saveDir = null;
  }
  return saveDir;
}

// 今の保存先(未設定なら null)
export function currentSaveDir() {
  return saveDir;
}

async function storeSaveDir(handle) {
  saveDir = handle;
  try {
    if (handle) await dbRequest('readwrite', (s) => s.put(handle, SAVE_DIR_KEY));
    else await dbRequest('readwrite', (s) => s.delete(SAVE_DIR_KEY));
  } catch {
    // 覚えられなくても保存自体は続ける
  }
}

// ===== フォルダへ直接保存 =====
// 保存先を選ぶ画面を出し、選ばれたフォルダを保存先として覚える。キャンセルされたら null。
// ボタンを押した直後に呼ぶこと(間に時間のかかる処理を挟むと、利用者の操作によるものと見なされず止められる)。
export async function pickSaveDir() {
  const options = { ...RW };
  const startIn = saveDir;
  if (startIn) options.startIn = startIn;
  try {
    const handle = await globalThis.showDirectoryPicker(options);
    await storeSaveDir(handle);
    return handle;
  } catch (err) {
    if (err?.name === 'AbortError') return null;
    if (startIn) {
      // 覚えていた保存先が消えた・移動したなどで画面を開けないときは、覚えた保存先を忘れる。
      // 画面はボタンを押した直後にしか開けないため、ここで開き直さず、もう一度押してもらう
      await storeSaveDir(null);
      throw new Error('保存先を選ぶ画面を開けませんでした。もう一度ボタンを押してください。');
    }
    throw err;
  }
}

// 「フォルダに保存」を押した直後に呼び、書き込める保存先を返す。
// - 保存先が未設定なら、保存先を選ぶ画面を出す(キャンセルされたら null)
// - 保存先が決まっていれば、選ぶ画面は出さない。書き込みの権限がまだ無いとき(次の訪問など)だけ、
//   ブラウザの確認を出す。許可されなければエラーを投げる
// - 保存先のフォルダが移動・削除されていたら、覚えた保存先を忘れてエラーを投げる
export async function prepareSaveDir() {
  if (!saveDir) return pickSaveDir();
  const dir = saveDir;
  if ((await dir.queryPermission(RW)) !== 'granted' && (await dir.requestPermission(RW)) !== 'granted') {
    throw new Error(
      `保存先「${dir.name}」への書き込みが許可されませんでした。` +
      'もう一度「フォルダに保存」を押して許可するか、「保存先を変更」で別のフォルダを選んでください。'
    );
  }
  if (!(await dirExists(dir))) {
    await storeSaveDir(null);
    throw new Error(
      `保存先のフォルダ「${dir.name}」が見つかりません(移動・削除された可能性があります)。` +
      'もう一度「フォルダに保存」を押して、保存先を選び直してください。'
    );
  }
  return dir;
}

// フォルダがまだあるか(中身の一覧を読み始めてみる)
async function dirExists(dir) {
  try {
    await dir.keys().next();
    return true;
  } catch (err) {
    if (err?.name === 'NotFoundError') return false;
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
    // Blob(Excel ファイル)はバイト列にしてから入れる(どの環境の JSZip でも確実に読めるように)
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
