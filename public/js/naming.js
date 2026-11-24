// フォルダ名・ファイル名の作成と、Windows で使えない文字の除去
// 日時はすべて日本時間で扱う。端末のタイムゾーン設定に左右されないよう、
// 9時間ずらした時刻を UTC として読み出す(getUTC〜)。
// 名前の最大文字数・写真のフォルダ名は constants.js に置く。
import { MAX_NAME_LENGTH, MAX_SOURCE_NAME_LENGTH, PHOTOS_DIR } from './constants.js';

export const JST_OFFSET_MS = 9 * 60 * 60 * 1000;

const pad2 = (n) => String(n).padStart(2, '0');
export const pad3 = (n) => String(n).padStart(3, '0');

// 日本時間の年・月・日・時・分・秒
function jstParts(ms) {
  const d = new Date(ms + JST_OFFSET_MS);
  return {
    y: d.getUTCFullYear(),
    mo: pad2(d.getUTCMonth() + 1),
    d: pad2(d.getUTCDate()),
    h: pad2(d.getUTCHours()),
    mi: pad2(d.getUTCMinutes()),
    s: pad2(d.getUTCSeconds())
  };
}

// 20261118(フォルダ名用)
export function formatJstDate(ms) {
  const p = jstParts(ms);
  return `${p.y}${p.mo}${p.d}`;
}

// 名前が日付で始まっているか(20261029-… / 2026-10-29_… など)。
// minoh-hiking の KMZ は出力時のファイル名(既定は yyyymmdd-nn)が Document 名になるため、
// 先頭に日付を持つことが多い。年月日として正しい値のときだけ日付とみなす。
export function startsWithDate(name) {
  const m = /^(\d{4})(?:(\d{2})(\d{2})|([-_.])(\d{2})\4(\d{2}))(?!\d)/.exec(name);
  if (!m) return false;
  const month = Number(m[2] ?? m[5]);
  const day = Number(m[3] ?? m[6]);
  return month >= 1 && month <= 12 && day >= 1 && day <= 31;
}

// 2026/11/18 10:32(画面表示用)
export function formatJstDateTime(ms) {
  const p = jstParts(ms);
  return `${p.y}/${p.mo}/${p.d} ${p.h}:${p.mi}`;
}

// 2026/11/18 10:32:15(地図のマーカーの表示用。KMZ・GPX に記録された日時を、秒まで日本時間で出す)
export function formatJstDateTimeSeconds(ms) {
  return `${formatJstDateTime(ms)}:${jstParts(ms).s}`;
}

// 2026/11/18 10:32〜12:45(画面表示用。日をまたぐときは終わりにも日付を付ける)
export function formatJstRange(startMs, endMs) {
  const s = jstParts(startMs);
  const e = jstParts(endMs);
  const sameDay = s.y === e.y && s.mo === e.mo && s.d === e.d;
  return `${formatJstDateTime(startMs)}〜${sameDay ? `${e.h}:${e.mi}` : formatJstDateTime(endMs)}`;
}

// 20261118-1032(写真のファイル名用)
function formatJstCompact(ms) {
  const p = jstParts(ms);
  return `${p.y}${p.mo}${p.d}-${p.h}${p.mi}`;
}

// ファイル名・フォルダ名に使えるように整える。
// - 改行・タブは空白1つにする
// - Windows で使えない文字 \ / : * ? " < > | と制御文字を除く
// - # と % も除く(Excel の「写真ファイル」のハイパーリンクで、# はファイル内の位置、
//   % は文字コードの記号と解釈され、リンクから写真を開けなくなるため)
// - 前後の空白を除き、maxLength 文字で切ったうえで、末尾のピリオド・空白を除く
//   (Windows は末尾のピリオド・空白を黙って削るため、名前が食い違う)
// eslint-disable-next-line no-control-regex
const INVALID_NAME_CHARS = /[\\/:*?"<>|#%\u0000-\u001f\u007f]/g;

export function sanitizeName(text, maxLength = MAX_NAME_LENGTH) {
  const cleaned = String(text ?? '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(INVALID_NAME_CHARS, '')
    .trim();
  // サロゲートペア(絵文字など)を途中で切らないよう、文字単位で数える
  const cut = Array.from(cleaned).slice(0, maxLength).join('');
  return cut.replace(/[.\s]+$/u, '');
}

// 「名前」の欄に入力された文字を、保存に使う名前にする(sanitizeName と同じ規則)。
// 入力欄の下で知らせるため、除いた文字(使えない文字。重なりは除く)と、長すぎて切ったかも返す。
//   { name, removed: ['/', ':' …], truncated }
export function checkSaveName(text) {
  const spaced = String(text ?? '').replace(/[\r\n\t]+/g, ' ');
  const removed = [...new Set(spaced.match(INVALID_NAME_CHARS) || [])];
  const cleaned = spaced.replace(INVALID_NAME_CHARS, '').trim();
  return {
    name: sanitizeName(text, MAX_NAME_LENGTH),
    removed,
    truncated: Array.from(cleaned).length > MAX_NAME_LENGTH
  };
}

// 画面と Excel に出す名称。空のときは「写真001」のように番号で補う
export function displayNameOf(photo, no) {
  return photo.name || `写真${pad3(no)}`;
}

// 写真のファイル名: 番号3桁_YYYYMMDD-HHmm_名称.拡張子
// 撮影日時が無ければ日時部分を、名称が空なら名称部分を省く。
// 名称部分の空白(全角の空白・改行・タブを含む)は _ に置き換える(箕面 大滝 → 箕面_大滝)。
// 名称に元からある - は、利用者が意図して付けたものとして残す(Minoh - Falls → Minoh_-_Falls)。
// 続く空白や、空白の前後にある _ は1つの _ にまとめる(見分けが付かないため)。
// 写真はもともと名前を付け直して保存するので、空白を含めなければ Excel の「写真ファイル」の
// ハイパーリンクで空白の扱い(%20 への置き換え)を気にしなくてよい。
export function buildPhotoFileName(no, photo) {
  const parts = [pad3(no)];
  if (Number.isFinite(photo.timeMs)) parts.push(formatJstCompact(photo.timeMs));
  // 先に sanitizeName で前後の空白を除いてから置き換える(名称の先頭・末尾が _ にならないように)
  const name = sanitizeName(photo.name).replace(/[\s_]*\s[\s_]*/gu, '_');
  if (name) parts.push(name);
  return `${parts.join('_')}.${photo.ext}`;
}

// 拡張子を除いたファイル名
export function baseNameOf(fileName) {
  return String(fileName).replace(/\.[^./\\]*$/, '');
}

// 写真の番号(保存の通し番号)・名称・保存するファイル名を付ける(画面の写真の一覧と保存で共通)。
// album は merge.js の buildAlbum の結果(photos は保存の順)
export function planPhotos(album) {
  return album.photos.map((photo, i) => {
    const no = i + 1;
    return {
      ...photo,
      no,
      displayName: displayNameOf(photo, no),
      fileName: buildPhotoFileName(no, photo)
    };
  });
}

// ===== 保存の「名前」(出力フォルダ名・Excel・ZIP の名前) =====
// album は merge.js の buildAlbum の結果、track は buildTrack の結果(どちらも読み込んでいなければ null)

// 名前の既定値: yyyymmdd_<名前のもと>
// 名前のもとは、KMZ があれば KML の Document 名(無ければ KMZ のファイル名)、GPX だけなら GPX のファイル名
// (拡張子なし)。KMZ・GPX が複数なら、いちばん早いもの(album.files・track.files の先頭)を使う。
// 名前のもとが日付で始まるとき(20261029-ようらく台園地 など)は、日付が2回入らないよう名前のもとだけにする。
// 前に付ける日付は、最も早い撮影日(GPX だけなら最も早い記録日。日本時間)。日時が無ければ nowMs の日付。
export function defaultSaveName(album, track, nowMs) {
  let title;
  let times;
  if (album) {
    title = sanitizeName(album.title) || sanitizeName(baseNameOf(album.fileName)) || '写真';
    times = album.photos.map((p) => p.timeMs);
  } else {
    title = sanitizeName(baseNameOf(track.files[0].source.fileName)) || '移動経路';
    times = track.files.flatMap((f) => f.segments.map((s) => s.startMs));
  }
  if (startsWithDate(title)) return title;
  const known = times.filter(Number.isFinite);
  const dateMs = known.length > 0 ? Math.min(...known) : nowMs;
  return sanitizeName(`${formatJstDate(dateMs)}_${title}`);
}

// 名前の候補: 既定値、KML の Document 名、KMZ・GPX のファイル名(拡張子なし)。使えない文字を除き、重なりは1つにする
export function saveNameCandidates(album, track, nowMs) {
  const names = [defaultSaveName(album, track, nowMs)];
  if (album) names.push(...album.titles, ...album.files.map((f) => baseNameOf(f.fileName)));
  if (track) names.push(...track.files.map((f) => baseNameOf(f.source.fileName)));
  return [...new Set(names.map((n) => sanitizeName(n)).filter(Boolean))];
}

// 一緒に保存する元のファイルの名前: 元のファイル名から使えない文字だけ除く。
// 拡張子(ext)で終わらなければ付ける(中身で KMZ・GPX と分かったファイルも、その形式として開けるように)
function sourceFileName(fileName, ext, fallback) {
  const suffix = `.${ext}`;
  const hasExt = fileName.toLowerCase().endsWith(suffix);
  const base = sanitizeName(hasExt ? fileName.slice(0, -suffix.length) : fileName, MAX_SOURCE_NAME_LENGTH);
  return `${base || fallback}${hasExt ? fileName.slice(-suffix.length) : suffix}`;
}

// 同じ名前が重なったら「名前 (2).kmz」のように番号を付ける(Windows に合わせて大文字・小文字は区別しない)。
// reserved は、ほかに使っている名前(写真のフォルダ名など)
function uniqueFileNames(names, reserved = []) {
  const used = new Set(reserved.map((n) => n.toLowerCase()));
  return names.map((name) => {
    const base = baseNameOf(name);
    const ext = name.slice(base.length);
    let result = name;
    for (let n = 2; used.has(result.toLowerCase()); n++) result = `${base} (${n})${ext}`;
    used.add(result.toLowerCase());
    return result;
  });
}

// 保存の計画: 出力フォルダ(名前 name)の中に置くものの名前を決める。
//   { folderName, xlsxFileName, photos, kmz: [{ file, fileName }], gpx: [{ file, fileName }] }
// - includeKmz が true なら、KMZ から作るもの全部(Excel(<名前>.xlsx)・写真(photos フォルダ)・元の KMZ)を保存する。
//   false か KMZ が無ければ、どれも無し(xlsxFileName は null、ほかは [])
// - includeGpx が true なら、元の GPX を保存する
// - 元の KMZ・GPX は、画面のファイルの一覧に出ているものを、元のファイルのまま出力フォルダの直下に置く。
//   file は readKmz・readGpx の結果(kmzBytes・bytes を持つ)
// - 直下のファイル名が重なれば番号を付ける(Excel の名前を優先して残す)
export function buildSavePlan({ album, track, name, includeKmz, includeGpx }) {
  const useKmz = !!album && includeKmz;
  const kmzFiles = useKmz ? album.files : [];
  const gpxFiles = track && includeGpx ? track.files.map((f) => f.source) : [];
  const names = uniqueFileNames([
    ...(useKmz ? [`${name}.xlsx`] : []),
    ...kmzFiles.map((f) => sourceFileName(f.fileName, 'kmz', 'original')),
    ...gpxFiles.map((f) => sourceFileName(f.fileName, 'gpx', 'track'))
  ], useKmz ? [PHOTOS_DIR] : []);
  return {
    folderName: name,
    xlsxFileName: useKmz ? names.shift() : null,
    photos: useKmz ? planPhotos(album) : [],
    kmz: kmzFiles.map((file) => ({ file, fileName: names.shift() })),
    gpx: gpxFiles.map((file) => ({ file, fileName: names.shift() }))
  };
}
