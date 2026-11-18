// フォルダ名・ファイル名の作成と、Windows で使えない文字の除去
// 日時はすべて日本時間で扱う。端末のタイムゾーン設定に左右されないよう、
// 9時間ずらした時刻を UTC として読み出す(getUTC〜)。

export const JST_OFFSET_MS = 9 * 60 * 60 * 1000;

// 名称部分の最大文字数(長すぎるとパス全体が Windows の上限に近づくため)
export const MAX_NAME_LENGTH = 40;

const pad2 = (n) => String(n).padStart(2, '0');
export const pad3 = (n) => String(n).padStart(3, '0');

// 日本時間の年・月・日・時・分
function jstParts(ms) {
  const d = new Date(ms + JST_OFFSET_MS);
  return {
    y: d.getUTCFullYear(),
    mo: pad2(d.getUTCMonth() + 1),
    d: pad2(d.getUTCDate()),
    h: pad2(d.getUTCHours()),
    mi: pad2(d.getUTCMinutes())
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
export function sanitizeName(text, maxLength = MAX_NAME_LENGTH) {
  const cleaned = String(text ?? '')
    .replace(/[\r\n\t]+/g, ' ')
    // eslint-disable-next-line no-control-regex
    .replace(/[\\/:*?"<>|#%\u0000-\u001f\u007f]/g, '')
    .trim();
  // サロゲートペア(絵文字など)を途中で切らないよう、文字単位で数える
  const cut = Array.from(cleaned).slice(0, maxLength).join('');
  return cut.replace(/[.\s]+$/u, '');
}

// 画面と Excel に出す名称。空のときは「写真001」のように番号で補う
export function displayNameOf(photo, no) {
  return photo.name || `写真${pad3(no)}`;
}

// 写真のファイル名: 番号3桁_YYYYMMDD-HHmm_名称.拡張子
// 撮影日時が無ければ日時部分を、名称が空なら名称部分を省く
export function buildPhotoFileName(no, photo) {
  const parts = [pad3(no)];
  if (Number.isFinite(photo.timeMs)) parts.push(formatJstCompact(photo.timeMs));
  const name = sanitizeName(photo.name);
  if (name) parts.push(name);
  return `${parts.join('_')}.${photo.ext}`;
}

// 拡張子を除いたファイル名
export function baseNameOf(fileName) {
  return String(fileName).replace(/\.[^./\\]*$/, '');
}

// 出力フォルダ名: yyyymmdd_<KML の Document 名>
// Document 名が無ければ KMZ のファイル名(拡張子なし)を使う。
// その名前が日付で始まるとき(20261029-ようらく台園地 など)は、日付が2回入らないよう名前だけにする。
// 前に付ける日付は最も早い撮影日(日本時間)。撮影日時が1枚も無ければ nowMs の日付。
export function buildFolderName(album, nowMs) {
  const title = sanitizeName(album.title) || sanitizeName(baseNameOf(album.fileName)) || '写真';
  if (startsWithDate(title)) return title;
  const times = album.photos.map((p) => p.timeMs).filter(Number.isFinite);
  const dateMs = times.length > 0 ? Math.min(...times) : nowMs;
  return `${formatJstDate(dateMs)}_${title}`;
}

// 出力の計画(フォルダ名・各写真の番号・名称・ファイル名・KMZ と Excel の保存名)を作る。
// album.photos は出力順(撮影日時の昇順)に並んでいること。
export function buildOutputPlan(album, nowMs) {
  // 元の KMZ はファイル名をそのまま使う(念のため使えない文字だけ除く)
  const kmzFileName = sanitizeName(album.fileName, 120) || 'original.kmz';
  const photos = album.photos.map((photo, i) => {
    const no = i + 1;
    return {
      ...photo,
      no,
      displayName: displayNameOf(photo, no),
      fileName: buildPhotoFileName(no, photo)
    };
  });
  return {
    folderName: buildFolderName(album, nowMs),
    photos,
    kmzFileName,
    // Excel は KMZ と同じ名前にする(20261029-ようらく台園地.kmz → 20261029-ようらく台園地.xlsx)
    xlsxFileName: `${baseNameOf(kmzFileName) || '写真一覧'}.xlsx`
  };
}
