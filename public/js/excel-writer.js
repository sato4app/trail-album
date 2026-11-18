// 写真一覧.xlsx の作成(ExcelJS を CDN から読み込んで使う)
// シート「写真一覧」に写真1枚を1行で並べ、シート「情報」に元の KMZ の情報を書く。
import { JST_OFFSET_MS } from './naming.js';

export const XLSX_FILE_NAME = '写真一覧.xlsx';
export const PHOTOS_DIR = 'photos';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const APP_NAME = 'Trail Album';
const INFO_NOTE = '元の画質の写真は photos フォルダにあります。Excel に埋め込んだ画像は縮小版です。';

// ===== サムネイルの配置 =====
// サムネイルは「写真」列(B列)のセルの内側に収め、縦横比を保ったまま表示枠に合わせて縮める。
// 配置は「セルに合わせて移動やサイズ変更をする」(twoCell)にする。並べ替えで行と一緒に動き、
// オートフィルターで行を隠したときも画像が一緒に隠れる(「移動するがサイズ変更はしない」では隠れない)。
const THUMB_BOX_PX = 150;       // 表示枠(長辺の最大 px)
const THUMB_MARGIN_PX = 8;      // 画像の上下の余白(px)
const PHOTO_COL_WIDTH = 25;     // B列の幅(文字数)。約175px で、左右に約12px の余白ができる
const NO_THUMB_ROW_HEIGHT = 30; // サムネイルを作れなかった行の高さ(pt)
const EMU_PER_PX = 9525;        // Excel の図形の位置の単位(EMU)。96dpi で 1px = 9525EMU
const PT_PER_PX = 0.75;         // 行の高さはポイント単位。96dpi で 1px = 0.75pt

// 列幅(文字数)を px に直す(既定のフォント Calibri 11 の数字幅 7px で計算する Excel の式)
function columnWidthToPx(width) {
  const digit = 7;
  return Math.trunc(((256 * width + Math.trunc(128 / digit)) / 256) * digit);
}

// ===== 書式 =====
const FONT = { name: '游ゴシック', size: 11 };
const HEADER_FONT = { ...FONT, bold: true };
const LINK_FONT = { ...FONT, color: { argb: 'FF0563C1' }, underline: true };
const MUTED_FONT = { ...FONT, color: { argb: 'FF888888' } };
const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBD9' } };
const THIN = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const DATE_FORMAT = 'yyyy/mm/dd hh:mm';
const COORD_FORMAT = '0.000000';

// 写真一覧の列(A〜J)。印刷で横1ページに収めたとき文字が小さくなりすぎないよう、
// 長くなりやすい名称・写真ファイルは幅に上限を設け、セルの中で折り返す
const COLUMNS = [
  { header: 'No.', align: 'center' },
  { header: '写真', align: 'center', width: PHOTO_COL_WIDTH },
  { header: '名称', align: 'left', wrap: true, maxWidth: 30 },
  { header: '撮影日時', align: 'center' },
  { header: '緯度', align: 'right' },
  { header: '経度', align: 'right' },
  { header: '地理院地図', align: 'center' },
  { header: 'Google マップ', align: 'center' },
  { header: '写真ファイル', align: 'left', wrap: true, maxWidth: 36 },
  { header: '備考', align: 'left', wrap: true, width: 24 }
];

// 印刷設定(A4 縦・横1ページに収める・縦は必要なだけ)
const PAGE_SETUP = {
  paperSize: 9,
  orientation: 'portrait',
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 0,
  horizontalCentered: true,
  margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
};

// 表示幅の目安(全角を2、半角を1と数える)
function textWidth(text) {
  let w = 0;
  for (const ch of String(text)) w += /[\u0000-ÿ｡-ﾟ]/.test(ch) ? 1 : 2;
  return w;
}

// ExcelJS は Date を UTC として書き込むため、日本時間で表示されるよう9時間ずらした値を渡す
function excelJstDate(ms) {
  return new Date(ms + JST_OFFSET_MS);
}

const coord = (v) => Number(v.toFixed(6));

// Excel の「写真ファイル」のリンク先(写真一覧.xlsx からの相対パス)。
// 空白は %20 にする(Excel 自身もリンクの空白をこの形で保存する)
export function photoLinkTarget(fileName) {
  return `${PHOTOS_DIR}/${fileName.replaceAll(' ', '%20')}`;
}

// 写真一覧.xlsx を作り、Blob で返す。
// plan は naming.js の buildOutputPlan の結果(各写真に thumb: { bytes, width, height } | null を持つ)。
export async function buildWorkbookBlob(album, plan, nowMs) {
  const ExcelJS = globalThis.ExcelJS;
  const wb = new ExcelJS.Workbook();
  wb.creator = APP_NAME;
  wb.created = new Date(nowMs);

  writePhotoSheet(wb, plan);
  writeInfoSheet(wb, album, plan, nowMs);

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: XLSX_MIME });
}

function writePhotoSheet(wb, plan) {
  const ws = wb.addWorksheet('写真一覧', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }],
    pageSetup: { ...PAGE_SETUP }
  });
  ws.pageSetup.printTitlesRow = '1:1'; // 見出し行を各ページに繰り返す

  // 見出し行
  const headerRow = ws.getRow(1);
  COLUMNS.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col.header;
    cell.font = HEADER_FONT;
    cell.fill = HEADER_FILL;
    cell.border = BORDER;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  headerRow.height = 22;

  // 列幅を決めるための表示文字列(列ごと)
  const texts = COLUMNS.map((col) => [col.header]);
  const photoColPx = columnWidthToPx(PHOTO_COL_WIDTH);

  plan.photos.forEach((photo, i) => {
    const rowNumber = i + 2;
    const row = ws.getRow(rowNumber);
    const latText = photo.lat.toFixed(6);
    const lngText = photo.lng.toFixed(6);

    const values = [
      photo.no,
      photo.thumb ? '' : '(表示できません)',
      photo.displayName,
      Number.isFinite(photo.timeMs) ? excelJstDate(photo.timeMs) : null,
      coord(photo.lat),
      coord(photo.lng),
      { text: '地図を開く', hyperlink: `https://maps.gsi.go.jp/#17/${latText}/${lngText}/` },
      { text: '地図を開く', hyperlink: `https://www.google.com/maps?q=${latText},${lngText}` },
      { text: photo.fileName, hyperlink: photoLinkTarget(photo.fileName) },
      null
    ];
    values.forEach((value, c) => {
      const cell = row.getCell(c + 1);
      cell.value = value;
      cell.border = BORDER;
      cell.font = value && typeof value === 'object' && 'hyperlink' in value ? LINK_FONT : FONT;
      cell.alignment = { horizontal: COLUMNS[c].align, vertical: 'middle', wrapText: !!COLUMNS[c].wrap };
    });
    row.getCell(4).numFmt = DATE_FORMAT;
    row.getCell(5).numFmt = COORD_FORMAT;
    row.getCell(6).numFmt = COORD_FORMAT;
    if (!photo.thumb) row.getCell(2).font = MUTED_FONT;

    texts[0].push(String(photo.no));
    texts[2].push(photo.displayName);
    texts[3].push(Number.isFinite(photo.timeMs) ? '2026/01/01 00:00' : '');
    texts[4].push(latText);
    texts[5].push(lngText);
    texts[6].push('地図を開く');
    texts[7].push('地図を開く');
    texts[8].push(photo.fileName);

    // サムネイル: 表示枠に収まる大きさにして、B列のセルの中央(上下は余白 THUMB_MARGIN_PX)に置く
    if (photo.thumb) {
      const scale = Math.min(1, THUMB_BOX_PX / Math.max(photo.thumb.width, photo.thumb.height));
      const w = Math.round(photo.thumb.width * scale);
      const h = Math.round(photo.thumb.height * scale);
      row.height = (h + THUMB_MARGIN_PX * 2) * PT_PER_PX;

      const imageId = wb.addImage({ buffer: photo.thumb.bytes, extension: 'jpeg' });
      const left = Math.max(0, Math.round((photoColPx - w) / 2));
      ws.addImage(imageId, {
        // 位置は EMU で直接指定する(ExcelJS の小数指定は列幅の換算が粗く、ずれるため)
        tl: { nativeCol: 1, nativeColOff: left * EMU_PER_PX, nativeRow: rowNumber - 1, nativeRowOff: THUMB_MARGIN_PX * EMU_PER_PX },
        br: {
          nativeCol: 1,
          nativeColOff: (left + w) * EMU_PER_PX,
          nativeRow: rowNumber - 1,
          nativeRowOff: (THUMB_MARGIN_PX + h) * EMU_PER_PX
        },
        editAs: 'twoCell'
      });
    } else {
      row.height = NO_THUMB_ROW_HEIGHT;
    }
  });

  // 列幅: 指定があればその幅、無ければ内容の最も長いものに合わせる
  COLUMNS.forEach((col, i) => {
    const column = ws.getColumn(i + 1);
    if (col.width) {
      column.width = col.width;
    } else {
      const longest = Math.max(...texts[i].map(textWidth));
      column.width = Math.min(col.maxWidth ?? 40, Math.max(6, longest + 3));
    }
  });

  // オートフィルター(見出し行〜最終行)
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: plan.photos.length + 1, column: COLUMNS.length } };
}

function writeInfoSheet(wb, album, plan, nowMs) {
  const ws = wb.addWorksheet('情報', { pageSetup: { ...PAGE_SETUP } });
  const rows = [
    ['項目', '内容'],
    ['元の KMZ ファイル', album.fileName],
    ['KML の Document 名', album.title || '(なし)'],
    ['写真の枚数', plan.photos.length],
    ['出力日時', excelJstDate(nowMs)],
    ['作成アプリ', APP_NAME],
    ['注意', INFO_NOTE]
  ];
  rows.forEach((values, r) => {
    const row = ws.getRow(r + 1);
    values.forEach((value, c) => {
      const cell = row.getCell(c + 1);
      cell.value = value;
      cell.border = BORDER;
      cell.font = r === 0 ? HEADER_FONT : FONT;
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: c === 1 };
      if (r === 0) cell.fill = HEADER_FILL;
    });
  });
  ws.getCell('B5').numFmt = 'yyyy/mm/dd hh:mm:ss';
  ws.getColumn(1).width = 22;
  // 内容の列幅は文字列の行だけで測る(日付・数値は短い)
  const longest = Math.max(...rows.map((r) => (typeof r[1] === 'string' ? textWidth(r[1]) : 0)));
  ws.getColumn(2).width = Math.min(80, Math.max(30, longest + 3));
}
