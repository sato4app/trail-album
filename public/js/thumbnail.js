// サムネイルの作成
// 画面のプレビューと、Excel に埋め込む縮小版の両方に使う。
// 長辺 約400px の JPEG にする(Excel ファイルが重くならないように)。
// 元の写真(photos フォルダに保存するもの)には手を加えない。
// 長辺の長さと JPEG の品質は constants.js に置く。
import { THUMB_LONG_SIDE, THUMB_JPEG_QUALITY } from './constants.js';

// 画像(Blob)を縮小した JPEG を作る。
// 戻り値: { blob, bytes(Uint8Array), width, height }。表示できない形式なら null。
export async function createThumbnail(blob) {
  let source;
  try {
    source = await decodeImage(blob);
  } catch {
    return null;
  }
  try {
    const scale = Math.min(1, THUMB_LONG_SIDE / Math.max(source.width, source.height));
    const width = Math.max(1, Math.round(source.width * scale));
    const height = Math.max(1, Math.round(source.height * scale));

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    // PNG の透明部分が JPEG で黒くならないよう、白で塗っておく
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(source.image, 0, 0, width, height);

    const thumbBlob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', THUMB_JPEG_QUALITY));
    if (!thumbBlob) return null;
    return {
      blob: thumbBlob,
      bytes: new Uint8Array(await thumbBlob.arrayBuffer()),
      width,
      height
    };
  } finally {
    source.close();
  }
}

// 画像(Blob)を PNG に変換する(クリップボードへのコピー用: ブラウザが書き込めるのは PNG だけのため)。
// 大きさは元のまま。EXIF の向きを反映する。PNG ならそのまま返す
export async function convertToPng(blob) {
  if (blob.type === 'image/png') return blob;
  const source = await decodeImage(blob);
  try {
    const canvas = document.createElement('canvas');
    canvas.width = source.width;
    canvas.height = source.height;
    canvas.getContext('2d').drawImage(source.image, 0, 0);
    const png = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!png) throw new Error('PNG に変換できませんでした');
    return png;
  } finally {
    source.close();
  }
}

// 画像を読み込む。EXIF の向き(縦向きで撮った写真など)を反映する。
// createImageBitmap が使えない・失敗したときは <img> で読み込む(<img> も EXIF の向きを反映する)。
async function decodeImage(blob) {
  if (typeof createImageBitmap === 'function') {
    try {
      const bitmap = await createImageBitmap(blob, { imageOrientation: 'from-image' });
      return { image: bitmap, width: bitmap.width, height: bitmap.height, close: () => bitmap.close() };
    } catch {
      // <img> での読み込みを試す
    }
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    return { image: img, width: img.naturalWidth, height: img.naturalHeight, close: () => URL.revokeObjectURL(url) };
  } catch (err) {
    URL.revokeObjectURL(url);
    throw err;
  }
}
