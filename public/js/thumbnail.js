// サムネイル(縮小版)の作成
// 縮小版は1種類(長辺 THUMB_LONG_SIDE の JPEG)。読み込んだときに写真ごとに作っておき(createThumbnail)、
// 画面の写真の一覧・地図の吹き出し・Excel に埋め込む画像・「縮小版をコピー」のすべてに使う。
// 「縮小版をコピー」では、クリップボードに書き込めるのが PNG だけのため、縮小版を同じ大きさの PNG にする(convertToPng)。
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
    // PNG の透明部分が JPEG で黒くならないよう、白で塗っておく
    const canvas = drawScaled(source, THUMB_LONG_SIDE, '#ffffff');
    const thumbBlob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', THUMB_JPEG_QUALITY));
    if (!thumbBlob) return null;
    return {
      blob: thumbBlob,
      bytes: new Uint8Array(await thumbBlob.arrayBuffer()),
      width: canvas.width,
      height: canvas.height
    };
  } finally {
    source.close();
  }
}

// 画像(Blob)を、同じ大きさの PNG にする(「縮小版をコピー」で、縮小版をクリップボードに書き込むため:
// ブラウザがクリップボードに書き込める画像は PNG だけ)。PNG ならそのまま返す
export async function convertToPng(blob) {
  if (blob.type === 'image/png') return blob;
  const source = await decodeImage(blob);
  try {
    const canvas = drawScaled(source, Math.max(source.width, source.height)); // 大きさはそのまま
    const png = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!png) throw new Error('PNG に変換できませんでした');
    return png;
  } finally {
    source.close();
  }
}

// 読み込んだ画像を、長辺 longSide まで縮小して(小さければそのままの大きさで)描いた canvas を作る。
// background を渡すと、先にその色で塗る
function drawScaled(source, longSide, background = '') {
  const scale = Math.min(1, longSide / Math.max(source.width, source.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(source.width * scale));
  canvas.height = Math.max(1, Math.round(source.height * scale));
  const ctx = canvas.getContext('2d');
  if (background) {
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(source.image, 0, 0, canvas.width, canvas.height);
  return canvas;
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
