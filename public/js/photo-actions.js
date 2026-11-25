// 写真の一覧のボタンの処理(縮小版をコピー・写真を保存)
import { convertToPng } from './thumbnail.js';
import { downloadBlob } from './saver.js';

// クリップボードに画像をコピーできるか(https か localhost で、Clipboard API に対応したブラウザ)
export const canCopyImage = !!(globalThis.isSecureContext && navigator.clipboard?.write && globalThis.ClipboardItem);

// 写真の縮小版(Blob。thumbnail.js の createThumbnail で作ったもの)を、同じ大きさの PNG にしてクリップボードにコピーする。
// ボタンを押した処理の中で、待たずに呼ぶこと
// (ブラウザは、押した操作の直後でないとクリップボードへの書き込みを許さない。
//  Safari に合わせ、PNG への変換は終わるのを待たずに Promise のまま渡す)
export async function copyImage(thumbBlob) {
  await navigator.clipboard.write([new ClipboardItem({ 'image/png': convertToPng(thumbBlob) })]);
}

// 元の写真を、保存するときと同じファイル名でダウンロードする(KMZ の中の写真をそのまま。再圧縮しない)
export function saveImage(photo) {
  downloadBlob(photo.blob, photo.fileName);
}

// ボタンの文字を少しの間だけ結果に変えて知らせる(「コピーしました」など)
export function flashButton(button, text, ms = 1500) {
  const original = button.dataset.label || button.textContent;
  button.dataset.label = original;
  button.textContent = text;
  clearTimeout(Number(button.dataset.timer));
  button.dataset.timer = String(setTimeout(() => {
    button.textContent = original;
  }, ms));
}
