// 「内容を確かめる」の KMZ(左)と GPX(右)の境界線を動かして、左右の幅の比率を変える。
// 左(KMZ)の列の幅を px で決め、残りを右(GPX・地図)の列に回す(画面の幅が変わっても KMZ の幅は変えない)。
// 決めた幅はこのブラウザに覚えておき、次に開いたときも使う。
//   ドラッグ: 境界線(どこを掴んでもよい)を左右に動かす
//   キーボード: 境界線にフォーカスして ← → (Shift を押しながらだと大きく動かす)
//   ダブルクリック・Home: 元の幅(KMZ_COL_DEFAULT_WIDTH。画面が狭ければ左右半分ずつ)に戻す
// どちらの列も COL_MIN_WIDTH より狭くはしない。
// 幅の設定値は constants.js に置き、列の並べ方を決める style.css へは CSS 変数(--col-min・--kmz-default-width)で渡す。
import { KMZ_COL_DEFAULT_WIDTH, COL_MIN_WIDTH, SPLITTER_KEY_STEP, SPLITTER_KEY_STEP_LARGE } from './constants.js';

const STORAGE_KEY = 'trail-album.kmzColWidth';

// cols: 列を並べる要素(.preview-cols)、splitter: 境界線、leftCol: 左(KMZ)の列
export function setupColSplitter(cols, splitter, leftCol) {
  cols.style.setProperty('--col-min', `${COL_MIN_WIDTH}px`);
  cols.style.setProperty('--kmz-default-width', `${KMZ_COL_DEFAULT_WIDTH}px`);
  const maxWidth = () => cols.clientWidth - splitter.offsetWidth - COL_MIN_WIDTH;
  const clampWidth = (px) => Math.round(Math.max(COL_MIN_WIDTH, Math.min(px, maxWidth())));

  // 左の列の幅を px で決める(null なら元の幅に戻す)
  const setWidth = (px) => {
    if (px === null) cols.style.removeProperty('--kmz-width');
    else cols.style.setProperty('--kmz-width', `${px}px`);
    updateAria();
  };

  // 支援技術向けに、左の列が占める割合(%)を伝える
  const updateAria = () => {
    const total = cols.clientWidth - splitter.offsetWidth;
    if (total <= 0) return; // 隠れている間は測れない
    splitter.setAttribute('aria-valuenow', String(Math.round((leftCol.offsetWidth / total) * 100)));
  };

  const save = (px) => {
    try {
      if (px === null) localStorage.removeItem(STORAGE_KEY);
      else localStorage.setItem(STORAGE_KEY, String(px));
    } catch {
      // 保存できない(プライベートモードなど)ときは、今の画面だけで使う
    }
  };

  // 前に決めた幅を使う(列の大きさは CSS の clamp で画面に収める)
  try {
    const saved = Number(localStorage.getItem(STORAGE_KEY));
    if (saved > 0) setWidth(saved);
  } catch {
    // 読み込めないときは元の幅のまま
  }
  splitter.setAttribute('aria-valuemin', '0');
  splitter.setAttribute('aria-valuemax', '100');
  new ResizeObserver(updateAria).observe(cols);

  // ドラッグ(マウス・タッチ・ペン)
  splitter.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    splitter.setPointerCapture(e.pointerId);
    const startX = e.clientX;
    const startWidth = leftCol.offsetWidth;
    let width = startWidth;
    splitter.classList.add('is-dragging');
    document.body.classList.add('is-resizing');

    const onMove = (ev) => {
      width = clampWidth(startWidth + ev.clientX - startX);
      setWidth(width);
    };
    const onEnd = () => {
      splitter.removeEventListener('pointermove', onMove);
      splitter.removeEventListener('lostpointercapture', onEnd);
      splitter.classList.remove('is-dragging');
      document.body.classList.remove('is-resizing');
      if (width !== startWidth) save(width);
    };
    splitter.addEventListener('pointermove', onMove);
    splitter.addEventListener('lostpointercapture', onEnd); // 指を離した・取り消されたときに起きる
  });

  splitter.addEventListener('dblclick', () => {
    setWidth(null);
    save(null);
  });

  splitter.addEventListener('keydown', (e) => {
    const step = e.shiftKey ? SPLITTER_KEY_STEP_LARGE : SPLITTER_KEY_STEP;
    let width;
    if (e.key === 'ArrowLeft') width = clampWidth(leftCol.offsetWidth - step);
    else if (e.key === 'ArrowRight') width = clampWidth(leftCol.offsetWidth + step);
    else if (e.key === 'Home') width = null;
    else return;
    e.preventDefault();
    setWidth(width);
    save(width);
  });
}
