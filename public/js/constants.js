// Trail Album の設定値
// 後から変えることのある値を、ここにまとめる。値を変えるときは、このファイルだけを直せばよい。
// (ファイル形式の決まり・単位の換算・ブラウザに保存するときのキー名など、変えると動かなくなる値は、
//  それぞれのモジュールに置く)

// ===== 地図(map-view.js) =====

// 地理院タイル(標準地図)。出典は国土地理院の決まりに従い、「地理院タイル」と書いて地理院タイル一覧ページへリンクする
export const GSI_TILE_URL = 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png';
export const GSI_ATTRIBUTION =
  '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>';
// 地理院タイル(標準地図)が用意されている最大のズームレベル。MAP_MAX_ZOOM をこれより大きくしたときは、
// この大きさの画像を引き伸ばして表示する
export const GSI_TILE_MAX_NATIVE_ZOOM = 18;

// ズームレベルの範囲(縮小の下限・拡大の上限)。地点の範囲が下限でも収まらないときは、下限のまま
// 範囲の真ん中を中心にする(端の地点は地図の外に出る)
export const MAP_MIN_ZOOM = 10;
export const MAP_MAX_ZOOM = 18;
// ズームレベルの刻み。地点の範囲が地図いっぱいに収まるよう、整数より細かくする
// (整数のままだと、範囲が地図の半分ほどにしかならないことがある)
export const ZOOM_SNAP = 0.1;
// 地図を拡大・縮小したときに自動で開いたズームレベルのスライドバーを、操作が無いまま閉じるまでの秒数
// (ズームレベルのボタンで開いたときは、その後にズームレベルを変えなければ、自動では閉じない)
export const ZOOM_SLIDER_AUTO_CLOSE_SEC = 5;

// 地点を読み込む前の地図(箕面大滝)。地図は地点を読み込んでから出すので、普段は目にしない
export const INITIAL_CENTER = [34.853667, 135.472041];
export const INITIAL_ZOOM = 15;

// 地点の範囲に地図を合わせるときの、ズームレベルの上限(地点が1つ・近いときに寄りすぎないように)
export const FIT_MAX_ZOOM = 17;
// いちばん端の地点と、地図の縁との間(px。マーカーの円が縁で切れないように)
export const FIT_PADDING = 24;

// マーカーと線の色。写真は1色、経路は GPX ごとに順に使い、足りなければ繰り返す(表の 1, 2 … の印と同じ色)
export const PHOTO_COLOR = '#e8590c';
export const TRACK_COLORS = ['#1c7ed6', '#7048e8', '#0c8599', '#d6336c', '#5c940d', '#495057'];

// マーカーと線の大きさ(px)
export const PHOTO_MARKER_RADIUS = 7; // 写真の撮影場所の円の半径
export const TRACK_POINT_RADIUS = 3;  // 記録地点の円の半径
export const TRACK_LINE_WEIGHT = 3;   // 経路の線の太さ
export const TRACK_LINE_OPACITY = 0.6; // 経路の線の不透明度(0〜1)

// 写真の撮影場所のマーカーを押したときに出す、写真の縮小版の大きさ(長辺の px。縮小版より大きくはしない)
export const PHOTO_POPUP_THUMB_SIZE = 240;

// 地図の右上の「別ウインドウで開く」で開くウインドウの大きさ(px。画面より大きければ画面に合わせる)。
// 開いた後は、利用者がウインドウの大きさを変えられる(地図も付いてくる)
export const MAP_WINDOW_WIDTH = 1200;
export const MAP_WINDOW_HEIGHT = 900;

// ===== KMZ と GPX の列(col-splitter.js が style.css に渡す) =====

// KMZ の列の既定の幅(px)。画面の幅が狭いときは、左右半分ずつにする
export const KMZ_COL_DEFAULT_WIDTH = 463;
// KMZ・GPX のどちらの列も、これより狭くしない(px)
export const COL_MIN_WIDTH = 280;
// 境界線を ← → キーで動かす幅(px)。Shift を押しながらだと大きく動かす
export const SPLITTER_KEY_STEP = 20;
export const SPLITTER_KEY_STEP_LARGE = 80;

// ===== 縮小版(thumbnail.js) =====

// 画面の一覧と Excel に使う縮小版の長辺(px)と、JPEG の品質(0〜1)
export const THUMB_LONG_SIDE = 400;
export const THUMB_JPEG_QUALITY = 0.85;

// ===== 保存(naming.js・excel-writer.js・app.js) =====

// Excel の作成者・「情報」シートに書くアプリ名
export const APP_NAME = 'Trail Album';
// 写真を入れるフォルダの名前(画面の説明文 index.html にも「photos」と書いている)
export const PHOTOS_DIR = 'photos';
// 写真のファイル名の名称部分と、保存の「名前」(出力フォルダ名・Excel・ZIP の名前)の最大文字数
// (長すぎるとパス全体が Windows の上限に近づくため)
export const MAX_NAME_LENGTH = 40;

// ===== Excel(excel-writer.js) =====

export const EXCEL_FONT_NAME = '游ゴシック';
export const EXCEL_FONT_SIZE = 11;
// 「写真」列(B列)の縮小版の表示枠(長辺の最大 px)と、画像の上下の余白(px)
export const EXCEL_THUMB_BOX_PX = 150;
export const EXCEL_THUMB_MARGIN_PX = 8;
// 「写真」列(B列)の幅(文字数)。25 で約175px になり、表示枠 150px の左右に約12px の余白ができる
export const EXCEL_PHOTO_COL_WIDTH = 25;
