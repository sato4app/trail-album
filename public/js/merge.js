// 読み込んだ KMZ・GPX をまとめる(後から読み込んだファイルは、今の内容に追加する)
// 同じ写真・同じ経路を重ねて追加しないよう、重複を見分けて除く。
// minoh-hiking は出力のたびに、その時点の全部の写真・経路を書き出すため、同じ日に2回出力した
// KMZ・GPX には同じ写真・経路が入る(経路は、記録の途中で出力すると途中までの形で入る)。

const unique = (list) => [...new Set(list)];

// 時刻の昇順。時刻の無い(null)ものは後ろに置く。同じとき(どちらも null を含む)は 0
function compareTime(a, b) {
  if (a !== null && b !== null) return a - b;
  if (a === null && b !== null) return 1;
  if (a !== null && b === null) return -1;
  return 0;
}

// ===== KMZ(写真) =====

// 写真の見分け: 撮影日時・位置・ファイルの大きさが同じなら同じ写真とみなす
// (名前は minoh-hiking で後から変えられるため、見分けに使わない)
function photoKey(photo) {
  return `${photo.timeMs}|${photo.lat.toFixed(6)}|${photo.lng.toFixed(6)}|${photo.bytes.length}`;
}

// 新しく読み込んだ写真(photos)から、読み込み済みの KMZ(loadedFiles)の写真と同じものを除いて返す。
// 同じ KMZ の中で重なっている写真も1枚にする
export function excludeLoadedPhotos(loadedFiles, photos) {
  const seen = new Set(loadedFiles.flatMap((file) => file.photos.map(photoKey)));
  return photos.filter((photo) => {
    const key = photoKey(photo);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// 読み込んだ KMZ(kmzFiles。読み込んだ順。どれも写真を1枚以上持つ)を1つのアルバムにまとめる。
// 何も無ければ null。
//   { files, titles, fileName, title, photos }
// - photos は出力順(撮影日時の昇順。撮影日時の無い写真は、読み込んだ順・KML の順で末尾)
// - files は、いちばん早く撮った写真の順。先頭の KMZ の名前を、出力フォルダ名と Excel のファイル名に使う
//   (fileName・title は files[0] のもの)
// - titles は KML の Document 名(重なりと空を除く)
export function buildAlbum(kmzFiles) {
  if (kmzFiles.length === 0) return null;
  const entries = kmzFiles.flatMap((file, order) => file.photos.map((photo) => ({ photo, order })));
  entries.sort((a, b) =>
    compareTime(a.photo.timeMs, b.photo.timeMs) || a.order - b.order || a.photo.kmlIndex - b.photo.kmlIndex
  );
  const files = unique(entries.map((e) => kmzFiles[e.order]));
  return {
    files,
    titles: unique(files.map((f) => f.title).filter(Boolean)),
    fileName: files[0].fileName,
    title: files[0].title,
    photos: entries.map((e) => e.photo)
  };
}

// ===== GPX(移動経路) =====

// 経路の見分け: 最初の記録地点(記録時刻・位置)が同じなら同じ経路とみなす
// (記録の途中で出力した GPX の経路は、後で出力した GPX の同じ経路の途中までと同じになるため)
function segmentKey(segment) {
  const p = segment.points[0];
  return `${p.timeMs}|${p.lat.toFixed(6)}|${p.lng.toFixed(6)}`;
}

// 読み込み済みの経路(loaded)に、新しく読み込んだ GPX(track: gpx-reader.js の readGpx の結果)の
// 経路を加える。同じ経路は記録地点の多いほう(より後まで記録したもの)を残す。
// loaded は変えずに、新しい配列を返す。
//   { segments, added, replaced, skipped }
//   added: 新しく加えた経路 / replaced: 読み込み済みの同じ経路と置き換えた経路 / skipped: 加えなかった経路
// 各経路には、読み込んだ GPX(source: track。同じ名前の別のファイルとも見分けられるよう、ファイル名ではなく
// 読み込んだもの自体)を付ける。
export function addSegments(loaded, track) {
  const segments = [...loaded];
  const indexByKey = new Map(segments.map((s, i) => [segmentKey(s), i]));
  const result = { segments, added: [], replaced: [], skipped: [] };
  for (const seg of track.segments) {
    const segment = { ...seg, source: track };
    const key = segmentKey(segment);
    const i = indexByKey.get(key);
    if (i === undefined) {
      indexByKey.set(key, segments.length);
      segments.push(segment);
      result.added.push(segment);
    } else if (segment.points.length > segments[i].points.length) {
      segments[i] = segment;
      result.replaced.push(segment);
    } else {
      result.skipped.push(segment);
    }
  }
  return result;
}

// 読み込んだ経路(loaded。読み込んだ順)を GPX ごとにまとめる。何も無ければ null。
//   { files: [{ source, segments }], segmentCount, pointCount, distanceM }
// - files は、いちばん早く記録した経路の順(画面の 1, 2 … の順)。source は readGpx の結果
// - 各 GPX の segments は記録の日時の順(記録時刻の無い経路は、読み込んだ順で末尾)
export function buildTrack(loaded) {
  if (loaded.length === 0) return null;
  const sorted = loaded
    .map((segment, order) => ({ segment, order }))
    .sort((a, b) => compareTime(a.segment.startMs, b.segment.startMs) || a.order - b.order)
    .map((e) => e.segment);
  const files = unique(sorted.map((s) => s.source)).map((source) => ({
    source,
    segments: sorted.filter((s) => s.source === source)
  }));
  return {
    files,
    segmentCount: sorted.length,
    pointCount: sorted.reduce((sum, s) => sum + s.points.length, 0),
    distanceM: sorted.reduce((sum, s) => sum + s.distanceM, 0)
  };
}
