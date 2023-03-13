/**
 * 指定の列のデータを返す（指定の列を縦方向にデータを取得したいときに使う）。
 * @param {Array} array 2次元配列
 * @param {Number} colIdx 列番号インデックス
 * @return {Array} colData 指定列データ（2次元配列）
 */
function getColData(array, colIdx) {
  let colData = [];
  for (let i = 0; i < array.length; i++) {
    colData.push([array[i][colIdx]]);
  }
  return colData;
}