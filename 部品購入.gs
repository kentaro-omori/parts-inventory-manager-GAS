/**
 * 部品追加処理をする（部品の購入数を在庫管理表に加算する）。
 */
function kounyu(ss = SpreadsheetApp.getActiveSpreadsheet(), zaikoSheet = ss.getSheetByName('部品在庫マスタ'), tsuikaSheet = ss.getSheetByName('部品購入'), tsuikaRirekiSheet = ss.getSheetByName('部品購入履歴')) {
  // 処理開始ダイアログ
  const question = Browser.msgBox("購入部品を追加します。在庫数が更新されます。", Browser.Buttons.OK_CANCEL);
  if (question == "cancel") {
    return;
  }

  const Z_ITEM_NAME = "型番";
  const Z_QUANTITY = "在庫数";
  const N_ITEM_NAME = "型番";
  const N_QUANTITY = "購入数";
  const N_DATE = "購入日";

  //「在庫管理マスタ」のデータ取得
  const zaikoData = zaikoSheet.getDataRange().getValues();
  const zaikoHeads = zaikoData.shift();
  const zItemColIdx = zaikoHeads.indexOf(Z_ITEM_NAME);
  const zQuantityColIdx = zaikoHeads.indexOf(Z_QUANTITY);
  const zLastRowNumber = zaikoSheet.getRange(zaikoSheet.getMaxRows(), zItemColIdx + 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //「部品購入」のデータ取得
  const tsuikaData = tsuikaSheet.getDataRange().getValues();
  const tsuikaHeads = tsuikaData.shift();
  const nItemColIdx = tsuikaHeads.indexOf(N_ITEM_NAME);
  const nQuantityColIdx = tsuikaHeads.indexOf(N_QUANTITY);
  const nDateColIdx = tsuikaHeads.indexOf(N_DATE);
  const nLastColNumber = tsuikaSheet.getRange(1, tsuikaSheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();

  //「部品購入履歴」の最終行取得
  const rLastRowNumber = tsuikaRirekiSheet.getLastRow();


  // 入庫対象の部品購入データを取得する
  let kounyuData = [];
  let kounyuRowNumbers = [];
  const today = new Date();
  for (let i = 0; i < tsuikaData.length; i++) {
    const rowData = tsuikaData[i].flat();
/*
    // 購入日が今日以前か判定
    if (rowData[nDateColIdx] != "" && rowData[nDateColIdx] <= today) {
      // 購入数が0以上の数値か判定
      if (isFinite(rowData[nQuantityColIdx]) && rowData[nQuantityColIdx] >= 0) {
        // 追加対象としてリストに追加する
        kounyuData.push(rowData);
        kounyuRowNumbers.push(i + 2);
      } else {
        continue;
      }
    }
*/
    // 購入数が0以上の数値か判定
    if (isFinite(rowData[nQuantityColIdx]) && rowData[nQuantityColIdx] > 0) {
      //日付入力が無ければ今日の日付を追加
      if(tsuikaSheet.getRange(i + 2, 5).isBlank()){
        tsuikaSheet.getRange(i + 2, 5).setValue(today);
        rowData.splice(4, 1, today)
      }
      // 追加対象としてリストに追加する
      kounyuData.push(rowData);
      kounyuRowNumbers.push(i + 2);
    } else {
      continue;
    }
  }

  if (kounyuData.length == 0) {
    Browser.msgBox("追加対象はありません。");
    console.log("追加対象はありません。");
    return;
  }


  // 追加する（対象商品の購入数を在庫管理マスタに加算する）
  let shinShohinData = [];
  const zaikoShohin = getColData(zaikoData,zItemColIdx).flat();
  for (let i = 0; i < kounyuData.length; i++) {
    if (zaikoShohin.indexOf(kounyuData[i][nItemColIdx]) != -1) {
      const idx = zaikoShohin.indexOf(kounyuData[i][nItemColIdx]);
      const zQuantityBefore = zaikoData[idx][zQuantityColIdx];
      zaikoData[idx][zQuantityColIdx] = zQuantityBefore + kounyuData[i][nQuantityColIdx];
      console.log(kounyuData[i][nItemColIdx] + "の在庫数：" + zaikoData[idx][zQuantityColIdx] + "（現在庫" + zQuantityBefore + " + 入庫数" + kounyuData[i][nQuantityColIdx] + "）");
    } else {
      console.log("新商品：" + kounyuData[i][nItemColIdx] + "の在庫数：" + kounyuData[i][nQuantityColIdx]);
      shinShohinData.push(kounyuData[i]);
    }
  }
  // 在庫管理マスタを更新する
  if (zaikoData.length != 0) {
    zaikoSheet.getRange(2, zItemColIdx + 1, zaikoData.length, 1).setValues(getColData(zaikoData, zItemColIdx));
    zaikoSheet.getRange(2, zQuantityColIdx + 1, zaikoData.length, 1).setValues(getColData(zaikoData, zQuantityColIdx));
  }
  // 新商品データがあれば在庫管理マスタに追加する
  if (shinShohinData.length != 0) {
    zaikoSheet.getRange(zLastRowNumber + 1, zItemColIdx + 1, shinShohinData.length, 1).setValues(getColData(shinShohinData, nItemColIdx));
    zaikoSheet.getRange(zLastRowNumber + 1, zQuantityColIdx + 1, shinShohinData.length, 1).setValues(getColData(shinShohinData, nQuantityColIdx));
  }

  // 追加対象を削除する（削除前に部品購入履歴に転記）
  tsuikaRirekiSheet.getRange(rLastRowNumber + 1, 1, kounyuData.length, nLastColNumber).setValues(kounyuData);
  const reversedKounyuRowNumbers = kounyuRowNumbers.reverse();
//  for (let i = 0; i < reversedKounyuRowNumbers.length; i++) {
//    tsuikaSheet.deleteRow(reversedKounyuRowNumbers[i]);
//  }
  const tsuikaLastRowNumber = tsuikaSheet.getLastRow();
  let addRange = tsuikaSheet.getRange(2, 4, tsuikaLastRowNumber, 2);
  addRange.clearContent();

  // 処理終了ダイアログ
  Browser.msgBox("部品追加が完了しました。");
  console.log("部品追加が完了しました。");
}