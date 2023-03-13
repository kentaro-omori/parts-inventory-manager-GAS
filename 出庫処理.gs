// @ts-nocheck
/**
 * 出庫処理をする（出荷商品の出荷数を在庫管理表から減算する）。
 */
function shukko(ss = SpreadsheetApp.getActiveSpreadsheet(), zaikoSheet = ss.getSheetByName('部品在庫マスタ'), shukkaSheet = ss.getSheetByName('基板作成数'), shukkaRirekiSheet = ss.getSheetByName('基板作成履歴')) {
  // 処理開始ダイアログ
  const question = Browser.msgBox("出庫します。在庫数が更新されます。", Browser.Buttons.OK_CANCEL);
  if (question == "cancel") {
    return;
  }

  const Z_ITEM_NAME = "型番";
  const Z_QUANTITY = "在庫数";
  const Z_THRESHOLD = "アラート閾値";
  const S_ITEM_NAME = "型番";
  const S_PCB_NAME = "完成基板名";
  const S_QUANTITY = "出荷数";
  const S_DATE = "出荷日";
  const P_ITEM_NAME = "型番";
  const P_QUANTITY = "員数";

  //「在庫管理マスタ」のデータ取得
  const zaikoData = zaikoSheet.getDataRange().getValues();
  const zaikoHeads = zaikoData.shift();
  const zItemColIdx = zaikoHeads.indexOf(Z_ITEM_NAME);
  const zQuantityColIdx = zaikoHeads.indexOf(Z_QUANTITY);
  const zThresholdColIdx = zaikoHeads.indexOf(Z_THRESHOLD);  
  const zLastRowNumber = zaikoSheet.getRange(zaikoSheet.getMaxRows(), zItemColIdx + 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //「出荷データ」のデータ取得
  const shukkaData = shukkaSheet.getDataRange().getValues();
  const shukkaHeads = shukkaData.shift();
  const sPcbColIdx = shukkaHeads.indexOf(S_PCB_NAME);
  const sQuantityColIdx = shukkaHeads.indexOf(S_QUANTITY);
  const sDateColIdx = shukkaHeads.indexOf(S_DATE);
  const sLastColNumber = shukkaSheet.getRange(1, shukkaSheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();

  //「基板作成履歴」の最終行取得
  const rLastRowNumber = shukkaRirekiSheet.getLastRow();


  // 出庫対象の出荷データを取得する
  let shukkoData = [];
  let shukkoRowNumbers = [];
  const today = new Date();
  for (let i = 0; i < shukkaData.length; i++) {
    const rowData = shukkaData[i].flat();
    /*
    // 出荷日が今日以前か判定
    if (rowData[sDateColIdx] != "" && rowData[sDateColIdx] <= today) {
      // 出荷数が0以上の数値か判定
      if (isFinite(rowData[sQuantityColIdx]) && rowData[sQuantityColIdx] >= 0) {
        // 出庫対象としてリストに追加する
        shukkoData.push(rowData);
        shukkoRowNumbers.push(i + 2);
      } else {
        continue;
      }
    }
    */
    // 出荷数が0以上かつ空白でないの数値か判定
    if (isFinite(rowData[sQuantityColIdx]) && rowData[sQuantityColIdx] > 0) {
      //日付入力が無ければ今日の日付を追加
      if(shukkaSheet.getRange(i + 2, 3).isBlank()){
        shukkaSheet.getRange(i + 2, 3).setValue(today);
        rowData.splice(2, 1, today)
      }
      // 出庫対象としてリストに追加する
      shukkoData.push(rowData);
      shukkoRowNumbers.push(i + 2);
    } else {
      continue;
    }
  }
  if (shukkoData.length == 0) {
    Browser.msgBox("出庫対象はありません。");
    console.log("出庫対象はありません。");
    return;
  }

  // 作成した基板の部品表を取得する
  let usePartsData = [];
  const zaikoShohin = getColData(zaikoData,zItemColIdx).flat();  //在庫商品データを取得する関数
  let message = ""; //slackでの送信用メッセージ
  //let partsListSheet;
  //let partsData;
  for (let i = 0; i < shukkoData.length; i++) {
    const partsListSheet = ss.getSheetByName(shukkoData[i][sPcbColIdx]);
    const partsData = partsListSheet.getDataRange().getValues();
    const partsHeads = partsData.shift();
    const pItemColIdx = partsHeads.indexOf(P_ITEM_NAME);
    const pQuantityColIdx = partsHeads.indexOf(P_QUANTITY);
    message = message + shukkoData[i][sPcbColIdx] + "を" + shukkoData[i][sQuantityColIdx] + "個納品しました。\n";
    for (let j = 0; j < partsData.length; j++) {
      if (zaikoShohin.indexOf(partsData[j][pItemColIdx]) != -1) { //出庫商品が在庫商品にある場合(無い場合は未実装)
        const idx = zaikoShohin.indexOf(partsData[j][pItemColIdx]); //出庫商品のインデックスを取得する
        const zQuantityBefore = zaikoData[idx][zQuantityColIdx]; //出庫前の在庫数量を取得する
        zaikoData[idx][zQuantityColIdx] = zQuantityBefore - partsData[j][pQuantityColIdx] * shukkoData[i][sQuantityColIdx];  //在庫数量から出庫数量を引く
        console.log(zaikoData[idx][zItemColIdx] + "の在庫数：" + zaikoData[idx][zQuantityColIdx] + "（現在庫" + zQuantityBefore + " - 出庫数" + partsData[j][pQuantityColIdx] * shukkoData[i][sQuantityColIdx] + "）"); //ログに在庫数量を表示する      
     }
    }
  }

  console.log(message);

/*
  // 出庫する（対象商品の出荷数を在庫管理表に減算する）
  let shinShohinData = [];  //新商品データを格納する配列
  const zaikoShohin = getColData(zaikoData,zItemColIdx).flat();  //在庫商品データを取得する関数
  for (let i = 0; i < shukkoData.length; i++) {
    if (zaikoShohin.indexOf(shukkoData[i][sPcbColIdx]) != -1) { //出庫商品が在庫商品にある場合
      const idx = zaikoShohin.indexOf(shukkoData[i][sPcbColIdx]); //出庫商品のインデックスを取得する
      const zQuantityBefore = zaikoData[idx][zQuantityColIdx]; //出庫前の在庫数量を取得する
      zaikoData[idx][zQuantityColIdx] = zQuantityBefore - shukkoData[i][sQuantityColIdx];  //出庫前の在庫数量を取得する
      console.log(shukkoData[i][sPcbColIdx] + "の在庫数：" + zaikoData[idx][zQuantityColIdx] + "（現在庫" + zQuantityBefore + " - 出庫数" + shukkoData[i][sQuantityColIdx] + "）"); //ログに在庫数量を表示する
    } else { //出庫商品が在庫商品にない場合（新商品）
      console.log("新商品：" + shukkoData[i][sPcbColIdx] + "の在庫数：" + shukkoData[i][sQuantityColIdx]); //ログに新商品と在庫数量を表示する
      shinShohinData.push(shukkoData[i]); //新商品データに追加する
    }
  }
*/
  // 在庫管理表を更新する
  if (zaikoData.length != 0) { //在庫データが空でない場合
  /* 
  在庫管理シートに以下の処理を行う。
  - 商品名列に在庫商品名データをセットする。
  - 数量列に在庫数量データをセットする。
  */
    zaikoSheet.getRange(2, zItemColIdx + 1, zaikoData.length, 1).setValues(getColData(zaikoData, zItemColIdx));
    zaikoSheet.getRange(2, zQuantityColIdx + 1, zaikoData.length, 1).setValues(getColData(zaikoData, zQuantityColIdx));
  }

  // 出庫対象を削除する（削除前に出荷履歴に転記）
  shukkaRirekiSheet.getRange(rLastRowNumber + 1, 1, shukkoData.length, sLastColNumber).setValues(shukkoData);
  const reversedShukkoRowNumbers = shukkoRowNumbers.reverse();
//  for (let i = 0; i < reversedShukkoRowNumbers.length; i++) {
//    shukkaSheet.deleteRow(reversedShukkoRowNumbers[i]);
//  }
  const shukkaLastRowNumber = shukkaSheet.getLastRow();
  let addRange = shukkaSheet.getRange(2, 2, shukkaLastRowNumber, 2);
  addRange.clearContent();

  //納品通知をslackに送信
  //SlackAPIで登録したボットのトークンを設定する
  const token = "xoxb-************";
  //ライブラリから導入したSlackAppを定義し、トークンを設定する
  const slackApp = SlackApp.create(token);
  //Slackボットがメッセージを投稿するチャンネルを定義する
  const channelId = "#general";
  //Slackボットが投稿するメッセージを定義する
  //const message = "SlackボットによるGASからの投稿テストです。"
  //SlackAppオブジェクトのpostMessageメソッドでボット投稿を行う
  slackApp.postMessage(channelId, message);

  //在庫数がアラート閾値を下回っていたらメール送信
  zaikoHeads.splice(4, 3)
  let fewPartsData = [zaikoHeads];
  for (let i = 0; i < zaikoData.length; i++) {
    if(zaikoData[i][zQuantityColIdx] <= zaikoData[i][zThresholdColIdx]) {
      zaikoData[i].splice(4, 3)
      fewPartsData.push(zaikoData[i]);
      console.log(zaikoData[i][zItemColIdx] + "の在庫数：" + zaikoData[i][zQuantityColIdx] + "が閾値" + zaikoData[i][zThresholdColIdx] + "を下回っている");
    }
  } 
  //配列をhtmlの表形式に変換
  // Start the table tag
  var html = "<table>";
  // Loop through each row of the array
  for (var i = 0; i < fewPartsData.length; i++) {
    // Start the table row tag
    html += "<tr>";
    // Loop through each cell of the row
    for (var j = 0; j < fewPartsData[i].length; j++) {
      // Add the table cell tag with the array value
      html += "<td>" + fewPartsData[i][j] + "</td>";
    }
    // End the table row tag
    html += "</tr>";
  }
  // End the table tag
  html += "</table>";
  console.log(html);
  //メール送信
  var email = "your e-mail address";
  //本文
  var body = "以下の部品在庫が少なくなっています。<br>部品の発注を検討してください。<br><br>" + html;
  //送信者の名前
  var sender = "在庫管理システム";
  //メールの件名
  var subject = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd') + " 電子部品在庫情報"; 
  //オプション
  var options = {
  "htmlBody":body,
  "name" : sender
  };
  GmailApp.sendEmail(email, subject, body, options)

  // 処理終了ダイアログ
  Browser.msgBox("出庫が完了しました。");
  console.log("出庫が完了しました。");
}