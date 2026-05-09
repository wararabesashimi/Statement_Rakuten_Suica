function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    rakutenPay: ss.getSheetByName("利用明細"),
    rakutenCash: ss.getSheetByName("キャッシュ明細"),
    suicaCharge: ss.getSheetByName("Suicaチャージ明細")
  };
}

function fetchStatementByLabel() {
  const sheets = getSheets();
  const targetLabel = GmailApp.getUserLabelByName("未反映");
  const doneLabel = GmailApp.getUserLabelByName("反映済み");

  if (!targetLabel || !doneLabel) {
    console.log("ラベル「未反映」または「反映済み」が見つかりません。");
    return;
  }

  // 「未反映」ラベルがついているスレッドを取得
  const threads = targetLabel.getThreads();
  if (threads.length === 0) {
    console.log("対象のスレッドはありません。");
    return;
  }

  threads.forEach(thread => {
    // スレッドから1通1通のメールをリストで取り出す
    const messages = thread.getMessages();

    messages.forEach(message => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      let success = false;

      // 1. 楽天キャッシュ
      if (subject.includes("【楽天キャッシュ】チャージ完了のお知らせ")) {
        const dateMatch = body.match(/ご利用日時\s*([\d\/\(\)\s\u3000-\u30ff\u4e00-\u9faf\:]+?\d{1,2}:\d{2})/);
        const amountMatch = body.match(/\[金額\]\s*([\d,]+)\s*円/);
        if (dateMatch && amountMatch) {
          sheets.rakutenCash.appendRow([dateMatch[1].trim(), amountMatch[1].replace(/,/g, "").trim()]);
          success = true;
        }
      } 
      // 2. Suicaチャージ
      else if (subject.includes("Suica入金（チャージ）完了のお知らせ")) {
        const dateMatch = body.match(/ご利用日時\s*([\d\/\(\)\s\u3000-\u30ff\u4e00-\u9faf\:]+?\d{1,2}:\d{2})/);
        const amountMatch = body.match(/\[金額\]\s*([\d,]+)\s*円/);
        if (dateMatch && amountMatch) {
          sheets.suicaCharge.appendRow([dateMatch[1].trim(), amountMatch[1].replace(/,/g, "").trim()]);
          success = true;
        }
      } 
      // 3. 楽天ペイ（通常決済）
      else if("楽天ペイアプリご利用内容確認メール") {
        const dateMatch = body.match(/ご利用日時\s*([\d\/\(\)\s\u3000-\u30ff\u4e00-\u9faf\:]+?\d{1,2}:\d{2})/);
        const storeMatch = body.match(/ご利用店舗\s*([\s\S]+?)(?=伝票番号|電話番号|決済総額|ご利用金額)/);
        const amountMatch = body.match(/決済総額\s*[^\d]*([\d,]+)/);

        if (dateMatch && storeMatch && amountMatch) {
          sheets.rakutenPay.appendRow([
            dateMatch[1].replace(/\r?\n/g, "").trim(),
            storeMatch[1].replace(/\s+/g, " ").trim(),
            amountMatch[1].replace(/,/g, "").trim()
          ]);
          success = true;
        }
      } else {
        console.log("不明なメールです。");
      }
    });
      thread.addLabel(doneLabel);
      thread.removeLabel(targetLabel);
      console.log("スレッド処理完了: ラベルを移動しました。");
  });
}
