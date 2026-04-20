function importAnkiIdQuestionAnswer() {
  const spreadsheetId = '1uxufxlIikKVAcZ0Zpa4-Y-MtBAYFXPy0zSgsCCH2rNw';   // スプレッドシートID
  const sheetName     = 'ANKI_QA';               // 出力先シート名
  const folderId      = '1NTCCwPdLeXBP5x_LowoCtFB8MdkmTOnn';  // txt置き場フォルダID
  const fileName      = 'anki_export.txt';       // Ankiのエクスポートファイル名

  // Ankiの「どのフィールドを何にするか」指定
  const idColIndex = 0; // 0始まりで「IDのフィールド番号」
  const qColIndex  = 1; // 問題
  const aColIndex  = 2; // 答え

  const ss    = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  const folder = DriveApp.getFolderById(folderId);
  const files  = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    throw new Error('指定フォルダに ' + fileName + ' が見つからない');
  }
  const file = files.next();
  const text = file.getBlob().getDataAsString('UTF-8');

  // ヘッダー行 (#separator:tab など) を除外
  const rawLines  = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  const dataLines = rawLines.filter(l => !l.startsWith('#'));

  function stripHtml(html) {
    // HTML として解釈 → タグ削除 → 空白整理
    const out = HtmlService.createHtmlOutput(html).getContent();
    return out
      .replace(/<[^>]*>/g, ' ') // タグをスペースに置換
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const rows = [];
  dataLines.forEach(line => {
    const fields = line.split('\t');
    // フィールド数チェック（足りない行はスキップ）
    if (fields.length <= Math.max(idColIndex, qColIndex, aColIndex)) return;

    const id  = stripHtml(fields[idColIndex] || '');
    const q   = stripHtml(fields[qColIndex]  || '');
    const ans = stripHtml(fields[aColIndex]  || '');

    // 中身が全部空ならスキップ
    if (!id && !q && !ans) return;

    rows.push([id, q, ans]);
  });

  // シートに書き込み
  sheet.clear();
  sheet.appendRow(['ID', '問題', '答え']);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}