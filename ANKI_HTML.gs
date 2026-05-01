function transferFormulas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // 5行目以降のデータがある行に対して処理
  if (lastRow < 5) {
    Logger.log('データがありません');
    return;
  }
  
  const startRow = 5;
  const numRows = lastRow - startRow + 1;

  // データ範囲を取得(A〜K列まで)
  const dataRange = sheet.getRange(startRow, 1, numRows, 11);
  const values = dataRange.getValues();

  // ★ L列(フラグ)を取得
  const lRange = sheet.getRange(startRow, 12, numRows, 1);
  const lValues = lRange.getValues();

  // 結果を格納する配列（M〜P）
  const results = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const flag = lValues[i][0]; // L列の値
    const rowNum = startRow + i; // 実際の行番号

    // ★ 追加対象でない行は何もしない（M〜Pも書き換えない）
    if (flag !== '追加対象') {
      // 既存値を維持するため、その行の現在の M〜P を読み直して results に入れる
      const existingMP = sheet.getRange(rowNum, 13, 1, 4).getValues()[0];
      results.push(existingMP);
      continue;
    }

    // M列: =B列 & A列
    const colM = row[1] + row[0]; // B列(index 1) + A列(index 0)
    
    // N列: ="【部屋】" & G列 & CHAR(10) & "【場所】" & H列
    const colN = "【部屋】" + row[6] + "\n" + "【場所】" + row[7]; // G列(index 6), H列(index 7)
    
    // O列: =D列 & CHAR(10) & CHAR(10) & E列
    const colO = row[3] + "\n\n" + row[4]; // D列(index 3), E列(index 4)
    
    // P列: ="【内装】"&$I5&CHAR(10)&"【物体】"&$J5&CHAR(10)&"────"&CHAR(10)&$K5
    const colP =
      "\n" +    
      "【内装】" + row[8] + "\n" +
      "【物体】" + row[9] + "\n" +
      "────" + "\n" +
      row[10];

    results.push([colM, colN, colO, colP]);

    // ★ 実行完了したら L列を「追加済み」に更新
    sheet.getRange(rowNum, 12).setValue('追加済み');
  }
  
  // M5〜P列に一括書き込み
  sheet.getRange(startRow, 13, results.length, 4).setValues(results);
  
  // ★ 転記完了後、自動的に次の処理を実行
  convertRowsAndExtractModel();
}


function convertRowsAndExtractModel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // 必要ならシート名指定

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return; // データ開始行が5行目想定

  const startRow = 5;
  const numRows = lastRow - startRow + 1;

  // ★ M〜P列 (M=13, 4列: M,N,O,P)
  const mToP = sheet.getRange(startRow, 13, numRows, 4).getValues();
  // ★ Q列既存値
  const existingQ = sheet.getRange(startRow, 17, numRows, 1).getValues();

  const outMP = [];
  const outQ = [];

  for (let i = 0; i < mToP.length; i++) {
    const row = mToP[i];
    const m = row[0];            // M
    const n = row[1];            // N
    const o = row[2];            // O
    const p = row[3];            // P (Markdown)

    // ★ M〜P が空行（未設定）の場合は何もせず既存値を維持
    if (!m && !n && !o && !p) {
      outMP.push(row);
      outQ.push(existingQ[i]);
      continue;
    }

    // 1) HTML 変換
    const mHtml = textToHtml(m);
    const nHtml = textToHtml(n);
    const oHtml = textToHtml(o);
    const pHtml = markdownToHtml(p); // 表対応のMarkdown→HTML

    // 2) 動画ファイル名の取得（★ M列の値を使う）
    const mValue = row[0];  // row[0] は M列（mToP の先頭列）
    let pHtmlWithSound = pHtml;
    if (mValue && String(mValue).trim() !== '') {
      // M列に「親切1」などのファイル名だけが入っている前提
      const soundTag = '[sound:' + String(mValue).trim() + '.mp4]';
      pHtmlWithSound = soundTag + '\n\n' + pHtml;
    }

    // ★ M〜P を上書き出力
    outMP.push([mHtml, nHtml, oHtml, pHtmlWithSound]);

    // 3) 模範英文抽出（P列のHTMLから）
    const modelText = extractModelTextFromHtml(pHtml);
    outQ.push([modelText]);
  }

  // ★ M〜P, Q に書き込み
  sheet.getRange(startRow, 13, numRows, 4).setValues(outMP); // M〜P
  sheet.getRange(startRow, 17, numRows, 1).setValues(outQ);   // Q

}



function textToHtml(text) {
  if (text == null) return '';
  let s = String(text);
  s = s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  s = s.replace(/\r\n|\r|\n/g, '<br>');
  return s;
}


// テーブル対応版 Markdown → HTML
function markdownToHtml(md) {
  if (md == null) return '';
  let s = String(md);

  // まずテーブル部分をHTMLに
  s = convertMarkdownTables(s);

  // 見出し
  s = s.replace(/^###### (.*)$/gm, '<h6>$1</h6>');
  s = s.replace(/^##### (.*)$/gm, '<h5>$1</h5>');
  s = s.replace(/^#### (.*)$/gm, '<h4>$1</h4>');
  s = s.replace(/^### (.*)$/gm, '<h3>$1</h3>');
  s = s.replace(/^## (.*)$/gm, '<h2>$1</h2>');
  s = s.replace(/^# (.*)$/gm, '<h1>$1</h1>');

  // 太字・斜体・インラインコード
  s = s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/\*(.+?)\*/g, '<em>$1</em>');
  s = s.replace(/`([^`]+)`/g, 'de>$1</code>');

  // 区切り線
  s = s.replace(/^\s*---\s*$/gm, '<hr>');

  // 改行 → <br>
  s = s.replace(/\r\n|\r|\n/g, '<br>');

  return s;
}


function convertMarkdownTables(text) {
  const lines = text.split(/\r\n|\r|\n/);
  let inTable = false;
  let tableLines = [];
  const outLines = [];

  function flushTable() {
    if (!inTable || tableLines.length === 0) return;
    if (tableLines.length >= 2) {
      const headerLine = tableLines[0];
      const separatorLine = tableLines[1];
      const bodyLines = tableLines.slice(2);

      const headers = splitMarkdownRow(headerLine);
      const isValidSeparator = /---/.test(separatorLine);

      if (isValidSeparator && headers.length > 0) {
        let html = '<table><thead><tr>';
        headers.forEach(h => {
          html += '<th>' + escapeHtml(h.trim()) + '</th>';
        });
        html += '</tr></thead>';

        if (bodyLines.length > 0) {
          html += '<tbody>';
          bodyLines.forEach(line => {
            const cells = splitMarkdownRow(line);
            if (cells.length === 0) return;
            html += '<tr>';
            cells.forEach(c => {
              html += '<td>' + escapeHtml(c.trim()) + '</td>';
            });
            html += '</tr>';
          });
          html += '</tbody>';
        }
        html += '</table>';
        outLines.push(html);
      } else {
        outLines.push(...tableLines);
      }
    } else {
      outLines.push(...tableLines);
    }
    inTable = false;
    tableLines = [];
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isTableRow = /^\s*\|.*\|\s*$/.test(line);

    if (isTableRow) {
      if (!inTable) {
        inTable = true;
        tableLines = [];
      }
      tableLines.push(line);
    } else {
      if (inTable) flushTable();
      outLines.push(line);
    }
  }
  if (inTable) flushTable();

  return outLines.join('\n');
}


function splitMarkdownRow(line) {
  let inner = line.trim();
  if (inner.startsWith('|')) inner = inner.slice(1);
  if (inner.endsWith('|')) inner = inner.slice(0, -1);
  const cells = inner.split('|').map(s => s.trim());
  if (cells.every(c => c === '')) return [];
  return cells;
}


function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


// 「✅ 模範英文」部分だけ抽出
function extractModelTextFromHtml(html) {
  let s = String(html);

  const marker = '<h2>✅ 模範英文</h2>';
  const idx = s.indexOf(marker);
  if (idx === -1) {
    return '';
  }

  s = s.slice(idx + marker.length);

  const hrIndex = s.indexOf('<hr>');
  if (hrIndex !== -1) {
    s = s.slice(0, hrIndex);
  }

  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.trim();

  return s;
}



/**
 * B〜D列が空でない行だけ、
 * セル内改行を <br> に変換しつつHTMLエスケープして
 * E〜G列に転記する（Ankiインポート用）。
 */
function copyBDtoEG_forAnki() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const startRow = 2; // データ開始行（必要に応じて変更）
  const numRows = lastRow - startRow + 1;

  // B〜D列(2〜4列目)を取得
  const srcRange = sheet.getRange(startRow, 2, numRows, 3);
  const srcValues = srcRange.getValues();

  const outValues = [];

  for (let i = 0; i < numRows; i++) {
    const b = srcValues[i][0];
    const c = srcValues[i][1];
    const d = srcValues[i][2];

    // B〜Dのどれか1つでも入っていたら対象
    const hasAny =
      (b !== '' && b != null) ||
      (c !== '' && c != null) ||
      (d !== '' && d != null);

    if (hasAny) {
      outValues.push([
        textToHtmlForAnki_(b),
        textToHtmlForAnki_(c),
        textToHtmlForAnki_(d),
      ]);
    } else {
      // すべて空 → E〜Gも空（既存維持したいならここを変える）
      outValues.push(['', '', '']);
    }
  }

  // E〜G列(5〜7列目)に一括書き込み
  const dstRange = sheet.getRange(startRow, 5, numRows, 3);
  dstRange.setValues(outValues);
}

/**
 * Anki向け：テキストをHTML用に変換
 * - 特殊文字エスケープ
 * - 改行を <br> に変換
 */
function textToHtmlForAnki_(value) {
  if (value == null) return '';
  let s = String(value);

  // 特殊文字エスケープ（タグとして解釈させたくない部分を守る）
  s = s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  // スプレッドシートの改行コードを <br> に変換（Ankiで改行として表示）[cite:10]
  s = s.replace(/\r\n|\r|\n/g, '<br>');

  return s;
}