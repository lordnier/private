
function convertRowsAndExtractModel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // 必要ならシート名指定

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return; // データ開始行が5行目想定

  const startRow = 5;
  const numRows = lastRow - startRow + 1;

  // G〜K列 (G=7, 5列: G,H,I,J,K)
  const gToK = sheet.getRange(startRow, 7, numRows, 5).getValues();
  // L〜O列既存値（追加対象じゃない行の保持用）
  const existingLO = sheet.getRange(startRow, 12, numRows, 4).getValues();
  // P列既存値
  const existingP = sheet.getRange(startRow, 16, numRows, 1).getValues();

  const outLO = [];
  const outP = [];

  for (let i = 0; i < gToK.length; i++) {
    const row = gToK[i];
    const flag = row[0];   // G
    const h = row[1];      // H
    const iCol = row[2];   // I
    const j = row[3];      // J
    const k = row[4];      // K (Markdown)

    if (flag === '追加対象') {
      // 1) HTML 変換
      const hHtml = textToHtml(h);
      const iHtml = textToHtml(iCol);
      const jHtml = textToHtml(j);
      const kHtml = markdownToHtml(k); // 表対応のMarkdown→HTML

      outLO.push([hHtml, iHtml, jHtml, kHtml]);

      // 2) 模範英文抽出
      const modelText = extractModelTextFromHtml(kHtml);
      outP.push([modelText]);

      // 3) この行の G列を「追加済み」に更新
      const rowIndex = startRow + i; // 実際の行番号
      sheet.getRange(rowIndex, 7).setValue('追加済み'); // 7列目 = G列
    } else {
      // 追加対象でない行は既存の値を維持
      outLO.push(existingLO[i]);
      outP.push(existingP[i]);
    }
  }

  // L〜O, P に書き込み
  sheet.getRange(startRow, 12, numRows, 4).setValues(outLO); // L〜O
  sheet.getRange(startRow, 16, numRows, 1).setValues(outP);   // P
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
  s = s.replace(/`([^`]+)`/g, '<code>$1</code>');

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

// 既存の名前で呼ばれても動くようにラッパーを作る
function convertRowsToHtml() {
  convertRowsAndExtractModel();
}