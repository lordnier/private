// メイン：選択範囲のMarkdownをHTMLにしてセルに書き戻す
function convertSelectedRangeToHtml() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange(); // ユーザーが選択している範囲[web:7]
  if (!range) {
    SpreadsheetApp.getUi().alert('セル範囲を選択してください。');
    return;
  }

  const values = range.getValues();
  const converted = values.map(row =>
    row.map(cell => {
      if (typeof cell !== 'string') return cell;
      const trimmed = cell.trim();
      if (trimmed === '') return cell;
      return markdownToHtml(trimmed);
    })
  );

  range.setValues(converted);
  exportSelectedRangeAsCSV();
}

// 簡易Markdown→HTML変換
function markdownToHtml(md) {
  // まず行ごとに処理
  const lines = md.replace(/\r\n/g, '\n').split('\n');

  const htmlLines = [];
  let inCodeBlock = false;
  let codeBuffer = [];
  let tableBuffer = [];

  const flushCodeBlock = () => {
    if (codeBuffer.length > 0) {
      const escaped = escapeHtml(codeBuffer.join('\n'));
      htmlLines.push('<pre><code>' + escaped + '</code></pre>');
      codeBuffer = [];
    }
  };

  const flushTable = () => {
    if (tableBuffer.length > 0) {
      htmlLines.push(convertMarkdownTableToHtml(tableBuffer));
      tableBuffer = [];
    }
  };

  for (let line of lines) {
    // コードブロック ``` の判定
    if (/^```/.test(line)) {
      if (!inCodeBlock) {
        // コードブロック開始
        flushTable();
        inCodeBlock = true;
        codeBuffer = [];
      } else {
        // コードブロック終了
        inCodeBlock = false;
        flushCodeBlock();
      }
      continue;
    }

    if (inCodeBlock) {
      codeBuffer.push(line);
      continue;
    }

    // Markdownの表行かどうか（|で始まり|を含む行）
    if (/^\s*\|.*\|\s*$/.test(line)) {
      tableBuffer.push(line);
      continue;
    } else {
      // 表が終わったタイミングでフラッシュ
      flushTable();
    }

    const trimmed = line.trim();

    if (trimmed === '') {
      // 空行 → 段落の区切りとして<br>にしておく（好みで<p>などに変更可）
      htmlLines.push('<br>');
      continue;
    }

    // 見出し #, ##, ### ...
    const headingMatch = trimmed.match(/^(#{1,6})\s+(.*)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const content = inlineMarkdownToHtml(headingMatch[2]);
      htmlLines.push('<h' + level + '>' + content + '</h' + level + '>');
      continue;
    }

    // 箇条書き（ul）: -, * など
    const ulMatch = trimmed.match(/^[-*+]\s+(.*)$/);
    if (ulMatch) {
      // ひとまず1行ずつ<li>として扱い、後で< ul >でラップする
      htmlLines.push('<ul><li>' + inlineMarkdownToHtml(ulMatch[1]) + '</li></ul>');
      continue;
    }

    // 番号付きリスト（ol）: 1. など
    const olMatch = trimmed.match(/^\d+\.\s+(.*)$/);
    if (olMatch) {
      htmlLines.push('<ol><li>' + inlineMarkdownToHtml(olMatch[1]) + '</li></ol>');
      continue;
    }

    // それ以外は通常の段落扱い
    htmlLines.push('<p>' + inlineMarkdownToHtml(trimmed) + '</p>');
  }

  // 最後に残っているバッファをフラッシュ
  if (inCodeBlock) {
    flushCodeBlock();
  }
  flushTable();

  // 連続した<ul>や<ol>をまとめる（簡易処理）
  let html = htmlLines.join('\n');
  html = mergeListTags(html, 'ul');
  html = mergeListTags(html, 'ol');

  return html;
}

// インライン要素のMarkdown→HTML (**太字**, *斜体*, `code`, リンクなど)
function inlineMarkdownToHtml(text) {
  let result = escapeHtml(text);

  // 強調（太字） **text**
  result = result.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

  // 斜体 *text*
  result = result.replace(/(^|[^\*])\*(?!\s)(.+?)(?!\s)\*(?!\*)/g, '$1<em>$2</em>');

  // インラインコード `code`
  result = result.replace(/`([^`]+)`/g, '<code>$1</code>');

  // リンク [text](url)
  result = result.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');

  // 改行（セル内の残りの\n）→<br>
  result = result.replace(/\n/g, '<br>');

  return result;
}

// Markdownの表をHTMLテーブルに変換
function convertMarkdownTableToHtml(lines) {
  // 先頭行: ヘッダ、2行目: 区切り、以降: 本文
  if (lines.length < 2) {
    return '<p>' + escapeHtml(lines.join('\n')) + '</p>';
  }

  const rows = lines.map(l =>
    l.trim()
      .replace(/^\|/, '')
      .replace(/\|$/, '')
      .split('|')
      .map(c => c.trim())
  );

  const header = rows[0];
  const body = rows.slice(2); // 2行目は --- | --- の区切り行とみなしてスキップ

  let html = '<table border="1" cellspacing="0" cellpadding="4">\n<thead><tr>';
  header.forEach(h => {
    html += '<th>' + inlineMarkdownToHtml(h) + '</th>';
  });
  html += '</tr></thead>\n<tbody>';

  body.forEach(row => {
    if (row.length === 1 && row[0] === '') return;
    html += '<tr>';
    row.forEach(cell => {
      html += '<td>' + inlineMarkdownToHtml(cell) + '</td>';
    });
    html += '</tr>';
  });

  html += '</tbody>\n</table>';
  return html;
}

// & < > などのエスケープ
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// <ul>や<ol>の連続を1つのリストにまとめる
function mergeListTags(html, tagName) {
  const openTag = '<' + tagName + '>';
  const closeTag = '</' + tagName + '>';

  // </ul>\n<ul> のような連続を削除して1つにまとめる
  const pattern = new RegExp(closeTag + '\\s*' + openTag, 'g');
  return html.replace(pattern, '');
}


function exportSelectedRangeAsCSV() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  
  // CSV形式に変換
  let csvContent = values.map(row => 
    row.map(cell => {
      // カンマやダブルクォートを含む場合はエスケープ
      let cellStr = cell.toString();
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        cellStr = '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',')
  ).join('\n');
  
  // ダウンロード
  const blob = Utilities.newBlob(csvContent, 'text/csv', 'selected_range.csv');
  const url = DriveApp.createFile(blob).getDownloadUrl().replace('?e=download&gd=true', '&export=download');
  
  const html = '<a href="' + url + '" target="_blank">CSVをダウンロード</a>';
  const ui = HtmlService.createHtmlOutput(html).setWidth(200).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(ui, 'CSV出力');
}

