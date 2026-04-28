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