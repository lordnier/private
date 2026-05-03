function fillForAnki() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;          // データ開始行が5行目想定

  const startRow = 5;
  const numRows  = lastRow - startRow + 1;

  // A〜L列(1〜12列目) と M列(13列目: ステータス) を取得
  const values = sheet.getRange(startRow, 1, numRows, 12).getValues();  // A〜L
  const flags  = sheet.getRange(startRow, 13, numRows, 1).getValues();  // M

  // 既存の N〜V 列を取得しておき、非対象行は保持
  const existingNV = sheet.getRange(startRow, 14, numRows, 9).getValues(); // N〜V (14〜22列目)

  const outNV = [];

  for (let i = 0; i < numRows; i++) {
    const flag = (flags[i][0] || '').toString();
    const row  = values[i];  // A〜L が入っている1行分

    if (flag === '追加対象') {
      const A = row[0] || '';  // A列
      const B = row[1] || '';  // B列
      const D = row[3] || '';  // D列 ★追加
      const E = row[4] || '';  // E列 ★追加
      const G = row[6] || '';  // G列
      const H = row[7] || '';  // H列
      const I = row[8] || '';  // I列
      const J = row[9] || '';  // J列
      const K = row[10] || ''; // K列
      const L = row[11] || ''; // L列

      // N列: B列とA列を先頭から連結（例: 親切 + 1 → 親切1）
      const colN = String(B) + String(A);

      // O〜S列: G〜K列をそのままコピー
      const colO = G; // G → O
      const colP = H; // H → P
      const colQ = I; // I → Q
      const colR = J; // J → R
      const colS = K; // K → S

      // ★ T列: D列 & 2つの改行 & E列
      const colT = String(D) + '\n\n' + String(E);

      // U列: [sound: + N列 + .mp4]
      const colU = '[sound:' + colN + '.mp4]';

      // V列: L列の内容
      const colV = L;

      // N〜V の 9 列ぶんを配列で追加
      outNV.push([colN, colO, colP, colQ, colR, colS, colT, colU, colV]);
    } else {
      // 追加対象でない行は既存の N〜V を維持
      outNV.push(existingNV[i]);
    }
  }

  // N〜V 列に一括書き込み
  sheet.getRange(startRow, 14, numRows, 9).setValues(outNV);
  extractModelSentencesToColumnW();
}



function extractModelSentencesToColumnW() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 5) {
    SpreadsheetApp.getUi().alert('データが見つかりません');
    return;
  }
  
  const columnLData = sheet.getRange(5, 12, lastRow - 4, 1).getValues(); // 5行目から開始
  const columnWRange = sheet.getRange(5, 23, lastRow - 4, 1); // W列
  
  const writeData = [];
  let count = 0;
  
  for (let i = 0; i < columnLData.length; i++) {
    const cellValue = String(columnLData[i][0]);
    
    // バックスラッシュを削除して正規化
    const normalized = cellValue.replace(/\\/g, '');
    
    // マーカー検出
    if (normalized.includes('模範英文') || normalized.includes('模範A文')) {
      
      // 英文を抽出(同じセル内から)
      const lines = cellValue.split('\n');
      let sentences = [];
      let foundMarker = false;
      
      for (let line of lines) {
        const cleanLine = line.trim();
        
        // マーカーを見つけた
        if (cleanLine.includes('模範英文') || cleanLine.includes('模範A文')) {
          foundMarker = true;
          continue;
        }
        
        // マーカー後、次のセクションまで
        if (foundMarker) {
          // 区切り文字で終了
          if (cleanLine.includes('---') || 
              cleanLine.includes('🔍') ||
              cleanLine.includes('日本語') ||
              (cleanLine.match(/^#{1,3}\s/) && !cleanLine.includes('模範'))) {
            break;
          }
          
          // 英文を収集
          if (/[a-zA-Z]/.test(cleanLine) && cleanLine.length > 5) {
            sentences.push(cleanLine);
          }
        }
      }
      
      // 抽出した英文をW列に書き込み
      if (sentences.length > 0) {
        const combinedText = sentences.join('\n');
        writeData.push([combinedText]);
        count++;
      } else {
        writeData.push(['']);
      }
    } else {
      writeData.push(['']);
    }
  }
  
  // W列に一括書き込み
  if (writeData.length > 0) {
    columnWRange.setValues(writeData);
  }
  
}