/**
 * 料理ログ解析のメイン処理
 */
function runCookingLogAnalysis() {
  const config = loadConfig_();
  const GEMINI_API_KEY = config.GEMINI_API_KEY;
  const MODEL_NAME = config.MODEL_NAME;
  const DIARY_SHEET_NAME = config.DIARY_SHEET;
  const COOK_EXP_SHEET_NAME = config.COOK_EXP_SHEET;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diarySheet = ss.getSheetByName(DIARY_SHEET_NAME);
  const cookExpSheet = ss.getSheetByName(COOK_EXP_SHEET_NAME);

  if (!diarySheet || !cookExpSheet) {
    SpreadsheetApp.getUi().alert('シート名が見つかりません：' + DIARY_SHEET_NAME + " または " + COOK_EXP_SHEET_NAME);
    return;
  }

  const lastRow = diarySheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('料理データがありません。');
    return;
  }
  
  // A列(1)からE列(5)までを取得
  const diaryData = diarySheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const targets = [];
  for (let i = 0; i < diaryData.length; i++) {
    const rowNum = i + 2;
    const date = diaryData[i][0];
    const cookMemo = diaryData[i][3]; // D列：料理ログ
    const status = String(diaryData[i][4]).trim(); // E列：ステータス

    // D列に内容があり、かつE列が空の場合のみ処理
    if (cookMemo && status === "") {
      const dateStr = date instanceof Date ? Utilities.formatDate(date, "JST", "yyyy-MM-dd") : String(date);
      const combinedInput = `【料理ログ】\n${cookMemo}`;
      targets.push({ rowNum, dateStr, combinedInput });
    }
  }

  if (targets.length === 0) {
    SpreadsheetApp.getUi().alert('対象の未処理料理ログ（D列記入・E列空欄）が見つかりませんでした。');
    return;
  }

  let totalAdded = 0;
  const allResults = [];
  const errorLogs = [];
  
  SpreadsheetApp.getActive().toast(`${targets.length}件を解析中...`, '🍳');

  for (let i = 0; i < targets.length; i++) {
    const t = targets[i];
    const request = buildCookingAiRequest(t.combinedInput, t.dateStr, MODEL_NAME, GEMINI_API_KEY);
    const resText = executeCookingApiWithRetry(request, t.dateStr, errorLogs);
    
    if (resText) {
      try {
        const json = JSON.parse(resText);
        if (!json.error && json.candidates && json.candidates[0].content.parts[0].text) {
          const results = parseCookingMarkdownTable(json.candidates[0].content.parts[0].text);
          if (results && results.length > 0) {
            allResults.push(...results);
            // E列（5列目）に「済」を入力
            diarySheet.getRange(t.rowNum, 5).setValue('済');
            totalAdded += results.length;
          } else {
            errorLogs.push(`[${t.dateStr}] AIの回答形式が不正（表が見つからない）`);
          }
        } else if (json.error) {
          errorLogs.push(`[${t.dateStr}] AIエラー: ${json.error.message}`);
        }
      } catch (e) {
        errorLogs.push(`[${t.dateStr}] JSONパース失敗: ${e.toString()}`);
      }
    }
    Utilities.sleep(600); 
  }

  if (allResults.length > 0) {
    saveToCookingExperimentSheet(cookExpSheet, allResults);
  }

  let finalMessage = `${totalAdded}件の料理ログを転記しました。\n`;
  if (errorLogs.length > 0) {
    finalMessage += `\n【失敗 (${errorLogs.length}件)】\n` + errorLogs.join('\n');
  }
  SpreadsheetApp.getUi().alert(finalMessage);
}


/**
 * 料理用AIリクエスト生成
 */
function buildCookingAiRequest(text, dateStr, MODEL_NAME, GEMINI_API_KEY) {
  const url = `https://generativelanguage.googleapis.com/v1/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `あなたは「料理実験ログ整理アシスタント」です。
    ユーザーが雑に語る料理の実験内容を、事実ベースで整理し、表形式にまとめてください。

    ■ 絶対ルール
    - ユーザーが言っていない事実・手順・分量を一切補完しない
    - 曖昧な表現（適量・少し等）はそのまま残す
    - **重要：Markdownテーブル内での改行制御**
      - セル内で改行が必要な場合は、必ず「  <br>」という文字列を使用してください。実際の改行コードは絶対に使用禁止。

    ■ 出力フォーマット
    Markdown表：
    日付 | ジャンル | 主食材 | 料理名 | レシピ | 結果 | 点数 | ステータス | 次のレシピ | 改善意図

    ■ 各項目のルール
    ▼ レシピ：【材料】箇条書き、【手順】番号付き。間は「  <br><br>」で区切る。
    ▼ 結果：自然な独り言形式。要約禁止。主観や納得感を残す。
    ▼ 点数：発言があれば記載（10点満点）。なければ空欄。
    ▼ ステータス：失敗なら「🏆学習」、成功・実験済なら「☑️実験済」。
    ▼ 次のレシピ：そのまま再現できる具体的な材料と手順。改善は原則1要素のみ。
    ▼ 改善意図：なぜその改善をしたか、科学的・料理的根拠を含めて説明。不健康な改善は禁止。

    【日付指定】日付列には「${dateStr}」と入れてください。
    【入力】\n${text}`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
}

/**
 * リトライ付きAPI実行
 */
function executeCookingApiWithRetry(request, dateStr, errorLogs, maxRetries = 2) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      if (attempt > 0) Utilities.sleep(Math.pow(2, attempt) * 1000);
      const res = UrlFetchApp.fetch(request.url, {
        method: request.method,
        contentType: request.contentType,
        payload: request.payload,
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 200) return res.getContentText();
      if (code !== 429 && code < 500) {
        errorLogs.push(`[${dateStr}] HTTP ${code}`);
        return null;
      }
    } catch (e) {
      if (attempt === maxRetries) errorLogs.push(`[${dateStr}] 通信失敗`);
    }
  }
  return null;
}

/**
 * Markdown解析
 */
function parseCookingMarkdownTable(md) {
  return md.split('\n')
    .filter(line => line.includes('|') && !line.includes('---'))
    .map(line => {
      const cols = line.split('|').map(c => c.trim()).filter((c, i, arr) => i !== 0 && i !== arr.length - 1);
      return cols.map(cell => cell.replace(/\s*<br\s*\/?>\s*/gi, '\n'));
    })
    .filter(cols => cols.length >= 10 && cols[0] !== '日付');
}

/**
 * 書き込み
 */
function saveToCookingExperimentSheet(sheet, data) {
  const colA = sheet.getRange("A:A").getValues();
  let row = 5;
  for (let i = 4; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === "") {
      row = i + 1;
      break;
    }
  }
  const range = sheet.getRange(row, 1, data.length, data[0].length);
  range.setValues(data);
  range.setWrap(true);
  range.setVerticalAlignment("top");
}