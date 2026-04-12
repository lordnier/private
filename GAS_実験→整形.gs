// 【設定】APIキーを入れてください
const GEMINI_API_KEY = 'AIzaSyBV_NWM1r4Aiq1RaIrCttdfk0Dl8ihk8kA';
const DIARY_SHEET = '1.日記';
const EXP_SHEET = '2.実験';

// モデル名は指定通り維持
const MODEL_NAME = 'gemini-2.5-flash';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪実験管理')
    .addItem('日記を解析して転記', 'processDiaryToExperiment')
    .addToUi();
}

/**
 * メイン処理（並列化版）
 */
function processDiaryToExperiment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diarySheet = ss.getSheetByName(DIARY_SHEET);
  const expSheet = ss.getSheetByName(EXP_SHEET);

  const diaryData = diarySheet.getRange(2, 1, diarySheet.getLastRow() - 1, 5).getValues();

  // 1. 処理対象のデータをリストアップ
  const targets = [];
  for (let i = 0; i < diaryData.length; i++) {
    const rowNum = i + 2;
    const date = diaryData[i][0];
    const expMemo = diaryData[i][1]; // B列：実験
    const cookMemo = diaryData[i][3]; // D列：料理
    const status = String(diaryData[i][4]).trim();

    if ((expMemo || cookMemo) && status === "") {
      const dateStr = date instanceof Date ? Utilities.formatDate(date, "JST", "yyyy-MM-dd") : String(date);
      const combinedInput = `【日記・実験】\n${expMemo}\n\n【料理ログ】\n${cookMemo}`;
      targets.push({ rowNum, dateStr, combinedInput });
    }
  }

  if (targets.length === 0) {
    SpreadsheetApp.getUi().alert('対象の未処理日記が見つかりませんでした。');
    return;
  }

  // 2. AIへのリクエストを一括作成
  SpreadsheetApp.getActive().toast(`${targets.length}件を並列解析中...`, '🧪');
  const requests = targets.map(t => createAiRequest(t.combinedInput, t.dateStr));

  // 3. 一括送信
  const responses = UrlFetchApp.fetchAll(requests);

  // 4. 結果をまとめて処理
  let totalAdded = 0;
  const allResults = [];

  targets.forEach((t, index) => {
    const resText = responses[index].getContentText();
    const json = JSON.parse(resText);

    if (!json.error && json.candidates && json.candidates[0].content.parts[0].text) {
      const results = parseMarkdown(json.candidates[0].content.parts[0].text);
      if (results && results.length > 0) {
        allResults.push(...results);
        diarySheet.getRange(t.rowNum, 5).setValue('済');
        totalAdded += results.length;
      }
    } else {
      console.error(`Error for ${t.dateStr}:`, resText);
    }
  });

  // 5. 実験シートへ一括書き込み
  if (allResults.length > 0) {
    writeToExperimentSheet(expSheet, allResults);
  }

  SpreadsheetApp.getUi().alert(totalAdded > 0 ? `${totalAdded}件の実験を転記しました。` : '解析に失敗したか、有効なデータがありませんでした。');
}

/**
 * AIリクエストオブジェクトの生成
 */
function createAiRequest(text, dateStr) {
  const url = `https://generativelanguage.googleapis.com/v1/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  const prompt = `あなたは実験結果の管理アシスタント」です。
ユーザーが雑に語る「試したこと」を、事実ベースで構造化・評価し、再利用可能な形に整理してください。

■ 絶対ルール（最優先）
ユーザー未言及の情報は一切追加しない（推測・補完・具体化禁止）

■ 出力フォーマット（固定）
Markdown表：

日付 | スキル | ミニスキル | If (トリガー) | Then (アクション) | 結果 | ステータス

■ 構造化ルール
▼ If / Then（最重要）
情報は削らず整理する
※セルの内部で改行が必要な箇所（アイコンの区切りなど）には、必ず「<br>」という文字列を使用してください。
※実際の改行（リターンキー）はMarkdownテーブルの構造を壊すため、セル内では絶対に使用しないでください。
If： 👀 状況：〇〇<br>🎯 狙い：〇〇
Then： ⚡ 行動：〇〇<br>💬 具体例：〇〇

■ 結果（最重要）
・ユーザの発言内容をもとに、発言者の思考の流れが感じられる自然な独り言形式で整理すること
・要約は禁止（要点のみの箇条書き化も禁止）
▼表現ルール
・一文で完結させず、思考の流れがつながる文章にする
・主観・迷い・納得感の表現を適度に残す
・ただし同じ内容の繰り返しや言い直しは削除する
▼NG
・結論だけの短文化（議事録的表現）
・説明過多な整形（不自然にきれいな文章）
▼目標状態
・「本人に少しだけ言語化うまくなった状態」を再現すること

■ ステータス（上から優先して判定する）
①「失敗」「うまくいかなかった」など明確にネガティブ評価している場合：
　→「💡Not to do」と入力
② 実験したことが明確な場合：
　→「☑️実験済」と入力
③ 実験していないことが明確な場合：
　→「🔒未実験」と入力

【日付指定】日付列には「${dateStr}」と入れてください。
【入力】${text}`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
}

/**
 * 表データの解析（<br>をスプレッドシートの改行に置換）
 */
function parseMarkdown(md) {
  return md.split('\n')
    .filter(line => line.includes('|') && !line.includes('---'))
    .map(line => {
      // 各カラムを分割してトリミング
      const cols = line.split('|').map(c => c.trim()).filter((c, i, arr) => i !== 0 && i !== arr.length - 1);
      // セル内の <br> をスプレッドシートの改行コード \n に置換
      return cols.map(cell => cell.replace(/<br>/g, '\n'));
    })
    .filter(cols => cols.length >= 7 && cols[0] !== '日付');
}

/**
 * 実験シートの空行を探して書き込み（折り返し設定を有効化）
 */
function writeToExperimentSheet(sheet, data) {
  const colA = sheet.getRange("A:A").getValues();
  let row = 5;
  for (let i = 4; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === "") {
      row = i + 1;
      break;
    }
  }

  const range = sheet.getRange(row, 1, data.length, data[0].length);
  
  // データの書き込み
  range.setValues(data);

  // 見た目を美しく整える設定
  range.setWrap(true);               // テキストを折り返して改行を表示
  range.setVerticalAlignment("top"); // 長文でも見やすいよう上揃えに設定
}