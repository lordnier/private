// 【設定】APIキーを入れてください
const GEMINI_API_KEY = 'AIzaSyCNqONvfL5qm597r4vSj-cysBRjT_dpg6I';
const DIARY_SHEET = '1.日記';
const EXP_SHEET = '2.実験';

// 2026年時点での最新・安定版モデルを指定
const MODEL_NAME = 'gemini-2.5-flash'; 

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪実験管理')
    .addItem('日記を解析して転記', 'processDiaryToExperiment')
    .addToUi();
}

/**
 * メイン処理
 */
async function processDiaryToExperiment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diarySheet = ss.getSheetByName(DIARY_SHEET);
  const expSheet = ss.getSheetByName(EXP_SHEET);
  
  const diaryData = diarySheet.getRange(2, 1, diarySheet.getLastRow() - 1, 5).getValues();
  let totalAdded = 0;

  for (let i = 0; i < diaryData.length; i++) {
    const rowNum = i + 2;
    const date = diaryData[i][0];
    const expMemo = diaryData[i][1]; // B列：実験
    const cookMemo = diaryData[i][3]; // D列：料理
    const status = String(diaryData[i][4]).trim();

    // 未処理（E列が空）かつ内容がある場合
    if ((expMemo || cookMemo) && status === "") {
      const dateStr = date instanceof Date ? Utilities.formatDate(date, "JST", "yyyy-MM-dd") : String(date);
      const combinedInput = `【日記・実験】\n${expMemo}\n\n【料理ログ】\n${cookMemo}`;

      SpreadsheetApp.getActive().toast(`${dateStr}を解析中...`, '🧪');

      // AIに投げて解析（文字数が多い場合は自動分割）
      const results = await fetchAiAnalysis(combinedInput, dateStr);

      if (results && results.length > 0) {
        writeToExperimentSheet(expSheet, results);
        diarySheet.getRange(rowNum, 5).setValue('済');
        totalAdded += results.length;
      }
    }
  }

  SpreadsheetApp.getUi().alert(totalAdded > 0 ? `${totalAdded}件の実験を転記しました。` : '対象の未処理日記が見つかりませんでした。');
}

/**
 * AI呼び出し
 */
async function fetchAiAnalysis(text, dateStr) {
  const url = `https://generativelanguage.googleapis.com/v1/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `あなたは実験結果の管理アシスタント」です。 ユーザーが雑に語る「試したこと」を、事実ベースで構造化・評価し、再利用可能な形に整理してください。 
■ 絶対ルール（最優先） 
ユーザー未言及の情報は一切追加しない（推測・補完・具体化禁止） 
■ 出力フォーマット（固定） 
Markdown表： 日付 / スキル / ミニスキル / If (トリガー) / Then (アクション) / 結果 / ステータス 
■ 構造化ルール 
▼ If / Then（最重要） 
情報は削らず整理する 
※改行は実際の改行として出力し、「\\n」などの文字列による改行表現は使用しない 
※If / Thenは「その時点で実際に起きたこと・考えたこと」のみ記述する 
※「〜すればよかった」「次は〜する」などの改善・未来の内容は記載しない（結果に含める） 
If： 
👀 状況：〇〇（実際に起きた事実） 
🎯 狙い：〇〇（その時に考えていた意図） 
Then： 
⚡ 行動：〇〇（実際に取った行動のみ） 
💬 具体例：〇〇（実際の発言・振る舞いのみ） 
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
・「本人が少しだけ言語化うまくなった状態」を再現すること 
■ ステータス（上から優先して判定する） 
①「失敗」「うまくいかなかった」など明確にネガティブ評価している場合： 　→「💡Not to do」と入力（※実験済みであっても最優先） 
② 実験したことが明確な場合： 　→「☑️実験済」と入力 
③ 実験していないことが明確な場合： 　→「🔒未実験」と入力

【日付指定】日付列には「${dateStr}」と入れてください。
【入力】
${text}`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const resText = response.getContentText();
  const json = JSON.parse(resText);
  if (json.error) {
    console.error(resText);
    return null;
  }
  return parseMarkdown(json.candidates[0].content.parts[0].text);
}

/**
 * 表データの解析
 */
function parseMarkdown(md) {
  return md.split('\n')
    .filter(line => line.includes('|') && !line.includes('---'))
    .map(line => line.split('|').map(c => c.trim()).filter((c, i, arr) => i !== 0 && i !== arr.length - 1))
    .filter(cols => cols.length >= 7 && cols[0] !== '日付');
}

/**
 * 実験シートの空行を探して書き込み
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
  sheet.getRange(row, 1, data.length, data[0].length).setValues(data);
}
