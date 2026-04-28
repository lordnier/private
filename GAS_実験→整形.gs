
// どの .gs ファイルに置いてもOK（実験.gsでも料理.gsでも可）
function runAllDailyTasks() {
  // 1. 実験管理
  processDiaryToExperiment();        // 実験.gs 側のメイン関数

  // 2. 料理管理
  runCookingLogAnalysis();   // 料理.gs 側のメイン関数
}


/**
 * Configシートから共通設定を読み込む
 * A列=キー, B列=値
 */
function loadConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  if (!sheet) {
    throw new Error('Config シートが見つかりません。シート名「Config」で作成してください。');
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const config = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0]).trim();
    const value = values[i][1];
    if (!key) continue;
    config[key] = value;
  }

  const requiredKeys = [
    'GEMINI_API_KEY',
    'MODEL_NAME',
    'DIARY_SHEET',
    'EXP_SHEET',
    'COOK_EXP_SHEET',
    'CONTEXT_INFO'
  ];
  requiredKeys.forEach(k => {
    if (!(k in config) || config[k] === '') {
      throw new Error('Config シートに必須キー「' + k + '」の値が設定されていません。');
    }
  });

  return config;
}

// ========= 実験用：メイン処理 =========

/**
 * メニュー作成関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪実験記録')
    .addItem('実験＋料理をまとめて実行', 'runAllDailyTasks')
    .addToUi();

  ui.createMenu('Paleo Research')
  .addItem('パレオな男をリサーチ', 'runPaleoResearch')
  .addToUi();

  ui.createMenu('📝CSV出力')
    .addItem('選択範囲をCSVダウンロード', 'exportSelectedRangeAsCSV')
    .addToUi();


}


/**
 * メイン処理（直列処理＋エラー詳細表示版）
 */function processDiaryToExperiment() {
  const config = loadConfig_();
  const GEMINI_API_KEY = config.GEMINI_API_KEY;
  const DIARY_SHEET = config.DIARY_SHEET;
  const EXP_SHEET = config.EXP_SHEET;
  const CONTEXT_INFO = config.CONTEXT_INFO;
  const MODEL_NAME = config.MODEL_NAME;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diarySheet = ss.getSheetByName(DIARY_SHEET);
  const expSheet = ss.getSheetByName(EXP_SHEET);

  if (!diarySheet) {
    SpreadsheetApp.getUi().alert('日記シートが見つかりません: ' + DIARY_SHEET);
    return;
  }
  if (!expSheet) {
    SpreadsheetApp.getUi().alert('実験シートが見つかりません: ' + EXP_SHEET);
    return;
  }

  const contextInfo = CONTEXT_INFO;
  const diaryData = diarySheet.getRange(2, 1, diarySheet.getLastRow() - 1, 4).getValues();

  const targets = [];
  for (let i = 0; i < diaryData.length; i++) {
    const rowNum = i + 2;
    const date = diaryData[i][0];
    const expMemo = diaryData[i][1];
    const status = String(diaryData[i][2]).trim();

    if (expMemo && status === "") {
      const dateStr = date instanceof Date
        ? Utilities.formatDate(date, "JST", "yyyy-MM-dd")
        : String(date);
      const combinedInput = `【日記・実験】\n${expMemo}`;
      targets.push({ rowNum, dateStr, combinedInput });
    }
  }

  if (targets.length === 0) {
    SpreadsheetApp.getActive().toast('対象の未処理日記が見つかりませんでした。', '実験管理', 5);
    return;
  }

  let totalAdded = 0;
  const allResults = [];
  const errorLogs = [];
  
  SpreadsheetApp.getActive().toast(`${targets.length}件を順番に解析中...`, '🧪');

  for (let i = 0; i < targets.length; i++) {
    const t = targets[i];
    const request = createAiRequest(
      t.combinedInput,
      t.dateStr,
      contextInfo,
      MODEL_NAME,
      GEMINI_API_KEY
    );
    const resText = fetchSingleWithRetry(request, t.dateStr, errorLogs);
    
    if (resText) {
      try {
        const json = JSON.parse(resText);
        if (!json.error && json.candidates && json.candidates[0].content.parts[0].text) {
          const results = parseMarkdown(json.candidates[0].content.parts[0].text);
          if (results && results.length > 0) {
            allResults.push(...results);
            diarySheet.getRange(t.rowNum, 3).setValue('済');
            totalAdded += results.length;
          } else {
            errorLogs.push(`[${t.dateStr}] AIの回答形式が不正（表が見つからない）`);
          }
        } else if (json.error) {
          errorLogs.push(`[${t.dateStr}] AIエラー: ${json.error.message}`);
        } else {
          errorLogs.push(`[${t.dateStr}] 解析エラー: 回答が空です`);
        }
      } catch (e) {
        errorLogs.push(`[${t.dateStr}] JSONパース失敗: ${e.toString()}`);
      }
    }
    Utilities.sleep(500); 
  }

  if (allResults.length > 0) {
    writeToExperimentSheet(expSheet, allResults);
  }

  let finalMessage = `${totalAdded}件の実験を転記しました。\n`;
  if (errorLogs.length > 0) {
    finalMessage += `\n【失敗した処理 (${errorLogs.length}件)】\n` + errorLogs.join('\n');
  }
  SpreadsheetApp.getActive().toast(finalMessage, '実験管理', 10);  // 10秒表示
}


// ========= 実験用：AIリクエスト生成（引数だけ変更） =========
/**
 * 単一リクエストのリトライ機能付き実行（直列用）
 * （ここは元コードから一切変更なし）
 */
function fetchSingleWithRetry(request, dateStr, errorLogs, maxRetries = 2) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      if (attempt > 0) {
        const waitTime = Math.pow(2, attempt) * 1000;
        Utilities.sleep(waitTime);
      }
      
      const res = UrlFetchApp.fetch(request.url, {
        method: request.method,
        contentType: request.contentType,
        payload: request.payload,
        muteHttpExceptions: true
      });
      
      const code = res.getResponseCode();
      if (code === 200) {
        return res.getContentText();
      } else {
        const errorMsg = `HTTP ${code}: ${res.getContentText().substring(0, 100)}...`;
        if (code === 429 || code >= 500) {
          if (attempt === maxRetries) {
            errorLogs.push(`[${dateStr}] 最大リトライ超過 (${errorMsg})`);
          }
          continue; 
        } else {
          errorLogs.push(`[${dateStr}] 致命的エラー (${errorMsg})`);
          return null;
        }
      }
    } catch (e) {
      if (attempt === maxRetries) {
        errorLogs.push(`[${dateStr}] 通信失敗: ${e.toString()}`);
      }
    }
  }
  return null;
}


/**
 * AIリクエストオブジェクトの生成
 * → MODEL_NAME / API_KEY を引数でもらうようにしただけ
 */
function createAiRequest(text, dateStr, contextInfo, MODEL_NAME, GEMINI_API_KEY) {
  const url = `https://generativelanguage.googleapis.com/v1/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
  const prompt = `あなたは「実験結果の管理アシスタント」です。ユーザーが雑に語る「試したこと」を、事実ベースで構造化・評価し、再利用可能な形に整理してください。

    ■ 前提知識（ユーザーの背景知識・独自の用語定義）
    ${contextInfo}

    ■ 絶対ルール（最優先）
    ユーザー未言及の情報は一切追加しない（推測・補完・具体化禁止）

    ■ 出力フォーマット（固定）
    Markdown表：

    日付 | スキル | ミニスキル | If (トリガー) | Then (アクション) | 結果 | ステータス

    ■ 構造化ルール
    ▼ If / Then（最重要）
    情報は削らず整理する
    ※セルの内部で改行が必要な箇所（アイコンの区切りなど）には、必ず「  <br>」という文字列を使用してください。
    ※実際の改行（リターンキー）はMarkdownテーブルの構造を壊すため、セル内では絶対に使用しないでください。
    If：
    👀 状況：〇〇  <br>🎯 狙い：〇〇（実際に起きた事実）  
    Then：
    ⚡ 行動：〇〇  <br>💬 具体例：〇〇（その時に考えていた意図）  

    ■ 結果（最重要）
    ・ユーザの発言内容をもとに、発言者の思考の流れが感じられる自然な独り言形式で整理すること
    ・要約は禁止（要点のみの箇取り化も禁止）

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
    　→「🏆学習」と入力
    ② 実験したことが明確な場合：
    　→「☑️実験済」と入力
    ③ 実験していないことが明確な場合：
    　→「🔒未実験」と入力

    【日付指定】
    日付列には「${dateStr}」と入れてください。

    【入力】
    ${text}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  return {
    url: url,
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
}


/**
 * 表データの解析（変更なし）
 */
function parseMarkdown(md) {
  return md.split('\n')
    .filter(line => line.includes('|') && !line.includes('---'))
    .map(line => {
      const cols = line.split('|').map(c => c.trim()).filter((c, i, arr) => i !== 0 && i !== arr.length - 1);
      return cols.map(cell => cell.replace(/\s*<br\s*\/?>\s*/gi, '\n'));
    })
    .filter(cols => cols.length >= 7 && cols[0] !== '日付');
}


/**
 * 実験シートの空行を探して書き込み（変更なし）
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
  range.setValues(data);
  range.setWrap(true);
  range.setVerticalAlignment("top");
}