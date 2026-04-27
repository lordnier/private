// ===== 設定値（固定項目）=====
const TAVILY_ENDPOINT = 'https://api.tavily.com/search';
const TAVILY_INCLUDE_DOMAIN = 'yuchrszk.blogspot.com'; // パレオな男


/**
 * エントリーポイント
 */


function runPaleoResearch() {
  // ここですべての設定を一括ロード（実験.gs の関数）
  const config = loadConfig_();

　// キーワードだけを UI から取得する
  const ui = SpreadsheetApp.getUi(); // [web:15]
  const response = ui.prompt(
    '検索キーワード入力',
    '「パレオな男」からリサーチしたいテーマ（例：筋トレ、睡眠など）を入力してください。',
    ui.ButtonSet.OK_CANCEL
  ); // [web:15][web:18]

  if (response.getSelectedButton() !== ui.Button.OK) {
    // キャンセルや閉じるが押されたら、何もせず終了
    return;
  }

  const keyword = response.getResponseText().trim(); // [web:19]
  if (!keyword) {
    ui.alert('検索キーワードが入力されていません。'); // ここはお好みで
    return;
  }

  const tavilyKey = config['TAVILY_API_KEY'];
  const geminiKey = config['GEMINI_API_KEY'];
  const modelName = config['MODEL_NAME'];

  const paleoArticles = searchPaleoArticlesWithTavily_(keyword, tavilyKey);

  if (!paleoArticles || paleoArticles.length === 0) {
    Logger.log('該当記事が見つかりませんでした。');
    writeResultTable_([['No data', '', '', '']]);
    return;
  }

  const prompt = buildGeminiPrompt_(keyword, paleoArticles);
  const tableRows = callGeminiAndParseTable_(prompt, geminiKey, modelName);

  writeResultTable_(tableRows);
}


/**
 * Tavily Search API を叩いて、「パレオな男」ブログだけを対象に記事一覧を取得
 */
function searchPaleoArticlesWithTavily_(keyword, apiKey) {
  const payload = {
    query: `site:${TAVILY_INCLUDE_DOMAIN} "${keyword}"`,
    topic: 'general',
    search_depth: 'basic',
    max_results: 10,
    include_raw_content: true,
    include_answer: false,
    include_domains: [TAVILY_INCLUDE_DOMAIN]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(TAVILY_ENDPOINT, options);
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('Tavily API error: ' + code + ' ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());

  const articles = (data.results || []).filter(r => {
    const url = r.url || '';
    if (url.indexOf('/search') !== -1) return false;
    if (url.indexOf('/label/') !== -1) return false;
    if (url.indexOf('_archive.html') !== -1) return false;
    return true;
  });

  return articles.map(a => ({
    title: a.title,
    url: a.url,
    content: a.content || ''
  }));
}

/**
 * Gemini に渡すプロンプトを構築
 */
function buildGeminiPrompt_(keyword, articles) {
  const articlesText = articles.map((a, i) => {
    return [
      `### Article ${i + 1}`,
      `タイトル: ${a.title}`,
      `URL: ${a.url}`,
      `本文抜粋:`,
      a.content.substring(0, 4000)
    ].join('\n');
  }).join('\n\n');

  const systemPrompt = `
あなたは「科学的情報整理」と「英語学習支援」を同時に行う専門家です。
以下の指示に従い、「パレオな男」のブログ記事から特定テーマに関する“科学的メリット”を5~8個抽出し、それをもとに英訳トレーニング用の和文を生成してください。

【目的】
・習慣化のモチベーションを高める（科学的メリットの理解）
・英語学習（和文英訳）を同時に行う

【対象テーマ】
「${keyword}」

【情報源の制約】
・必ず「パレオな男」ブログ内の情報のみを使用すること
・各メリットごとに、対応する記事本文ページのURL（個別記事ページ）を必ず明記すること
・検索結果一覧ページやタグ一覧ページのURLは使用しないこと
・リンクは、クリックすると該当内容が直接確認できるページであること
・科学的根拠（研究・論文・レビューなど）に基づく記述を優先すること


【出力形式】
以下の形式の表で出力すること：
No/英訳課題（日本語）/メリット/ソース

【各列のルール】
■ 英訳課題（日本語）
・各メリットごとに1つ作成

【目的】
・暗記しやすく、自然で滑らかな日本語にすること
・英訳しやすいシンプルな構造にすること

【構造ルール】
・1〜2文で構成（無理に2文にしない）
・1文で自然に収まる場合は1文を優先する
・意味の分断を避け、自然な流れを最優先する

【長さルール】
・全体で35〜55文字程度
・この範囲を外れた場合は自動で修正すること

【自然さルール（重要）】
・日本語として違和感がないことを最優先する
・不自然な言い換えや語尾調整を禁止
・意味を削って文字数を合わせることを禁止

【内容ルール】
・元のメリットの意味を保つ
・原因→結果の関係が伝わるようにする

■ メリット
以下の構成で記述すること
・【結論】（一言で端的に）<br>
・【理由】（科学的な仕組みや背景）<br>
・【具体例】（直感的に理解できる比喩や状況説明）<br>
※ソース情報（URLや出典）はメリット列には一切含めないこと

■ ソース
・該当記事の「タイトル」を表示し、そのテキストにURLを埋め込んだリンク形式で記載すること
・URLをそのまま裸で記載しないこと


【補足ルール】
・専門用語はできるだけ噛み砕く
・結論は断定的に書く（〜できる、〜が向上する）
・具体例は理解促進のため自由に補足OK
`;

  return systemPrompt + '\n\n' + '【利用可能な記事】\n' + articlesText;
}

/**
 * Gemini API を叩き、Markdown テーブルをパースして 2次元配列に変換
 */
function callGeminiAndParseTable_(prompt, apiKey, modelName) {
  const model = modelName || 'gemini-1.5-pro';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }]
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('Gemini API error: ' + code + ' ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());
  const text = (((data.candidates || [])[0] || {}).content || {}).parts
    ?.map(p => p.text || '')
    .join('') || '';

  return parseMarkdownTable_(text);
}

/**
 * Markdown テーブル文字列を 2次元配列に変換するユーティリティ
 */
function parseMarkdownTable_(mdTable) {
  const lines = mdTable
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  const tableLines = lines.filter(l => l.indexOf('|') !== -1);

  if (tableLines.length === 0) {
    throw new Error('Gemini 出力からテーブルを検出できませんでした。');
  }

  const result = [];
  for (let i = 0; i < tableLines.length; i++) {
    const line = tableLines[i];

    if (/^(\|\s*:?-{3,}:?\s*)+\|?$/.test(line)) {
      continue;
    }

    const cols = line
      .split('|')
      .map(c => c.trim())
      .filter((_, idx, arr) => !(idx === 0 && arr.length > 1) && !(idx === arr.length - 1 && arr.length > 1));

    if (cols.length > 0) {
      result.push(cols);
    }
  }

  return result;
}

/**
 * Result シートに 2次元配列を出力
 */
function writeResultTable_(rows) {
　// 1) 列数を4列に揃える（No / 英訳課題 / メリット / ソース 前提）
  const fixedRows = rows.map(r => {
    if (r.length > 4) {
      return r.slice(0, 4);         // 5列以上なら先頭4列だけ
    }
    if (r.length < 4) {
      const copy = r.slice();
      while (copy.length < 4) copy.push('');
      return copy;                  // 4列未満なら空文字で埋める
    }
    return r;                       // ちょうど4列ならそのまま
  });

  // 2) <br> をセル内改行に変換
  const cleanedRows = fixedRows.map(row =>
    row.map(cell =>
      typeof cell === 'string'
        ? cell.replace(/<br\s*\/?>/gi, '\n')  // HTML改行 → 改行コード[web:31][web:39]
        : cell
    )
  );
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Result');
  if (!sheet) {
    sheet = ss.insertSheet('Result');
  }
  sheet.clearContents();

  const numRows = cleanedRows.length;
  const numCols = cleanedRows[0].length;
  sheet.getRange(1, 1, numRows, numCols).setValues(cleanedRows);
}

