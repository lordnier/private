function callMiniMaxTTS_(apiKey, text) {
  const url = 'https://api.minimax.io/v1/t2a_v2';

  const payload = {
    model: 'speech-2.8-hd',
    text: text,
    stream: false,
    language_boost: 'auto',
    output_format: 'hex',
    voice_setting: {
      voice_id: 'English_Whispering_girl_v3',
      speed: 0.9,
      vol: 0.8,
      pitch: 0
    },
    audio_setting: {
      sample_rate: 32000,
      bitrate: 128000,
      format: 'mp3',
      channel: 1
    },
    voice_modify: {
      pitch: 0,
      intensity: -2,
      timbre: 0
      // sound_effects を削除
    }
  };

  const headers = {
    Authorization: 'Bearer ' + apiKey,
    'Content-Type': 'application/json'
  };

  const options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const bodyText = res.getContentText();

  console.log('MiniMax response code:', code);
  console.log('MiniMax response body:', bodyText);

  if (code !== 200) {
    throw new Error('HTTP ' + code + ' : ' + bodyText);
  }

  const json = JSON.parse(bodyText);

  if (!json.base_resp || json.base_resp.status_code !== 0) {
    throw new Error(
      'MiniMax base_resp error: ' +
      (json.base_resp ? json.base_resp.status_msg + ' (code ' + json.base_resp.status_code + ')' : bodyText)
    );
  }

  if (!json.data || !json.data.audio) {
    throw new Error('MiniMaxレスポンスにaudioが含まれていません: ' + bodyText);
  }

  const audioHex = json.data.audio;
  if (!audioHex) {
    throw new Error('audioフィールドが空です: ' + bodyText);
  }

  return hexToBytes_(audioHex);
}

/**
 * 選択範囲の各セルテキストをMiniMaxでMP3化
 */
function generateTTSForSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();

  // Configシートから設定読み込み（既存のloadConfig_を使い回し）
  const config = loadConfig_();
  const apiKey = config['MINIMAX_API_KEY'];
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Config シートに MINIMAX_API_KEY が設定されていません。');
    return;
  }

  const folderName = 'MiniMax_TTS_Output';
  const folder = getOrCreateFolderByName_(folderName);

  const startRow = range.getRow();
  const startCol = range.getColumn();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const text = values[r][c];
      if (!text || typeof text !== 'string') {
        continue;
      }

      const row = startRow + r;
      const col = startCol + c;
      const a1Notation = sheet.getRange(row, col).getA1Notation();

      try {
        const fileName = `MiniMax_TTS_${sheet.getName()}_${a1Notation}.mp3`;
        const audioBytes = callMiniMaxTTS_(apiKey, text);
        folder.createFile(
          Utilities.newBlob(audioBytes, 'audio/mpeg', fileName)
        );
        console.log(`✅ ${a1Notation} を音声化し、ファイル「${fileName}」を作成しました`);
      } catch (e) {
        console.error(`❌ ${a1Notation} の音声生成に失敗: ${e}`);
      }

      Utilities.sleep(500);
    }
  }

  SpreadsheetApp.getUi().alert(
    'MiniMaxでのMP3生成が完了しました。Googleドライブの「' + folderName + '」フォルダを確認してください。'
  );
}

/**
 * hex文字列をバイト列に変換
 */
function hexToBytes_(hex) {
  const cleanHex = hex.replace(/[^0-9a-fA-F]/g, '');
  const bytes = [];
  for (let i = 0; i < cleanHex.length; i += 2) {
    bytes.push(parseInt(cleanHex.substr(i, 2), 16));
  }
  return Utilities.newBlob(bytes).getBytes();
}

/**
 * 指定名のフォルダを取得、なければ作成
 */
function getOrCreateFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}

function testMiniMaxOnce() {
  const config = loadConfig_();
  const apiKey = config['MINIMAX_API_KEY'];
  const text = 'This is a test for MiniMax whispering voice.';

  const bytes = callMiniMaxTTS_(apiKey, text);
  Logger.log('bytes length = ' + bytes.length);
}