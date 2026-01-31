/**
 * スプレッドシート操作モジュール
 *
 * 投稿、回答、ユーザー、トークルームの各シートへのアクセスと
 * データの記録・取得を行う関数群を提供します。
 */

// シート名の定数定義
var SHEET_NAMES = {
  POSTS: '投稿',
  ANSWERS: '回答',
  USERS: 'ユーザー',
  ROOMS: 'トークルーム',
  DEBUG: 'デバッグ',
  TRANSLATION_LOG: '翻訳ログ'
};

// 翻訳用定数
var GEMINI_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta/models/';
var GEMINI_MODELS = [
  'gemini-2.5-flash-lite',
  'gemini-2.5-flash',
  'gemini-3-flash-preview',
  'gemma-3-27b-it'
];
var MAX_HISTORY_COUNT = 2;

/**
 * デバッグシートにログを出力する関数
 *
 * @param {string} message エラーメッセージ
 * @param {string} stack スタックトレース
 */
function debugToSheet(message, stack) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.DEBUG);

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.DEBUG);
      sheet.appendRow(['timestamp', 'message', 'stack']);
    }

    sheet.appendRow([new Date(), message, stack || '']);
  } catch (e) {
    // console.log は使用しない
  }
}

/**
 * スプレッドシートを取得する関数
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} アクティブなスプレッドシート
 */
function getSpreadsheet() {
  var id = getScriptProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

/**
 * 投稿を記録する関数
 *
 * @param {string} postId LINEのメッセージID
 * @param {Date} timestamp 投稿日時
 * @param {string} userId ユーザーID
 * @param {string} roomId トークルームID（個人チャットの場合はnullまたはuserIdと同じ）
 * @param {string} messageText メッセージ内容
 * @param {boolean} hasPoll アンケートが含まれているかどうか
 * @param {string} translatedText 翻訳されたメッセージ内容（任意）
 */
function recordPost(postId, timestamp, userId, roomId, messageText, hasPoll, translatedText) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.POSTS);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.POSTS);
    sheet.appendRow(['post_id', 'timestamp', 'user_id', 'room_id', 'message_text', 'has_poll', 'translated_text']);
  } else {
    // ヘッダー行を確認し、translated_text列がない場合は追加（既存シートへの対応）
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.indexOf('translated_text') === -1) {
      sheet.getRange(1, header.length + 1).setValue('translated_text');
    }
  }

  sheet.appendRow([postId, timestamp, userId, roomId, messageText, hasPoll, translatedText || '']);
}

/**
 * 回答を記録する関数
 *
 * @param {string} pollPostId アンケートの元投稿ID
 * @param {Date} timestamp 回答日時
 * @param {string} userId 回答したユーザーID
 * @param {string} answerValue 回答内容 (OK/NG)
 */
function recordAnswer(pollPostId, timestamp, userId, answerValue) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ANSWERS);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ANSWERS);
    sheet.appendRow(['answer_id', 'timestamp', 'poll_post_id', 'user_id', 'answer_value']);
  }

  var data = sheet.getDataRange().getValues();
  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    // poll_post_id は 3列目 (index 2)
    // user_id は 4列目 (index 3)
    if (data[i][2] === pollPostId && data[i][3] === userId) {
      // 既存の回答を更新
      // timestamp (index 1 -> 列2)
      // answer_value (index 4 -> 列5)
      sheet.getRange(i + 1, 2).setValue(timestamp);
      sheet.getRange(i + 1, 5).setValue(answerValue);
      return;
    }
  }

  // 存在しない場合は新規追加
  var answerId = Utilities.getUuid();
  sheet.appendRow([answerId, timestamp, pollPostId, userId, answerValue]);
}

/**
 * ユーザーが存在しない場合に新規登録する関数
 *
 * @param {string} userId ユーザーID
 */
function ensureUser(userId) {
  if (!userId) return;

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USERS);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    sheet.appendRow(['user_id', 'display_name']);
  }

  var data = sheet.getDataRange().getValues();
  // ヘッダー行を除くデータから検索
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return; // 既に存在する
    }
  }

  // 存在しない場合のみ追加。名前は空欄（管理者が手動入力）
  sheet.appendRow([userId, '']);
}

/**
 * トークルームが存在しない場合に新規登録する関数
 *
 * @param {string} roomId トークルームID
 */
function ensureRoom(roomId) {
  if (!roomId) return;

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ROOMS);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ROOMS);
    sheet.appendRow(['room_id', 'room_name']);
  }

  var data = sheet.getDataRange().getValues();
  // ヘッダー行を除くデータから検索
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === roomId) {
      return; // 既に存在する
    }
  }

  // 存在しない場合のみ追加。ルーム名は空欄（管理者が手動入力）
  sheet.appendRow([roomId, '']);
}

/**
 * ユーザー名を更新する関数
 *
 * @param {string} userId ユーザーID
 * @param {string} newName 新しいユーザー名
 */
function updateUserName(userId, newName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      // 2列目 (index 1) を更新
      sheet.getRange(i + 1, 2).setValue(newName);
      return;
    }
  }
  // ユーザーが存在しない場合は追加
  sheet.appendRow([userId, newName]);
}

/**
 * 指定された投稿IDに対する回答を集計する関数
 *
 * @param {string} postId アンケートの投稿ID
 * @returns {Object} {ok: number, ng: number} 集計結果
 */
function getPollResults(postId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ANSWERS);
  if (!sheet) return { ok: 0, ng: 0 };

  var data = sheet.getDataRange().getValues();
  var okCount = 0;
  var ngCount = 0;

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    // poll_post_id は 3列目 (index 2)
    // answer_value は 5列目 (index 4)
    if (data[i][2] === postId) {
      var value = data[i][4];
      if (value === 'OK') okCount++;
      if (value === 'NG') ngCount++;
    }
  }

  return { ok: okCount, ng: ngCount };
}

/**
 * アンケートの詳細結果を取得する関数
 *
 * @param {string} postId アンケートの投稿ID
 * @returns {Array} 回答詳細の配列 [{timestamp, userName, answerValue}, ...]
 */
function getPollResultDetails(postId) {
  var ss = getSpreadsheet();

  // ユーザー情報を取得してマッピングを作成
  var userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  var userMap = {};
  if (userSheet) {
    var userData = userSheet.getDataRange().getValues();
    for (var i = 1; i < userData.length; i++) {
      userMap[userData[i][0]] = userData[i][1];
    }
  }

  // 回答データを取得
  var answerSheet = ss.getSheetByName(SHEET_NAMES.ANSWERS);
  if (!answerSheet) return [];

  var data = answerSheet.getDataRange().getValues();
  var results = [];

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    // poll_post_id は 3列目 (index 2)
    if (data[i][2] === postId) {
      var timestamp = new Date(data[i][1]);
      var userId = data[i][3];
      var answerValue = data[i][4];
      var userName = userMap[userId] || '未登録';

      results.push({
        timestamp: timestamp,
        userName: userName,
        answerValue: answerValue
      });
    }
  }

  // 日時の降順でソート（新しい順）
  results.sort(function(a, b) {
    return b.timestamp - a.timestamp;
  });

  return results;
}

/**
 * 指定された投稿IDのメッセージ内容と翻訳内容を取得する関数
 *
 * @param {string} postId 投稿ID
 * @returns {Object} { text: string, translatedText: string }
 */
function getPollContent(postId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.POSTS);
  if (!sheet) return { text: "投稿が見つかりません", translatedText: "" };

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var translatedTextIndex = header.indexOf('translated_text');

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    // post_id は 1列目 (index 0)
    // message_text は 5列目 (index 4)
    if (data[i][0] === postId) {
      var text = data[i][4];
      var translatedText = "";
      if (translatedTextIndex !== -1) {
        translatedText = data[i][translatedTextIndex];
      }
      return { text: text, translatedText: translatedText };
    }
  }
  return { text: "投稿が見つかりません", translatedText: "" };
}

/**
 * LINE Messaging API操作モジュール
 *
 * メッセージの返信、Flex Messageの生成など、LINE関連の機能を提供します。
 */

/**
 * プロパティサービスから設定値を取得
 */
function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * メッセージを返信する関数
 *
 * @param {string} replyToken 返信用トークン
 * @param {Array} messages 送信するメッセージオブジェクトの配列
 */
function replyMessages(replyToken, messages) {
  var token = getScriptProperty('CHANNEL_ACCESS_TOKEN');
  var url = 'https://api.line.me/v2/bot/message/reply';
  var payload = {
    'replyToken': replyToken,
    'messages': messages
  };

  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + token,
    },
    'method': 'post',
    'payload': JSON.stringify(payload)
  });
}

/**
 * Loadingアニメーションを表示する関数
 *
 * @param {string} userId ユーザーID
 * @param {number} seconds 表示秒数 (デフォルト2秒)
 */
function sendLoadingAnimation(userId, seconds) {
  var token = getScriptProperty('CHANNEL_ACCESS_TOKEN');
  var url = 'https://api.line.me/v2/bot/chat/loading/start';
  var payload = {
    'chatId': userId,
    'loadingSeconds': seconds || 5
  };

  try {
    UrlFetchApp.fetch(url, {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': 'Bearer ' + token,
      },
      'method': 'post',
      'payload': JSON.stringify(payload)
    });
  } catch (e) {
    debugToSheet('sendLoadingAnimation failed: ' + e.message, e.stack);
  }
}

/**
 * アンケート用のFlex Messageを作成する関数
 *
 * @param {string} originalPostId アンケート対象の投稿ID
 * @returns {Object} Flex Messageオブジェクト
 */
function createPollFlexMessage(originalPostId) {
  var webAppUrl = getScriptProperty('WEB_APP_URL');
  var resultsUrl = webAppUrl + '?postId=' + originalPostId;

  return {
    "type": "flex",
    "altText": "アンケート",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "Select one. Can be changed.",
            "weight": "bold",
            "size": "sm"
          }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "box",
            "layout": "horizontal",
            "spacing": "sm",
            "contents": [
              {
                "type": "button",
                "style": "primary",
                "height": "sm",
                "action": {
                  "type": "postback",
                  "label": "OK",
                  "data": "action=answer&value=OK&postId=" + originalPostId
                }
              },
              {
                "type": "button",
                "style": "secondary",
                "height": "sm",
                "action": {
                  "type": "postback",
                  "label": "NG",
                  "data": "action=answer&value=NG&postId=" + originalPostId
                }
              },
              {
                "type": "button",
                "style": "secondary",
                "height": "sm",
                "action": {
                  "type": "postback",
                  "label": "N/A",
                  "data": "action=answer&value=N/A&postId=" + originalPostId
                }
              }
            ]
          },
          {
            "type": "separator",
            "margin": "sm"
          },
          {
            "type": "button",
            "style": "link",
            "height": "sm",
            "action": {
              "type": "uri",
              "label": "See results",
              "uri": resultsUrl
            }
          }
        ],
        "flex": 0
      }
    }
  };
}

/**
 * LINE Bot メインエントリーポイント
 *
 * Webhookからのリクエストを受け取り、適切な処理に振り分けます。
 */

/**
 * 重複イベント（リトライ）かどうかを判定する関数
 *
 * @param {string} eventId WebhookイベントID
 * @returns {boolean} 処理済みの場合はtrue、未処理の場合はfalse
 */
function isProcessed(eventId) {
  var cache = CacheService.getScriptCache();
  // キャッシュに存在する場合は処理済みとみなす
  if (cache.get(eventId)) {
    return true;
  }
  // 処理済みとしてマーク（10分間キャッシュ）
  cache.put(eventId, 'processed', 600);
  return false;
}

/**
 * WebhookへのPOSTリクエストを処理する関数
 *
 * @param {Object} e イベントオブジェクト
 */
function doPost(e) {
  try {
    // LINEプラットフォームからの検証用リクエストの場合
    if (!e || !e.postData) {
      return ContentService.createTextOutput("OK");
    }

    var json = JSON.parse(e.postData.contents);
    var events = json.events;

    events.forEach(function(event) {
      // リトライガード: 処理済みのイベントIDはスキップ
      if (event.webhookEventId && isProcessed(event.webhookEventId)) {
        return;
      }

      if (event.type === 'message' && event.message.type === 'text') {
        handleMessageEvent(event);
      } else if (event.type === 'postback') {
        handlePostbackEvent(event);
      }
    });

    return ContentService.createTextOutput("OK");
  } catch (error) {
    debugToSheet(error.message, error.stack);
    // LINEプラットフォームにエラーを返さないようにOKを返す
    return ContentService.createTextOutput("OK");
  }
}

/**
 * メッセージイベントを処理する関数
 *
 * @param {Object} event LINEイベントオブジェクト
 */
function handleMessageEvent(event) {
  var messageId = event.message.id;
  var timestamp = new Date(event.timestamp);
  var userId = event.source.userId;

  // 翻訳処理開始前にLoadingアニメーションを表示 (60秒)
  sendLoadingAnimation(userId, 60);

  // グループまたはルームIDを取得。個人チャットの場合は空文字
  var roomId = event.source.roomId || event.source.groupId || "";
  var text = event.message.text;

  // ユーザーの確認・登録
  ensureUser(userId);

  // ルームIDがある場合のみ、ルームの確認・登録
  if (roomId) {
    ensureRoom(roomId);
  }

  // アンケートキーワードの判定
  var checkRegex = /\[check\]/i;
  var hasPoll = checkRegex.test(text);
  // ユーザー名更新コマンドの判定
  var nameMatch = text.match(/私(?:の名前|)は"(.+?)"/);

  // ユーザー名更新コマンド処理
  if (nameMatch) {
    var newName = nameMatch[1];
    updateUserName(userId, newName);
    var detectedLanguage = detectLanguage(text);

    // コンテキスト把握のため履歴に追加
    updateUserHistory(userId, text, detectedLanguage);

    var confirmationMessage = "名前を「" + newName + "」に更新しました。";
    var translatedUserMessage = "";

    try {
      var history = getUserHistory(userId);
      // ユーザーの入力を翻訳 (detectedLanguage -> Target)
      var translationResult = translateWithContext(text, history, detectedLanguage);
      translatedUserMessage = translationResult.translation;
    } catch (e) {
      debugToSheet('Name update translation failed: ' + e.message);
    }

    // 投稿を記録 (翻訳結果も含める)
    recordPost(messageId, timestamp, userId, roomId, text, hasPoll, translatedUserMessage);

    var messagesToSend = [];
    if (translatedUserMessage) {
      messagesToSend.push({
        "type": "text",
        "text": translatedUserMessage
      });
    }
    messagesToSend.push({
      "type": "text",
      "text": confirmationMessage
    });

    replyMessages(event.replyToken, messagesToSend);
    return;
  }

  // アンケートがある場合はFlex Messageを返信
  if (hasPoll) {
    var pollContent = text.replace(checkRegex, '').trim();
    var translatedPoll = '';
    var detectedLanguage = detectLanguage(text);

    // コンテキスト把握のため履歴に追加
    updateUserHistory(userId, text, detectedLanguage);

    // アンケート内容の翻訳
    if (pollContent) {
      try {
        var history = getUserHistory(userId);
        var translationResult = translateWithContext(pollContent, history, detectedLanguage);
        translatedPoll = translationResult.translation;
      } catch (e) {
        debugToSheet('Poll translation failed: ' + e.message);
      }
    }

    // 投稿を記録 (翻訳結果も含める)
    recordPost(messageId, timestamp, userId, roomId, text, hasPoll, translatedPoll);

    var messagesToSend = [];

    // 翻訳メッセージがある場合は先に追加
    if (translatedPoll) {
      messagesToSend.push({
        "type": "text",
        "text": translatedPoll
      });
    }

    // Flex Messageを追加
    var flexMessage = createPollFlexMessage(messageId);
    messagesToSend.push(flexMessage);

    replyMessages(event.replyToken, messagesToSend);
    return;
  }

  // 翻訳処理 (コマンドでもアンケートでもない場合)
  var translatedText = "";
  try {
    // 履歴取得
    var history = getUserHistory(userId);
    // 言語検出
    var detectedLanguage = detectLanguage(text);
    // 翻訳実行
    var translationResult = translateWithContext(text, history, detectedLanguage);
    translatedText = translationResult.translation;

    // 翻訳結果を返信
    replyMessages(event.replyToken, [{
      "type": "text",
      "text": translatedText
    }]);

    // 履歴更新
    updateUserHistory(userId, text, detectedLanguage);

    // ログ保存
    recordTranslationLog({
      timestamp: new Date(),
      userId: userId,
      language: detectedLanguage,
      originalMessage: text,
      translation: translationResult.translation,
      prompt: translationResult.prompt,
      historyCount: history.length
    });

  } catch (error) {
    // エラーハンドリング
    debugToSheet("Translation Error: " + error.message, error.stack);

    var errorMessage = '申し訳ございません。翻訳中にエラーが発生しました。';
    // レートリミットエラーの場合のメッセージ
    if (error.message && error.message.indexOf('RATE_LIMIT_EXCEEDED') !== -1) {
      errorMessage = 'AIサービスのレートリミットに到達しました。５分ほど置いて試してください';
    }

    replyMessages(event.replyToken, [{
      "type": "text",
      "text": errorMessage
    }]);
  }

  // 投稿を記録 (通常の翻訳)
  recordPost(messageId, timestamp, userId, roomId, text, hasPoll, translatedText);
}

/**
 * ポストバックイベントを処理する関数
 *
 * @param {Object} event LINEイベントオブジェクト
 */
function handlePostbackEvent(event) {
  var data = event.postback.data;
  var params = parseQuery(data);

  if (params['action'] === 'answer') {
    var userId = event.source.userId;
    var timestamp = new Date(event.timestamp);
    var answerValue = params['value'];
    var pollPostId = params['postId'];

    // アニメーションを表示 (5秒)
    sendLoadingAnimation(userId, 5);
    
    // 回答を記録
    recordAnswer(pollPostId, timestamp, userId, answerValue);
  }
}

/**
 * クエリ文字列をパースするヘルパー関数
 *
 * @param {string} queryString クエリ文字列 (key=value&key2=value2)
 * @returns {Object} パース結果のオブジェクト
 */
function parseQuery(queryString) {
  var query = {};
  var pairs = (queryString[0] === '?' ? queryString.substr(1) : queryString).split('&');
  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i].split('=');
    query[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1] || '');
  }
  return query;
}

/**
 * Webアプリケーションモジュール
 *
 * アンケート結果を表示するWebページを提供します。
 */

/**
 * HTTP GETリクエストを処理する関数
 *
 * @param {Object} e イベントオブジェクト
 */
function doGet(e) {
  try {
    // index.html からテンプレートを作成
    var template = HtmlService.createTemplateFromFile('index');

    var postId = e.parameter.postId;
    var results = [];

    // postId が指定されている場合、詳細結果を取得
    var pollData = { text: "", translatedText: "" };
    if (postId) {
      results = getPollResultDetails(postId);
      pollData = getPollContent(postId);
    }

    // テンプレート変数に値を設定
    template.postId = postId || "指定されていません";
    template.results = results;
    template.pollContent = pollData.text;
    template.translatedPollContent = pollData.translatedText;

    return template.evaluate()
        .setTitle('See results')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    debugToSheet(error.message, error.stack);
    return ContentService.createTextOutput("エラーが発生しました。");
  }
}

/* ==========================================================================================
 * 以下、翻訳機能モジュール
 * ========================================================================================== */

/**
 * ユーザー履歴取得
 */
function getUserHistory(userId) {
  try {
    var historyKey = 'HISTORY_' + userId;
    var historyJson = getScriptProperty(historyKey);

    if (!historyJson) {
      return [];
    }

    return JSON.parse(historyJson);
  } catch (error) {
    debugToSheet('getUserHistoryエラー: ' + error.toString());
    return [];
  }
}

/**
 * ユーザー履歴更新
 */
function updateUserHistory(userId, message, language) {
  try {
    var properties = PropertiesService.getScriptProperties();
    var historyKey = 'HISTORY_' + userId;

    var history = getUserHistory(userId);
    history.push({
      message: message,
      language: language,
      timestamp: new Date().getTime()
    });

    if (history.length > MAX_HISTORY_COUNT) {
      history = history.slice(-MAX_HISTORY_COUNT);
    }

    properties.setProperty(historyKey, JSON.stringify(history));
  } catch (error) {
    debugToSheet('updateUserHistoryエラー: ' + error.toString());
  }
}

/**
 * 言語検出
 */
function detectLanguage(text) {
  if (/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(text)) {
    return 'ja';
  }
  if (/[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]/.test(text)) {
    return 'pl';
  }
  return 'en';
}

/**
 * 文脈を考慮した翻訳
 */
function translateWithContext(message, history, sourceLanguage) {
  try {
    var targetLanguage = determineTargetLanguage(sourceLanguage);
    var prompt = buildTranslationPrompt(message, history, sourceLanguage, targetLanguage);
    var translation = callGeminiAPI(prompt);

    return {
      translation: translation,
      prompt: prompt
    };
  } catch (error) {
    throw error;
  }
}

/**
 * ターゲット言語決定
 */
function determineTargetLanguage(sourceLanguage) {
  if (sourceLanguage === 'ja') {
    return 'en';
  } else {
    return 'ja';
  }
}

/**
 * 翻訳プロンプト作成
 */
function buildTranslationPrompt(message, history, sourceLanguage, targetLanguage) {
  var prompt = '';
  if (sourceLanguage === 'ja') {
    prompt += 'あなたはプロの通訳アシスタントです。以下の日本語テキストを「英語」と「ポーランド語」の両方に翻訳してください。\n\n';
    prompt += '【出力形式】\n';
    prompt += 'Polish: [ポーランド語の翻訳結果]\n';
    prompt += 'English: [英語の翻訳結果]\n\n';
  } else {
    prompt += 'あなたはプロの通訳アシスタントです。以下のテキストを自然な日本語に翻訳してください。\n\n';
  }

  if (history && history.length > 0) {
    prompt += '【会話の文脈】\n';
    prompt += '以下は過去のユーザーの発言です。代名詞や省略表現を翻訳する際の参考にしてください。\n\n';
    history.forEach(function(item, index) {
      prompt += (index + 1) + '. ' + item.message + '\n';
    });
    prompt += '\n';
  }

  prompt += '【翻訳対象】\n';
  prompt += message + '\n\n';
  prompt += '【指示】\n';
  prompt += '- 翻訳結果のみを出力してください（説明や追加情報は不要）\n';
  prompt += '- 子供バレエ教室のチャットでのメッセージです。バレエ用語は正しく訳してください。ポーランド語は先生で、日本語は生徒の保護者です。バレエ教室の先生とのやりとりとして自然な文章にしてください。\n';
  prompt += '- 原文に含まれるニュアンス（感情、皮肉、丁寧さの度合い、ユーモアなど）を鋭敏に汲み取り、それをターゲット言語で適切に表現してください。直訳よりも、この「空気感」の再現を優先してください。\n';
  prompt += '- ポーランド人が言葉に込める親密さを表現してください\n';
  prompt += '- 翻訳した文章が長くなっても構いませんので、元の文章の意図が完全に伝わるようにしてください\n';

  if (history && history.length > 0) {
    prompt += '- 代名詞や省略表現は、上記の文脈を考慮して適切に翻訳してください\n';
  }

  return prompt;
}

/**
 * Gemini API呼び出し (リトライ機能付き)
 */
function callGeminiAPI(prompt) {
  var apiKey = getScriptProperty('GEMINI_API_KEY');

  var payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 8192
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var lastError = null;

  // モデルごとのループ
  for (var i = 0; i < GEMINI_MODELS.length; i++) {
    var model = GEMINI_MODELS[i];
    var url = GEMINI_BASE_URL + model + ':generateContent?key=' + apiKey;

    try {
      var response;
      var responseCode;
      var maxRetries = 3;

      // 503エラー用のリトライループ
      for (var attempt = 0; attempt < maxRetries; attempt++) {
        response = UrlFetchApp.fetch(url, options);
        responseCode = response.getResponseCode();

        if (responseCode !== 503) {
          break;
        }

        if (attempt < maxRetries - 1) {
          var waitTime = Math.floor(Math.random() * 3001) + 2000;
          Utilities.sleep(waitTime);
        }
      }

      var responseContent = response.getContentText();

      // 429 (Rate Limit) の場合は次のモデルへ
      if (responseCode === 429) {
        debugToSheet('Model ' + model + ' hit rate limit (429). Switching to next model.');
        lastError = new Error('RATE_LIMIT_EXCEEDED');
        continue; // 次のモデルへ
      }

      if (responseCode !== 200) {
        throw new Error('Gemini API error (' + model + '): ' + responseCode + ' - ' + responseContent);
      }

      var result = JSON.parse(responseContent);

      if (!result.candidates || result.candidates.length === 0) {
        throw new Error('No translation result from Gemini API (' + model + ')');
      }

      return result.candidates[0].content.parts[0].text.trim();

    } catch (error) {
      debugToSheet('callGeminiAPI error with ' + model + ': ' + error.toString());
      lastError = error;

      // エラーオブジェクトのメッセージにRATE_LIMITが含まれている場合も次へ
      if (error.message && error.message.indexOf('RATE_LIMIT_EXCEEDED') !== -1) {
        continue;
      }

      // その他のエラーの場合は、今のところ次のモデルを試さず終了する（APIキーエラー等のため）
      // ただし、モデル固有のエラー(404 Not Found等)の可能性もあるため、
      // 404の場合は次へ行くべきかもしれないが、今回は要件「429の時」に絞る。
      throw error;
    }
  }

  // 全てのモデルで失敗した場合
  throw lastError || new Error('All models failed.');
}

/**
 * 翻訳ログを記録する関数
 *
 * @param {Object} data ログデータ
 */
function recordTranslationLog(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.TRANSLATION_LOG);

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.TRANSLATION_LOG);
      sheet.appendRow([
        'timestamp',
        'user_id',
        'language',
        'original_message',
        'translation',
        'prompt',
        'history_count'
      ]);
    }

    sheet.appendRow([
      data.timestamp,
      data.userId,
      data.language,
      data.originalMessage,
      data.translation,
      data.prompt,
      data.historyCount
    ]);

  } catch (error) {
    debugToSheet('recordTranslationLogエラー: ' + error.toString());
  }
}
