// =============================================================================
// eVote - GASを用いたオンライン匿名投票システム
// コード.gs - サーバーサイドロジック全体
// =============================================================================

// ---- シート名定数 ----
var SHEET_SETTINGS   = '設定';
var SHEET_ROSTER     = '名簿とトークン';
var SHEET_BALLOT_BOX = '投票箱';
var SHEET_SYSLOG     = 'システム管理';

// ---- 名簿とトークンシートの列インデックス（1始まり） ----
var COL_EMAIL = 1; // A列: メールアドレス
var COL_TOKEN = 2; // B列: トークン
var COL_URL   = 3; // C列: 投票用URL
var COL_VOTED = 4; // D列: 投票済みフラグ

// ---- 設定シートの構造 ----
// A1: 主催者メール（カンマ区切り複数可）
// A2: 締め切り日時
// A3: 事務局パスワード
// A4: WebアプリURL（デプロイ後に手動で貼り付ける。例: https://script.google.com/macros/s/XXXXX/exec）
// A5: GitHub PagesのURL（例: https://{user}.github.io/{repo}）
//     ※ トークン付き投票URLの生成に使用。末尾スラッシュなしで入力。
// B列以降: 各投票。1行目がタイトル、2行目以降が選択肢。列が空になるまで続く。
//   例) B1=投票タイトル1, B2~=選択肢 / C1=投票タイトル2, C2~=選択肢 / D1=...
// ※ B列は必須。C列以降は1行目（タイトル）が空であれば存在しないとみなす。

// =============================================================================
// メニュー登録
// =============================================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('eVote 管理')
    .addItem('① トークン生成 ＋ 案内メール送信（GAS直接）', 'generateTokensAndSendEmails')
    .addItem('② トークン生成のみ（URLリスト作成・外部メーラー用）', 'generateTokensOnly')
    .addSeparator()
    .addItem('③ 集計・結果通知（手動実行）', 'tallySendResults')
    .addItem('④ 締め切りトリガーをセットアップ', 'setupTrigger')
    .addToUi();
}

// =============================================================================
// エントリポイント: doGet / doPost
// =============================================================================

/**
 * GETリクエスト: 生存確認用のシンプルなレスポンスを返す。
 * HTML配信は GitHub Pages が担当するため、ここでは何も返さない。
 *
 * @param {Object} e
 * @returns {TextOutput}
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POSTリクエスト: JSON API として動作する。
 * リクエストボディの JSON に action フィールドで処理を振り分ける。
 *
 * action: 'getVoteFormData' → トークン検証・フォームデータ取得
 * action: 'submitVote'      → 投票記録
 * action: 'verifyPassword'  → 事務局パスワード検証
 * action: 'registerAttendee'→ 当日参加者登録
 *
 * @param {Object} e - GASイベントオブジェクト
 * @returns {TextOutput} JSON レスポンス
 */
function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';

    if (action === 'getVoteFormData') {
      result = getVoteFormData(body.token);

    } else if (action === 'submitVote') {
      result = submitVote(body.token, body.choices);

    } else if (action === 'verifyPassword') {
      result = { success: verifyAdminPassword(body.password) };

    } else if (action === 'registerAttendee') {
      result = registerAttendee(body.email, body.password);

    } else {
      result = { success: false, message: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'サーバーエラー: ' + err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// 機能A: 事前準備・案内メール送信（カスタムメニューから実行）
// =============================================================================

/**
 * 名簿とトークンシートのA列にある未処理（B列が空）のアドレスに対し、
 * トークンとURLを生成し、GASのMailAppを使って案内メールを直接送信する。
 * 処理件数をダイアログで通知する。
 */
function generateTokensAndSendEmails() {
  var settings    = _getSettings();
  var pagesUrl    = settings.pagesUrl || settings.appUrl; // GitHub Pages URL（なければGAS URLにフォールバック）
  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();
  var sentCount   = 0;

  for (var i = 2; i <= lastRow; i++) {
    var email         = rosterSheet.getRange(i, COL_EMAIL).getValue();
    var existingToken = rosterSheet.getRange(i, COL_TOKEN).getValue();

    // メールアドレスがあり、まだトークンが未生成の行だけ処理
    if (email && !existingToken) {
      var voteUrl = _generateTokenForRow(i, rosterSheet, pagesUrl);
      _sendInvitationEmail(email, voteUrl, settings);
      sentCount++;
    }
  }

  SpreadsheetApp.getUi().alert(
    sentCount + ' 件のアドレスにトークンを生成し、案内メールを送信しました。'
  );
}

/**
 * 名簿とトークンシートの未処理アドレスに対してトークンとURLを生成するが、
 * メールは送信しない（外部メーラーでの差し込み送信用）。
 * C列（URL）が更新されたことをダイアログで通知する。
 */
function generateTokensOnly() {
  var settings    = _getSettings();
  var pagesUrl    = settings.pagesUrl || settings.appUrl;
  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();
  var genCount    = 0;

  for (var i = 2; i <= lastRow; i++) {
    var email         = rosterSheet.getRange(i, COL_EMAIL).getValue();
    var existingToken = rosterSheet.getRange(i, COL_TOKEN).getValue();

    if (email && !existingToken) {
      _generateTokenForRow(i, rosterSheet, pagesUrl);
      genCount++;
    }
  }

  SpreadsheetApp.getUi().alert(
    genCount + ' 件のトークンとURLを生成しました。\n' +
    '「名簿とトークン」シートのC列（URL）をご確認ください。'
  );
}

// =============================================================================
// 機能B: 事務局ポータル（google.script.run 経由で呼び出し）
// =============================================================================

/**
 * 事務局用パスワードを検証する。
 * 設定シートのA3セルの値と照合し、一致すれば true を返す。
 *
 * @param {string} password - 入力されたパスワード
 * @returns {boolean} 認証結果
 */
function verifyAdminPassword(password) {
  var settings = _getSettings();
  return String(settings.adminPassword).trim() === String(password).trim();
}

/**
 * 当日参加者を名簿に即時登録し、トークン付き投票URLを発行して返す。
 * 同時に、参加者のアドレス宛に案内メールを送信する。
 * 締め切り後・同一アドレスの重複登録も適切にハンドリングする。
 * GAS URL が公開されているため、事務局パスワードをサーバー側でも検証する。
 *
 * @param {string} email    - 登録するメールアドレス
 * @param {string} password - 事務局パスワード（必須）
 * @returns {Object} { success: boolean, url: string, message: string }
 */
function registerAttendee(email, password) {
  // サーバー側でもパスワードを検証（GAS URL が公開されているため二重チェック）
  if (!verifyAdminPassword(password)) {
    return { success: false, url: '', message: '認証エラー：パスワードが正しくありません。' };
  }

  // 簡易バリデーション
  if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
    return { success: false, url: '', message: '有効なメールアドレスを入力してください。' };
  }

  var settings = _getSettings();
  var appUrl   = settings.appUrl;
  var pagesUrl = settings.pagesUrl || appUrl;

  // 締め切りチェック
  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    return { success: false, url: '', message: '投票の締め切りを過ぎています。登録できません。' };
  }

  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();

  // 二重登録チェック: 既に同一アドレスが存在すれば既存URLを返す
  for (var i = 2; i <= lastRow; i++) {
    if (rosterSheet.getRange(i, COL_EMAIL).getValue() === email) {
      var existingUrl = rosterSheet.getRange(i, COL_URL).getValue();
      return {
        success: true,
        url:     existingUrl,
        message: 'このアドレスは既に登録済みです。既存のURLを返します。'
      };
    }
  }

  // 新規行に追加してトークンを生成
  var newRow  = lastRow + 1;
  rosterSheet.getRange(newRow, COL_EMAIL).setValue(email);
  var voteUrl = _generateTokenForRow(newRow, rosterSheet, pagesUrl);

  // 案内メール送信（失敗しても登録自体は成功とする）
  try {
    _sendInvitationEmail(email, voteUrl, settings);
  } catch (mailErr) {
    Logger.log('メール送信エラー: ' + mailErr.message);
    return {
      success: true,
      url:     voteUrl,
      message: '登録完了。ただし案内メールの送信に失敗しました。URLをコピーして直接お伝えください。'
    };
  }

  return { success: true, url: voteUrl, message: '登録完了。案内メールを送信しました。' };
}

// =============================================================================
// 機能C: 投票フォームデータ取得・投票記録（google.script.run 経由で呼び出し）
// =============================================================================

/**
 * 投票フォームの表示に必要な情報を取得して返す。
 * index.html のページ読み込み時に google.script.run から呼び出す。
 *
 * @param {string} token - 有権者トークン
 * @returns {Object} {
 *   valid: boolean,
 *   message: string,
 *   settings: {
 *     votes: Array<{ title: string, options: string[] }>,  // 投票ごとの設定（1件以上）
 *     deadline: string
 *   } | null
 * }
 */
function getVoteFormData(token) {
  var validation = _validateToken(token);
  if (!validation.valid) {
    return { valid: false, message: validation.message, settings: null };
  }

  var settings = _getSettings();
  return {
    valid:   true,
    message: 'OK',
    settings: {
      votes:    settings.votes,
      deadline: settings.deadline
                  ? Utilities.formatDate(new Date(settings.deadline), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
                  : ''
    }
  };
}

/**
 * 投票を記録する。index.html の google.script.run から呼び出す。
 * 内部で _recordVote() を呼び出し、結果をオブジェクトで返す。
 *
 * @param {string}   token   - 有権者トークン
 * @param {string[]} choices - 各投票の選択肢の配列。votes の順序に対応する。
 * @returns {Object} { success: boolean, message: string }
 */
function submitVote(token, choices) {
  try {
    _recordVote(token, choices);
    return { success: true, message: '投票が完了しました。ご参加ありがとうございました。' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 投票の有効性を確認し、LockServiceで排他制御しながら投票箱に記録する。
 * 投票箱には「日時・各投票の選択肢」のみを記録し、匿名性を担保する。
 * 投票数は動的（設定シートの投票列数に追従する）。
 *
 * @param {string}   token   - 有権者トークン
 * @param {string[]} choices - 各投票の選択肢の配列
 * @throws {Error} 検証失敗・締め切り超過・二重投票・未選択の場合
 */
function _recordVote(token, choices) {
  // トークン検証（締め切り・有効性）
  var validation = _validateToken(token);
  if (!validation.valid) {
    throw new Error(validation.message);
  }

  var settings = _getSettings();

  // 各投票の選択肢が入力されているかチェック
  for (var v = 0; v < settings.votes.length; v++) {
    if (!choices[v] || choices[v] === '') {
      throw new Error('「' + settings.votes[v].title + '」の選択肢を選んでください。');
    }
  }

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  var ballotSheet = ss.getSheetByName(SHEET_BALLOT_BOX);

  // LockService で排他制御（最大30秒待機）
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // ロック取得後に再度投票済みフラグを確認（競合状態での二重投票を防止）
    var alreadyVoted = rosterSheet.getRange(validation.row, COL_VOTED).getValue();
    if (alreadyVoted === true) {
      throw new Error('このトークンは既に使用されています。');
    }

    // ---- 匿名性の核心: 投票箱にはトークン・メールを一切記録しない ----
    // 1列目: 日時、2列目以降: 各投票の選択肢
    var row = [new Date()];
    for (var i = 0; i < settings.votes.length; i++) {
      row.push(choices[i] || '');
    }
    ballotSheet.appendRow(row);

    // 投票済みフラグをセット
    rosterSheet.getRange(validation.row, COL_VOTED).setValue(true);

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/**
 * トークンの有効性と締め切り日時を検証する。
 *
 * @param {string} token - 検証するトークン
 * @returns {Object} { valid: boolean, message: string, row: number|null }
 */
function _validateToken(token) {
  if (!token) {
    return { valid: false, message: '投票URLが正しくありません。', row: null };
  }

  var settings = _getSettings();

  // 締め切りチェック
  if (settings.deadline && new Date() > new Date(settings.deadline)) {
    return { valid: false, message: '投票の受付は終了しました。', row: null };
  }

  // 名簿シートでトークンを照合
  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  var lastRow     = rosterSheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var rowToken = rosterSheet.getRange(i, COL_TOKEN).getValue();
    if (rowToken === token) {
      var voted = rosterSheet.getRange(i, COL_VOTED).getValue();
      if (voted === true) {
        return { valid: false, message: 'このURLはすでに使用済みです。', row: i };
      }
      return { valid: true, message: 'OK', row: i };
    }
  }

  return { valid: false, message: '有効なトークンが見つかりません。', row: null };
}

// =============================================================================
// 機能D: 自動集計・結果通知
// =============================================================================

/**
 * 投票箱を集計し、主催者メールアドレスに結果を通知する。
 * システム管理シートのA1セル（集計完了フラグ）を確認し、二重集計を防止する。
 * 投票数は設定シートの列数に追従する（3件以上にも対応）。
 * 時間主導型トリガーからの呼び出しと、カスタムメニューからの手動実行の両方に対応する。
 */
function tallySendResults() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sysSheet = ss.getSheetByName(SHEET_SYSLOG);

  // 集計完了フラグのチェック（A1セル）
  var alreadyDone = sysSheet.getRange('A1').getValue();
  if (alreadyDone === true) {
    Logger.log('集計は既に完了済みです。重複実行をスキップしました。');
    try {
      SpreadsheetApp.getUi().alert('集計は既に完了しています（重複実行を防止しました）。');
    } catch (e) { /* トリガー実行時はUIなし */ }
    return;
  }

  var settings    = _getSettings();
  var ballotSheet = ss.getSheetByName(SHEET_BALLOT_BOX);
  var lastRow     = ballotSheet.getLastRow();

  if (lastRow < 1) {
    Logger.log('集計: 投票データがありません。');
    try {
      SpreadsheetApp.getUi().alert('投票データがまだありません。');
    } catch (e) { /* トリガー実行時はUIなし */ }
    return;
  }

  var numVotes = settings.votes.length;

  // 各投票の集計オブジェクトを初期化
  // tallies[v] = { '選択肢A': 0, '選択肢B': 0, ... }
  var tallies = settings.votes.map(function(vote) {
    var t = {};
    vote.options.forEach(function(opt) { t[opt] = 0; });
    return t;
  });

  // 投票箱を全行スキャン（1列目=日時、2列目以降=各投票の選択肢）
  // ヘッダー行なし運用のため1行目から読み込む。日時が空の行は自動スキップ。
  var totalVotes = 0;
  var allData = ballotSheet.getRange(1, 1, lastRow, numVotes + 1).getValues();
  allData.forEach(function(rowData) {
    if (!rowData[0]) return; // 日時が空の行はスキップ
    totalVotes++;
    for (var v = 0; v < numVotes; v++) {
      var choice = rowData[v + 1]; // 0列目=日時なので+1
      if (choice && tallies[v].hasOwnProperty(choice)) {
        tallies[v][choice]++;
      }
    }
  });

  // ---- 結果メール本文の組み立て ----
  var now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var body = '【eVote 投票結果通知】\n\n';
  body += '集計日時: ' + now + '\n';
  body += '総投票数: ' + totalVotes + ' 票\n\n';

  settings.votes.forEach(function(vote, idx) {
    body += '■ ' + vote.title + '\n';
    // 票数の多い順に並べ替えて出力
    Object.keys(tallies[idx])
      .sort(function(a, b) { return tallies[idx][b] - tallies[idx][a]; })
      .forEach(function(opt) {
        body += '  ' + opt + ': ' + tallies[idx][opt] + ' 票\n';
      });
    body += '\n';
  });

  // 主催者へメール送信（1件目はTO、2件目以降はBCC）
  if (settings.organizerEmails.length > 0) {
    var mailOptions = {
      to:      settings.organizerEmails[0],
      subject: '【eVote】投票結果のお知らせ',
      body:    body
    };
    if (settings.organizerEmails.length > 1) {
      mailOptions.bcc = settings.organizerEmails.slice(1).join(',');
    }
    MailApp.sendEmail(mailOptions);
  }

  // 集計完了フラグをセット（二重集計防止）
  sysSheet.getRange('A1').setValue(true);
  SpreadsheetApp.flush();

  Logger.log('集計・通知が完了しました。\n' + body);

  // UIから呼ばれた場合はダイアログ表示
  try {
    SpreadsheetApp.getUi().alert('集計・結果通知が完了しました。\n\n' + body);
  } catch (e) {
    // 時間主導型トリガーからの呼び出し時はUIが使えないため無視
  }
}

// =============================================================================
// トリガーセットアップ
// =============================================================================

/**
 * 設定シートの締め切り日時に基づいて、tallySendResults を実行する
 * 時間主導型トリガーを登録する。
 * 既存の同名トリガーは削除してから新しいものを登録する。
 * 締め切りの5分後に実行されるよう設定する。
 */
function setupTrigger() {
  var settings = _getSettings();

  if (!settings.deadline) {
    SpreadsheetApp.getUi().alert(
      '設定シートのA2に締め切り日時が設定されていません。\n例: 2026/04/10 12:00'
    );
    return;
  }

  var deadline = new Date(settings.deadline);
  if (isNaN(deadline.getTime())) {
    SpreadsheetApp.getUi().alert(
      '締め切り日時の形式が正しくありません。\n例: 2026/04/10 12:00'
    );
    return;
  }

  // 既存の tallySendResults トリガーをすべて削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tallySendResults') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 締め切りの5分後に集計トリガーを設定
  var triggerTime = new Date(deadline.getTime() + 5 * 60 * 1000);
  ScriptApp.newTrigger('tallySendResults')
    .timeBased()
    .at(triggerTime)
    .create();

  var triggerTimeStr = Utilities.formatDate(triggerTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  SpreadsheetApp.getUi().alert(
    '集計トリガーをセットしました。\n実行予定日時: ' + triggerTimeStr
  );
}

// =============================================================================
// ユーティリティ（内部関数）
// =============================================================================

/**
 * 設定シートからすべての設定値を読み込んでオブジェクトで返す。
 *
 * 設定シートのレイアウト:
 *   A1: 主催者メール（カンマ区切り複数可）
 *   A2: 締め切り日時
 *   A3: 事務局パスワード
 *   A4: WebアプリURL（デプロイ後のURLを手動で貼り付ける）
 *   A5: GitHub PagesのベースURL（例: https://user.github.io/repo）
 *   B列: 投票1（B1=タイトル、B2以降=選択肢）
 *   C列: 投票2（C1=タイトル、C2以降=選択肢）※空欄なら存在しない
 *   D列以降: 投票3, 4, ... （同様。タイトルが空欄になった列で終了）
 *
 * @returns {Object} 設定オブジェクト
 *   {string[]}   organizerEmails  - 主催者メールアドレス一覧
 *   {Date|''}    deadline         - 締め切り日時
 *   {string}     adminPassword    - 事務局用パスワード
 *   {string}     appUrl           - WebアプリURL（A4セル）
 *   {string}     pagesUrl         - GitHub PagesのベースURL（A5セル）
 *   {Array<{title:string, options:string[]}>} votes - 投票ごとの設定（1件以上）
 */
function _getSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);

  // A列: 固定設定項目
  var organizerEmails = _parseEmails(sheet.getRange('A1').getValue());
  var deadline        = sheet.getRange('A2').getValue();
  var adminPassword   = sheet.getRange('A3').getValue();
  // A4: WebアプリURL（必須）。ScriptApp.getService().getUrl() はアカウント依存で
  // 誤ったURLを返す場合があるため、ここでは一切使わない。
  var appUrl = _normalizeAppUrl(String(sheet.getRange('A4').getValue()).trim());
  // A5: GitHub PagesのベースURL（トークン付き投票URL生成に使用）
  var pagesUrl = String(sheet.getRange('A5').getValue()).trim();

  var maxRow = sheet.getLastRow();

  // B列（col=2）以降を動的にスキャンして投票設定を構築する。
  // タイトルセル（1行目）が空になった列で走査を終了する。
  var votes   = [];
  var col     = 2; // B列から開始

  while (true) {
    var title = sheet.getRange(1, col).getValue();

    // タイトルが空なら、その列以降には投票が存在しないとみなして終了
    if (title === '' || title === null || title === undefined) break;

    // 2行目以降の選択肢を収集
    var options = [];
    for (var r = 2; r <= maxRow; r++) {
      var val = sheet.getRange(r, col).getValue();
      if (val !== '' && val !== null && val !== undefined) {
        options.push(String(val));
      }
    }

    // タイトルはあるが選択肢が0件の列はスキップ（設定ミス対策）
    if (options.length > 0) {
      votes.push({ title: String(title), options: options });
    }

    col++;
  }

  return {
    organizerEmails: organizerEmails,
    deadline:        deadline,
    adminPassword:   adminPassword,
    appUrl:          appUrl,
    pagesUrl:        pagesUrl,
    votes:           votes
  };
}

/**
 * UUID v4 を生成して返す。
 * GAS標準の Utilities.getUuid() を利用する。
 *
 * @returns {string} UUID文字列（例: "550e8400-e29b-41d4-a716-446655440000"）
 */
function _generateUUID() {
  return Utilities.getUuid();
}

/**
 * 名簿とトークンシートの指定行に UUID トークンと投票用URLを生成して書き込み、
 * 生成した投票用URLを返す。
 *
 * pagesUrl は GitHub Pages のベースURL（例: https://user.github.io/repo）を指定する。
 * 生成される投票URLは `{pagesUrl}/index.html?token={uuid}` の形式になる。
 *
 * @param {number} row      - 対象行番号（1始まり）
 * @param {Sheet}  sheet    - 名簿とトークンシートオブジェクト
 * @param {string} pagesUrl - GitHub PagesのベースURL（末尾スラッシュなし）
 * @returns {string} 生成した投票用URL
 */
function _generateTokenForRow(row, sheet, pagesUrl) {
  var token   = _generateUUID();
  // GitHub Pages の index.html に token パラメータを付与した URL を生成する
  var base    = String(pagesUrl).replace(/\/$/, ''); // 末尾スラッシュを除去
  var voteUrl = base + '/index.html?token=' + token;

  sheet.getRange(row, COL_TOKEN).setValue(token);
  sheet.getRange(row, COL_URL).setValue(voteUrl);
  sheet.getRange(row, COL_VOTED).setValue(false);

  SpreadsheetApp.flush();
  return voteUrl;
}

/**
 * 有権者に案内メールを送信する。
 * 締め切り日時を日本語形式でフォーマットして本文に含める。
 *
 * @param {string} email    - 宛先メールアドレス
 * @param {string} voteUrl  - 専用投票URL
 * @param {Object} settings - _getSettings() の戻り値
 */
function _sendInvitationEmail(email, voteUrl, settings) {
  var deadlineStr = settings.deadline
    ? Utilities.formatDate(new Date(settings.deadline), 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm')
    : '（未設定）';

  var subject = '【eVote】投票のご案内';
  var body = [
    'このたびは投票へのご参加ありがとうございます。',
    '',
    '以下の専用URLからご投票ください。',
    'このURLはあなた専用です。他の方と共有しないでください。',
    '',
    '▼ 投票URL',
    voteUrl,
    '',
    '締め切り: ' + deadlineStr,
    '',
    '※ このURLは一度しか使用できません。',
    '※ 投票は匿名で処理されます（誰が何に投票したかは記録されません）。'
  ].join('\n');

  MailApp.sendEmail({ to: email, subject: subject, body: body });
}

/**
 * GAS WebアプリURLを、ブラウザのアカウント状態に依存しない正規形式に変換する。
 *
 * GASはアクセスするアカウントの状態によって以下のような形式のURLを生成する場合がある：
 *   - /a/{domain}/macros/s/{id}/exec   （Google Workspaceドメイン経由）
 *   - /macros/u/{N}/s/{id}/exec        （複数アカウントのN番目）
 * これらをすべて標準形式 /macros/s/{id}/exec に統一する。
 *
 * @param {string} url - 変換対象のURL
 * @returns {string} 正規化されたURL
 */
function _normalizeAppUrl(url) {
  if (!url) return '';

  // /a/{domain}/macros/s/{id}/exec → /macros/s/{id}/exec
  url = url.replace(/\/a\/[^\/]+\/macros\//, '/macros/');

  // /macros/u/{N}/s/{id}/exec → /macros/s/{id}/exec
  url = url.replace(/\/macros\/u\/\d+\/s\//, '/macros/s/');

  return url;
}

/**
 * カンマ区切りのメールアドレス文字列を配列に変換する。
 * 各アドレスの前後にある空白・全角スペースを除去し、空文字は除外する。
 * セル値が数値や Date 型で入っている場合も文字列に変換して処理する。
 *
 * 例: "a@example.com, b@example.com , c@example.com"
 *   → ["a@example.com", "b@example.com", "c@example.com"]
 *
 * @param {string} raw - A1セルの生の値（カンマ区切り複数可）
 * @returns {string[]} メールアドレスの配列
 */
function _parseEmails(raw) {
  if (!raw) return [];
  return String(raw)
    .split(',')
    .map(function(addr) { return addr.replace(/[\s\u3000]+/g, ''); }) // 半角・全角スペース除去
    .filter(function(addr) { return addr.length > 0; });
}
