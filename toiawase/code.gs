/**
 * 問い合わせフォームの自動化（Google Apps Script）
 *
 * できること
 *   1. Googleフォームに問い合わせが届くと、スプレッドシートに自動でたまる
 *   2. 受付番号と「対応状況」の欄を自動でつける
 *   3. 担当者に通知メールを送る
 *   4. お客様に「受け付けました」の自動返信メールを送る
 *
 * 使い方
 *   新しいスプレッドシートを作り、「拡張機能」→「Apps Script」にこのコードを貼る。
 *   下の設定（NOTIFY_TO など）を書き換えて、関数「setup」を1回だけ実行する。
 *   フォームとトリガー（自動で動くしくみ）が作られる。
 */

// ===== 設定（ここを書き換える） =====
const NOTIFY_TO = 'tantou@example.com';       // 通知を受け取る担当者のメールアドレス
const SHOP_NAME = 'サンプル工務店';            // 自動返信メールに出す名前
const FORM_TITLE = 'お問い合わせフォーム';
const KINDS = ['見積りの依頼', '工事のご相談', 'アフターサービス', 'その他'];
const SEND_AUTO_REPLY = true;                  // お客様への自動返信を送るなら true

// 回答シートで使う見出し（フォームの質問名と同じにする）
const Q_NAME = 'お名前';
const Q_MAIL = 'メールアドレス';
const Q_KIND = 'お問い合わせの種類';
const Q_BODY = 'お問い合わせ内容';
const COL_NO = '受付番号';
const COL_STATUS = '対応状況';

/**
 * 最初に1回だけ実行する。フォームを作ってこのスプレッドシートにつなぎ、
 * 回答が届いたときに onFormSubmit が動くように設定する。
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const form = FormApp.create(FORM_TITLE)
    .setDescription(SHOP_NAME + 'へのお問い合わせはこちらから。2営業日以内にご連絡します。')
    .setCollectEmail(false);
  form.addTextItem().setTitle(Q_NAME).setRequired(true);
  form.addTextItem().setTitle(Q_MAIL).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build());
  form.addMultipleChoiceItem().setTitle(Q_KIND).setChoiceValues(KINDS).setRequired(true);
  form.addParagraphTextItem().setTitle(Q_BODY).setRequired(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 同じトリガーが二重にできないよう、古いものを消してから作る
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();

  Logger.log('フォームのURL（お客様に案内する）: ' + form.getPublishedUrl());
  Logger.log('フォームの編集URL: ' + form.getEditUrl());
}

/**
 * フォームに回答が届くたびに自動で動く。
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const answer = name => (e.namedValues[name] || [''])[0];

  // 受付番号と対応状況の列がなければ作る
  const noCol = ensureColumn_(sheet, COL_NO);
  const statusCol = ensureColumn_(sheet, COL_STATUS);

  const receiptNo = 'Q-' + ('0000' + (row - 1)).slice(-4);
  sheet.getRange(row, noCol).setValue(receiptNo);
  const status = sheet.getRange(row, statusCol);
  status.setValue('未対応');
  status.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['未対応', '対応中', '完了'], true).build());

  const name = answer(Q_NAME);
  const mail = answer(Q_MAIL);
  const kind = answer(Q_KIND);
  const body = answer(Q_BODY);

  // 担当者への通知
  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject: '【問い合わせ ' + receiptNo + '】' + kind + '（' + name + ' 様）',
    body: [
      '新しい問い合わせが届きました。',
      '',
      '受付番号: ' + receiptNo,
      'お名前: ' + name,
      'メールアドレス: ' + mail,
      '種類: ' + kind,
      '',
      '内容:',
      body,
      '',
      '一覧: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    ].join('\n'),
  });

  // お客様への自動返信
  if (SEND_AUTO_REPLY && mail) {
    MailApp.sendEmail({
      to: mail,
      subject: '【' + SHOP_NAME + '】お問い合わせを受け付けました（' + receiptNo + '）',
      body: [
        name + ' 様',
        '',
        'お問い合わせいただき、ありがとうございます。',
        '以下の内容で受け付けました。2営業日以内に担当者よりご連絡します。',
        '',
        '受付番号: ' + receiptNo,
        '種類: ' + kind,
        '内容:',
        body,
        '',
        '※このメールは自動で送信しています。',
        SHOP_NAME,
      ].join('\n'),
    });
  }
}

/** 見出し行に name の列がなければ右端に足し、その列番号を返す */
function ensureColumn_(sheet, name) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(name);
  if (idx >= 0) return idx + 1;
  sheet.getRange(1, lastCol + 1).setValue(name);
  return lastCol + 1;
}
