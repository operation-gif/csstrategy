/**
 * 売上方程式 診断ツール（Google Apps Script）
 *
 * 「売上方程式 診断マスタ」スプレッドシートに組み込んで使う。
 *  1. 設問マスタからGoogleフォームを作成する
 *  2. フォームの回答を0〜4点に変換して「回答」シートに書き込む
 *  3. 企業ごとの診断シート（別ファイル）をDriveのフォルダに作成する
 *  4. SlackのCSチャンネルに「〇〇株式会社 診断結果が出ました」と通知する
 *
 * 設定手順は README.md を参照。
 */

const CONFIG = {
  SHEET_MASTER: '設問マスタ',
  SHEET_HABITS: '口癖チェック',
  SHEET_MAP: '施策マッピング',
  SHEET_ANSWER: '回答',
  SHEET_RESULT: '診断結果',
  FORM_TITLE: '売上方程式 診断アンケート',
  FOLDER_NAME: '売上方程式 診断シート',
  // 診断シートの格納先フォルダのID（フォルダURLの /folders/ の後ろの文字列）。
  // 空のときは、診断マスタと同じ場所に「売上方程式 診断シート」フォルダを自動で作る。
  REPORT_FOLDER_ID: '',
  ROLE_QID: '③-2',      // 複数選択で、選んだ数が点数になる設問
  CHANNEL_QID: '①-4',   // 直後に「対象チャネル」を聞く設問
  FONT: 'Noto Sans JP',
  COLOR_HEADER: '#F6B26B',
  COLOR_WEAK: '#F4CCCC',
  TIMEZONE: 'Asia/Tokyo',
};

// 回答シートの列（1始まり）
const ANSWER_COL = {
  DATE: 1, COMPANY: 2, PRODUCT: 3, POSITION: 4,
  FIRST_SCORE: 5, // E列（①-1）から T列（④-4）まで
  HABITS: 21,     // U列
  CHANNEL: 22,    // V列
  REPORT_URL: 23, // W列
};

const POSITION_OPTIONS = ['代表・役員', '販売・営業', 'マーケティング・広報', '商品開発', 'その他'];

const ROLE_OPTIONS = [
  '信頼づくり（有名店・百貨店・認定マークなど）',
  '売上をつくる（EC・卸・スーパーなど）',
  '試してもらう（試食・サンプル配布・イベントなど）',
  '価格の理由を伝える（接客・説明のある売り場）',
];

const CHANNEL_OPTIONS = [
  '卸先のスタッフ',
  '委託先のスタッフ',
  'イベントの販売スタッフ',
  '協会店舗のスタッフ',
  'その他',
  '自社スタッフ以外が売ったことはない',
];

const SECTION_TEXT = {
  '①商品の魅力': 'お客様に「何の商品か」が伝わる言葉になっているかをお聞きします。',
  '②売り方（伝え方）': '誰に向けて、どんな順番で伝えているかをお聞きします。',
  '③売り場': 'どこで、どんな役割の売り場で売っているかをお聞きします。',
  '④競合': 'お客様がこの商品と比べているものは何かをお聞きします。',
};

const FORM_DESCRIPTION = [
  '日本マタニティフード協会の契約企業様向け「売上方程式 診断」です。',
  '',
  '売上 ＝ ①商品の魅力 × ②売り方（伝え方） × ③売り場 × ④競合',
  '',
  'ご回答いただくと、4つの項目のうち御社の商品でいちばん伸びしろのある項目と、その項目を伸ばすためのおすすめ施策を、担当よりお返しします。',
  '',
  '・所要時間：約10分（全21問）',
  '・1回のご回答で1商品についてお答えください。複数の商品を診断したい場合は、商品ごとにご回答ください。',
  '・良い・悪いを評価するものではありません。今の状態にいちばん近いものをお選びください。',
  '・ご回答内容は、協会内で施策のご提案にのみ使用します。',
].join('\n');

const CONFIRMATION_MESSAGE =
  'ご回答ありがとうございました。診断結果とおすすめ施策は、担当よりご連絡いたします。';

// ---------------------------------------------------------------------------
// メニュー
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('診断ツール')
    .addItem('1. Slackの通知先を設定', 'setSlackWebhook')
    .addItem('2. 診断フォームを作成', 'setupForm')
    .addSeparator()
    .addItem('Slack通知のテスト', 'testSlack')
    .addItem('指定した行の診断シートを作り直す', 'rebuildReportForRow')
    .addToUi();
}

function setSlackWebhook() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Slackの通知先',
    'CSチャンネル用のIncoming Webhook URL（https://hooks.slack.com/services/…）を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const url = res.getResponseText().trim();
  if (!/^https:\/\/hooks\.slack\.com\//.test(url)) {
    ui.alert('Slack の Incoming Webhook URL ではないようです。もう一度お試しください。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URL', url);
  ui.alert('通知先を保存しました。「Slack通知のテスト」で届くか確認できます。');
}

function testSlack() {
  const sent = postToSlack_('【テスト】売上方程式 診断ツールからの通知テストです。');
  SpreadsheetApp.getUi().alert(sent
    ? 'Slackにテスト通知を送りました。CSチャンネルを確認してください。'
    : 'Slackの通知先がまだ設定されていません。「1. Slackの通知先を設定」から設定してください。');
}

function rebuildReportForRow() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('診断シートを作り直す', '回答シートの行番号を入力してください（例：3）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const row = parseInt(res.getResponseText(), 10);
  if (!(row >= 2)) {
    ui.alert('2以上の行番号を入力してください。');
    return;
  }
  const report = buildReport_(row);
  const sent = notifyResult_(report);
  ui.alert('診断シートを作成しました' + (sent ? '（Slackにも通知しました）' : '') + '。\n' + report.url);
}

// ---------------------------------------------------------------------------
// フォーム作成
// ---------------------------------------------------------------------------

function setupForm() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FORM_ID')) {
    const ok = ui.alert('フォームはすでに作成済みです',
      '新しいフォームを作り直しますか？（今のフォームへの回答の自動処理は止まります）',
      ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
    removeFormTriggers_();
  }

  const ss = ss_();
  const questions = readMaster_();
  const habits = readHabits_();

  const form = FormApp.create(CONFIG.FORM_TITLE)
    .setDescription(FORM_DESCRIPTION)
    .setConfirmationMessage(CONFIRMATION_MESSAGE)
    .setCollectEmail(false)
    .setAllowResponseEdits(false)
    .setShowLinkToRespondAgain(true);

  const itemMap = {}; // フォームの設問ID → 設問マスタのID（または種別）

  form.addSectionHeaderItem().setTitle('はじめに');
  itemMap[form.addTextItem().setTitle('会社名').setRequired(true).getId()] = 'COMPANY';
  itemMap[form.addTextItem().setTitle('診断する商品名').setRequired(true).getId()] = 'PRODUCT';
  itemMap[form.addMultipleChoiceItem().setTitle('ご回答者の役割')
    .setChoiceValues(POSITION_OPTIONS).setRequired(true).getId()] = 'POSITION';

  let currentCategory = '';
  questions.forEach(q => {
    if (q.category !== currentCategory) {
      currentCategory = q.category;
      form.addSectionHeaderItem()
        .setTitle(q.category)
        .setHelpText(SECTION_TEXT[q.category] || '');
    }

    if (q.id === CONFIG.ROLE_QID) {
      const item = form.addCheckboxItem()
        .setTitle('今の売り場で、次の役割を持っているものをすべて選んでください。')
        .setHelpText('当てはまるものがなければ、選ばずに次へ進んでください。')
        .setChoiceValues(ROLE_OPTIONS)
        .setRequired(false);
      itemMap[item.getId()] = q.id;
      return;
    }

    const item = form.addMultipleChoiceItem()
      .setTitle(q.text)
      .setHelpText(q.id === CONFIG.CHANNEL_QID
        ? '卸先がない場合は、イベントの販売スタッフや協会店舗のスタッフが売った場合でお答えください。'
        : '今の状態にいちばん近いものを1つ選んでください。')
      .setChoiceValues(q.levels)
      .setRequired(true);
    itemMap[item.getId()] = q.id;

    if (q.id === CONFIG.CHANNEL_QID) {
      const ch = form.addCheckboxItem()
        .setTitle('前の質問は、どの販売チャネルについてお答えいただきましたか？（当てはまるものをすべて）')
        .setChoiceValues(CHANNEL_OPTIONS)
        .setRequired(true);
      itemMap[ch.getId()] = 'CHANNEL';
    }
  });

  form.addSectionHeaderItem().setTitle('最後に');
  const habitItem = form.addCheckboxItem()
    .setTitle('ふだん社内で、次のような言葉が出ることはありますか？（当てはまるものをすべて）')
    .setHelpText('当てはまるものがなければ、選ばずに送信してください。')
    .setChoiceValues(habits.map(h => h.text))
    .setRequired(false);
  itemMap[habitItem.getId()] = 'HABITS';

  // 回答の生データは診断マスタの新しいタブ（フォームの回答）にも残す
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();

  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('FORM_ID', form.getId());
  props.setProperty('ITEM_MAP', JSON.stringify(itemMap));
  getReportFolder_();

  ensureReportUrlHeader_();

  ui.alert('フォームを作成しました',
    '回答用URL（企業様に送るリンク）：\n' + form.getPublishedUrl() +
    '\n\n編集用URL：\n' + form.getEditUrl() +
    '\n\n診断シートの保存先：Driveの「' + CONFIG.FOLDER_NAME + '」フォルダ',
    ui.ButtonSet.OK);
}

function removeFormTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ---------------------------------------------------------------------------
// 回答の処理
// ---------------------------------------------------------------------------

function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let company = '';
  try {
    const parsed = parseResponse_(e.response);
    company = parsed.company;
    const row = writeAnswerRow_(parsed);
    SpreadsheetApp.flush();
    const report = buildReport_(row);
    notifyResult_(report);
  } catch (err) {
    postToSlack_('⚠️ 売上方程式 診断：' + (company || '回答') + ' の診断シート作成に失敗しました。\n' + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function parseResponse_(response) {
  const itemMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('ITEM_MAP') || '{}');
  const questions = readMaster_();
  const levelsById = {};
  questions.forEach(q => { levelsById[q.id] = q.levels; });
  const habits = readHabits_();

  const parsed = {
    date: response.getTimestamp(),
    company: '', product: '', position: '',
    scores: {}, habits: [], channels: [],
  };

  response.getItemResponses().forEach(ir => {
    const key = itemMap[ir.getItem().getId()];
    const answer = ir.getResponse();
    if (!key) return;
    if (key === 'COMPANY') parsed.company = String(answer).trim();
    else if (key === 'PRODUCT') parsed.product = String(answer).trim();
    else if (key === 'POSITION') parsed.position = answer;
    else if (key === 'CHANNEL') parsed.channels = answer;
    else if (key === 'HABITS') {
      parsed.habits = answer.map(text => {
        const h = habits.find(x => x.text === text);
        return h ? h.no : text;
      });
    } else if (key === CONFIG.ROLE_QID) {
      parsed.scores[key] = Math.min(4, (answer || []).length);
    } else {
      const idx = (levelsById[key] || []).indexOf(answer);
      if (idx < 0) throw new Error(key + ' の回答「' + answer + '」が設問マスタの選択肢と一致しません。');
      parsed.scores[key] = idx;
    }
  });

  // 任意回答の③-2を選ばなかった場合は0点
  if (parsed.scores[CONFIG.ROLE_QID] === undefined) parsed.scores[CONFIG.ROLE_QID] = 0;
  return parsed;
}

function writeAnswerRow_(parsed) {
  const ss = ss_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ANSWER);
  const headers = sheet.getRange(1, 1, 1, ANSWER_COL.CHANNEL).getValues()[0];

  const companies = sheet.getRange('B:B').getValues();
  let row = companies.length;
  while (row > 1 && companies[row - 1][0] === '') row--;
  row += 1;

  const values = new Array(ANSWER_COL.CHANNEL).fill('');
  values[ANSWER_COL.DATE - 1] = Utilities.formatDate(parsed.date, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');
  values[ANSWER_COL.COMPANY - 1] = parsed.company;
  values[ANSWER_COL.PRODUCT - 1] = parsed.product;
  values[ANSWER_COL.POSITION - 1] = parsed.position;
  for (let c = ANSWER_COL.FIRST_SCORE; c < ANSWER_COL.HABITS; c++) {
    const qid = headers[c - 1];
    if (parsed.scores[qid] === undefined) throw new Error('回答シートの見出し「' + qid + '」に対応する回答がありません。');
    values[c - 1] = parsed.scores[qid];
  }
  values[ANSWER_COL.HABITS - 1] = parsed.habits.join(', ');
  values[ANSWER_COL.CHANNEL - 1] = parsed.channels.join('、');

  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  ensureResultFormulas_(row);
  return row;
}

// 診断結果シートの数式が回答の行まで届いていなければ、2行目の数式を下にコピーする
function ensureResultFormulas_(row) {
  const sheet = ss_().getSheetByName(CONFIG.SHEET_RESULT);
  if (sheet.getRange(row, 1).getFormula()) return;
  if (sheet.getMaxRows() < row) sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
  const lastCol = sheet.getLastColumn();
  sheet.getRange(2, 1, 1, lastCol).copyTo(sheet.getRange(row, 1, 1, lastCol));
}

function ensureReportUrlHeader_() {
  const sheet = ss_().getSheetByName(CONFIG.SHEET_ANSWER);
  const cell = sheet.getRange(1, ANSWER_COL.REPORT_URL);
  if (!cell.getValue()) {
    cell.setValue('診断シートURL')
      .setFontWeight('bold')
      .setBackground(CONFIG.COLOR_HEADER)
      .setFontFamily(CONFIG.FONT);
    sheet.setColumnWidth(ANSWER_COL.REPORT_URL, 260);
  }
}

// ---------------------------------------------------------------------------
// 企業ごとの診断シート
// ---------------------------------------------------------------------------

function buildReport_(row) {
  const ss = ss_();
  const answerSheet = ss.getSheetByName(CONFIG.SHEET_ANSWER);
  const resultSheet = ss.getSheetByName(CONFIG.SHEET_RESULT);

  const answerHeaders = answerSheet.getRange(1, 1, 1, ANSWER_COL.CHANNEL).getValues()[0];
  const answer = answerSheet.getRange(row, 1, 1, ANSWER_COL.CHANNEL).getValues()[0];
  const resultHeaders = resultSheet.getRange(1, 1, 1, 11).getValues()[0];
  const result = resultSheet.getRange(row, 1, 1, 11).getValues()[0];

  const company = answer[ANSWER_COL.COMPANY - 1];
  const product = answer[ANSWER_COL.PRODUCT - 1];
  if (!company) throw new Error('回答シートの' + row + '行目に会社名がありません。');

  const questions = readMaster_();
  const mapping = readMapping_();

  const scoreByQid = {};
  for (let c = ANSWER_COL.FIRST_SCORE; c < ANSWER_COL.HABITS; c++) {
    scoreByQid[answerHeaders[c - 1]] = Number(answer[c - 1]);
  }

  const weakest = result[7];
  const detailRows = questions.map(q => {
    const s = scoreByQid[q.id];
    const recs = mapping
      .filter(m => m.qid === q.id && m.low <= s && s <= m.high)
      .map(m => m.name)
      .filter((name, i, arr) => arr.indexOf(name) === i);
    return [
      q.id,
      q.category,
      q.text,
      s,
      q.levels[s],
      s >= 4 ? '最高段階に到達' : q.levels[s + 1],
      recs.length ? recs.join('、') : (s >= 4 ? '―' : '該当施策なし'),
    ];
  });

  const dateText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
  const file = SpreadsheetApp.create('診断シート_' + company + '_' + product + '_' + dateText);
  DriveApp.getFileById(file.getId()).moveTo(getReportFolder_());

  const sh = file.getSheets()[0].setName('診断結果');
  const round = v => (typeof v === 'number' ? Math.round(v) : v);

  const top = [
    ['売上方程式 診断シート', ''],
    ['会社名', company],
    ['商品名', product],
    ['回答日', answer[ANSWER_COL.DATE - 1]],
    ['ご回答者の役割', answer[ANSWER_COL.POSITION - 1]],
    ['', ''],
    ['項目', '点数（100点満点）'],
    [resultHeaders[2], round(result[2])],
    [resultHeaders[3], round(result[3])],
    [resultHeaders[4], round(result[4])],
    [resultHeaders[5], round(result[5])],
    ['総合点（4項目の幾何平均）', round(result[6])],
    ['一番弱い項目', weakest],
    ['進め方の判定', result[8]],
    ['一番弱い項目を伸ばした場合の総合点', round(result[9]) + '点（' + (round(result[10]) >= 0 ? '+' : '') + round(result[10]) + '点）'],
  ];
  sh.getRange(1, 1, top.length, 2).setValues(top);

  const tableTop = top.length + 2;
  const tableHeader = ['設問ID', '項目', '設問', '現在の段階', '現在の状態', '次の目標（1つ上の段階）', 'おすすめの施策'];
  sh.getRange(tableTop, 1, 1, tableHeader.length).setValues([tableHeader]);
  sh.getRange(tableTop + 1, 1, detailRows.length, tableHeader.length).setValues(detailRows);

  const extraTop = tableTop + detailRows.length + 2;
  sh.getRange(extraTop, 1, 2, 2).setValues([
    ['当てはまった口癖', answer[ANSWER_COL.HABITS - 1] || 'なし'],
    ['①-4の対象チャネル', answer[ANSWER_COL.CHANNEL - 1] || ''],
  ]);

  formatReport_(sh, top.length, tableTop, detailRows, extraTop, weakest);
  addScoreChart_(sh);

  const url = file.getUrl();
  answerSheet.getRange(row, ANSWER_COL.REPORT_URL).setValue(url);

  return {
    row: row,
    company: company,
    product: product,
    total: round(result[6]),
    weakest: weakest,
    judgement: result[8],
    url: url,
  };
}

function formatReport_(sh, topRows, tableTop, detailRows, extraTop, weakest) {
  const lastRow = extraTop + 1;
  sh.getRange(1, 1, lastRow, 7)
    .setFontFamily(CONFIG.FONT)
    .setFontSize(10)
    .setVerticalAlignment('top')
    .setWrap(true);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1, 4, 1).setFontWeight('bold');
  sh.getRange(7, 1, 1, 2).setFontWeight('bold').setBackground(CONFIG.COLOR_HEADER);
  sh.getRange(8, 2, 5, 1).setNumberFormat('0"点"').setHorizontalAlignment('right');
  sh.getRange(12, 1, 1, 2).setFontWeight('bold');
  sh.getRange(8, 1, topRows - 7, 2).setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(tableTop, 1, 1, 7).setFontWeight('bold').setBackground(CONFIG.COLOR_HEADER);
  sh.getRange(tableTop + 1, 4, detailRows.length, 1).setHorizontalAlignment('center');
  sh.getRange(tableTop, 1, detailRows.length + 1, 7)
    .setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  detailRows.forEach((r, i) => {
    if (r[1] === weakest) sh.getRange(tableTop + 1 + i, 1, 1, 7).setBackground(CONFIG.COLOR_WEAK);
  });
  sh.getRange(extraTop, 1, 2, 1).setFontWeight('bold');

  [80, 150, 280, 80, 280, 280, 320].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(0);
}

function addScoreChart_(sh) {
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange(7, 1, 5, 2))
    .setNumHeaders(1)
    .setPosition(1, 4, 0, 0)
    .setOption('title', '4項目の点数（100点満点）')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { minValue: 0, maxValue: 100 })
    .setOption('colors', ['#E69138'])
    .setOption('fontName', CONFIG.FONT)
    .setOption('width', 520)
    .setOption('height', 260)
    .build();
  sh.insertChart(chart);
}

function getReportFolder_() {
  if (CONFIG.REPORT_FOLDER_ID) return DriveApp.getFolderById(CONFIG.REPORT_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REPORT_FOLDER_ID');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      // フォルダが削除された場合は作り直す
    }
  }
  const ssFile = DriveApp.getFileById(ss_().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('REPORT_FOLDER_ID', folder.getId());
  return folder;
}

// ---------------------------------------------------------------------------
// Slack通知
// ---------------------------------------------------------------------------

function notifyResult_(report) {
  const text = [
    '*' + report.company + '*（' + report.product + '）',
    '診断結果が出ました',
    '総合点 ' + report.total + '点／一番弱い項目：' + report.weakest,
    '診断シートはこちら ' + report.url,
  ].join('\n');
  return postToSlack_(text);
}

// Slackの通知先が未設定のときは何もせず false を返す（診断シートの作成は止めない）
function postToSlack_(text) {
  const url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!url) {
    console.log('Slackの通知先が未設定のため、通知を送りませんでした：\n' + text);
    return false;
  }
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Slackへの通知に失敗しました（' + res.getResponseCode() + '）：' + res.getContentText());
  }
  return true;
}

// ---------------------------------------------------------------------------
// マスタの読み込み
// ---------------------------------------------------------------------------

// フォーム送信のトリガーから呼ばれたときも、診断マスタを確実に開く
function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

function readMaster_() {
  const sheet = ss_().getSheetByName(CONFIG.SHEET_MASTER);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
    .filter(r => r[0])
    .map(r => ({
      id: r[0],
      category: r[1],
      text: r[2],
      levels: r.slice(4, 9).map(String),
    }));
}

function readHabits_() {
  const sheet = ss_().getSheetByName(CONFIG.SHEET_HABITS);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ no: r[0], text: r[1] }));
}

function readMapping_() {
  const sheet = ss_().getSheetByName(CONFIG.SHEET_MAP);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
    .filter(r => r[0] && r[6])
    .map(r => ({ qid: r[0], low: Number(r[2]), high: Number(r[3]), name: r[6] }));
}
