/**
 * FIT365 各店ブック専用 Apps Script（未納・入金・SMS 等＋各店→全店チェック同期）
 *
 * 全店ブック用は別ファイル: FIT365_全店実施管理ブック_GAS/Code.gs（全店→各店。双方向は両方にトリガー＋店舗ブック）
 * 各店に複製するのはこの Code.gs ＋ dialog.html
 */

/** シート見出し（比較時は normalizeFit365HeaderLabel で正規化して一致判定） */
var FIT365_SHEET_HEADER_LABELS = {
  STATUS: "2ヶ月未納 貸倒予定者",
  PREMIUM: "プレミアム 家族会員有無",
  MEMBER_ID: "会員番号",
  NAME: "名前",
  PHONE: "電話番号",
  EMAIL: "メールアドレス"
};

/** 見出しに無い場合の列（従来テンプレ向けフォールバック） */
var FIT365_FALLBACK_COL_PAID_CHECK = 22;

/** 月次未納シートのコピー元（テンプレート） */
var FIT365_GENPON_SHEET_NAME = "原本";

/** 時間トリガーで実行する関数名（毎月1日・UI なし） */
var FIT365_MONTHLY_GENPON_HANDLER = "fit365MonthlyCreateSheetFromGenpon";
/** 手動専用（時間トリガーに付けると getUi で失敗する） */
var FIT365_MONTHLY_GENPON_HANDLER_MANUAL = "fit365MonthlyCreateSheetFromGenponManual";

/** シート名パターン: 【店舗】26年4月（西暦下2桁・月は1〜2桁） */
var FIT365_MONTHLY_SHEET_NAME_RE = /^【([^】]+)】(\d{2})年(\d{1,2})月$/;

/**
 * 各店→全店（月次「【店名】YY年M月」）:
 * - 6行目 B〜E・G〜K → 全店 C〜K（チェック）
 * - 全店 2 行目 K〜N（半角の日＝DL）→ 各店 Q4〜Q7（一方通行・全店のみ編集）
 * - 各店 P4〜P7（実施した日を記入）→ 全店 その店行の K〜N にチェック（P4=1回目→K列）
 * 店名はタブ名（無いとき B12）。ループ防止: _FIT365_SYNC（全店側 B1 は DL 行エコー用・予備）
 */
/** 各店→全店 の onEdit 同期 */
var FIT365_STORE_HUB_SYNC_ENABLED = true;
/** 全店実施管理ブック ID（書き込み先） */
var FIT365_STORE_HUB_SPREADSHEET_ID = "1jyeP8hZYLICEuZEQktwwDUqGgaatuQok69gC3vd0KDc";
/** スクリプト連携用の非表示シート（自動作成） */
var FIT365_SYNC_CONTROL_SHEET_NAME = "_FIT365_SYNC";
/** エコー無視の有効ミリ秒 */
var FIT365_SYNC_ECHO_MS = 8000;
var FIT365_STORE_HUB_SOURCE_ROW = 6;
var FIT365_STORE_HUB_SOURCE_STORE_CELL = "B12";
var FIT365_STORE_HUB_TARGET_STORE_COL = 2;
var FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW = 4;
/** 空でなければ、集計ブック内のこのシート名を最優先で使う（例: "26年4月"） */
var FIT365_STORE_HUB_FORCE_SHEET_NAME = "";

/** SMS: 全店 2 行目 K〜N（DL）→各店 Q4:Q7。各店 P4:P7（実施日）→全店その店行の K〜N チェック */
var FIT365_SMS_HUB_SCHEDULE_ROW = 2;
var FIT365_SMS_HUB_COL_FIRST = 11;
var FIT365_SMS_HUB_COL_LAST = 14;
var FIT365_SMS_STORE_ROWS_FIRST = 4;
var FIT365_SMS_STORE_ROWS_LAST = 7;
var FIT365_SMS_STORE_COL_DATE = 16; // P列: 実施日
var FIT365_SMS_STORE_COL_DL = 17;   // Q列: DL（全店指示日）

function fit365SmsHubScheduleRangeA1_() {
  var r = FIT365_SMS_HUB_SCHEDULE_ROW;
  return "K" + r + ":N" + r;
}

function fit365SmsHubStoreCheckboxRangeA1_(hitRow) {
  return "K" + hitRow + ":N" + hitRow;
}

function fit365SmsStoreQBlockA1_() {
  return "Q" + FIT365_SMS_STORE_ROWS_FIRST + ":Q" + FIT365_SMS_STORE_ROWS_LAST;
}

function fit365SmsStorePBlockA1_() {
  return "P" + FIT365_SMS_STORE_ROWS_FIRST + ":P" + FIT365_SMS_STORE_ROWS_LAST;
}

/**
 * 店舗ブックのシート名から、集計ブック内の対応シートを取得する。
 * 試す順: FIT365_STORE_HUB_FORCE_SHEET_NAME（設定時）→ 完全一致 → 【店舗】を外した名前 → YY年M月 のみ（店舗シートが【】付きのとき）
 * 店舗シートが「26年4月」のみのときは、【店名】+ 同じ月名 を試す（storeNamePlain が必要）
 */
function fit365GetHubSheetForStoreTab_(hub, storeSheetName, storeNamePlain) {
  var forced = String(FIT365_STORE_HUB_FORCE_SHEET_NAME || "").trim();
  if (forced) {
    var sh0 = hub.getSheetByName(forced);
    if (sh0) return sh0;
  }
  var n = String(storeSheetName || "").trim();
  var plainStore = String(storeNamePlain || "").trim();
  var candidates = [];
  if (n) candidates.push(n);
  var stripped = n.replace(/^【[^】]+】\s*/, "").trim();
  if (stripped && candidates.indexOf(stripped) === -1) candidates.push(stripped);
  var m = n.match(FIT365_MONTHLY_SHEET_NAME_RE);
  if (m) {
    var period = m[2] + "年" + m[3] + "月";
    if (candidates.indexOf(period) === -1) candidates.push(period);
    if (plainStore) {
      var br = "【" + plainStore + "】" + period;
      if (candidates.indexOf(br) === -1) candidates.push(br);
    }
  } else if (plainStore && /^\d{2}年\d{1,2}月$/.test(n)) {
    if (candidates.indexOf("【" + plainStore + "】" + n) === -1) {
      candidates.push("【" + plainStore + "】" + n);
    }
  }
  for (var i = 0; i < candidates.length; i++) {
    var sh = hub.getSheetByName(candidates[i]);
    if (sh) return sh;
  }
  return null;
}

/** セル値を真偽に（チェック以外は null） */
function fit365CoerceCheckboxValue_(v) {
  if (v === true || v === false) return v;
  if (v === "" || v === null || v === undefined) return null;
  if (v === "TRUE" || v === "true" || v === "True") return true;
  if (v === "FALSE" || v === "false" || v === "False") return false;
  return null;
}

function fit365IsLikelyCheckboxEdit_(vRaw) {
  if (vRaw === "" || vRaw === null || vRaw === undefined) return true;
  if (fit365CoerceCheckboxValue_(vRaw) !== null) return true;
  return false;
}

function fit365SyncToast_(ss, title, msg) {
  try {
    ss.toast(String(msg).slice(0, 280), String(title || "FIT365各店"), 12);
  } catch (ignore) {
    Logger.log(title + ": " + msg);
  }
}

/** 全店ブック未共有などで出るが、別トリガーで同期済みのときは画面通知だけ抑える */
function fit365IsDocumentPermissionError_(errOrMsg) {
  var s =
    errOrMsg && errOrMsg.message ? String(errOrMsg.message) : String(errOrMsg || "");
  return (
    s.indexOf("アクセスする権限がありません") !== -1 ||
    s.indexOf("permission to access") !== -1 ||
    s.indexOf("Authorization is required") !== -1
  );
}

function fit365StoreHubNotifySyncIssue_(ss, title, errOrMsg) {
  var msg =
    errOrMsg && errOrMsg.message ? String(errOrMsg.message) : String(errOrMsg || "");
  Logger.log(String(title || "FIT365各店") + ": " + msg);
  if (fit365IsDocumentPermissionError_(msg)) return;
  fit365SyncToast_(ss, title, msg);
}

function fit365EnsureControlSheet_(ss) {
  var name = FIT365_SYNC_CONTROL_SHEET_NAME;
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.hideSheet();
  return sh;
}

function fit365MarkStoreEchoOnHub_(hubSS, hubRow) {
  fit365EnsureControlSheet_(hubSS);
  hubSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME).getRange(1, 1).setValue("STORE_ECHO|" + hubRow + "|" + Date.now());
}

/** 全店 _FIT365_SYNC の B1: 各店から DL 行を書いた直後（onEdit ループ防止・現状未使用） */
function fit365MarkStoreEchoOnHubRow2_(hubSS) {
  fit365EnsureControlSheet_(hubSS);
  hubSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME).getRange(1, 2).setValue("STORE_ECHO_ROW2|" + Date.now());
}

function fit365StoreShouldSkipEchoFromHub_(storeSS) {
  fit365EnsureControlSheet_(storeSS);
  var sh = storeSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME);
  var parts = String(sh.getRange(1, 1).getDisplayValue() || "").split("|");
  if (parts[0] !== "HUB_ECHO") return false;
  var ts = parseInt(parts[1], 10);
  if (!ts || Date.now() - ts > FIT365_SYNC_ECHO_MS) return false;
  sh.getRange(1, 1).clearContent();
  return true;
}

function fit365StoreSheetIsMonthlyTab_(sheet) {
  return FIT365_MONTHLY_SHEET_NAME_RE.test(String(sheet.getName() || ""));
}

function fit365StoreIsSmsSyncCell_(row, col) {
  return (
    row >= FIT365_SMS_STORE_ROWS_FIRST &&
    row <= FIT365_SMS_STORE_ROWS_LAST &&
    (col === FIT365_SMS_STORE_COL_DATE || col === FIT365_SMS_STORE_COL_DL)
  );
}

/** 1〜31 の日（半角数字）を返す。空は null。解釈不能も null */
function fit365ParseSmsDayNumber_(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "boolean") return null;
  var n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^\d.]/g, ""));
  if (isNaN(n)) return null;
  n = Math.floor(Math.abs(n));
  if (n < 1 || n > 31) return null;
  return n;
}

/** 全店 DL 行（K〜N）に書く値（数値または空） */
function fit365NormalizeSmsHubScheduleCell_(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (v instanceof Date && !isNaN(v.getTime())) {
    var dom = v.getDate();
    return dom >= 1 && dom <= 31 ? dom : "";
  }
  if (typeof v === "number" && !isNaN(v)) return Math.floor(Math.abs(v));
  var n = fit365ParseSmsDayNumber_(v);
  return n === null ? "" : n;
}

function fit365GetStoreNameFromSheet_(sheet) {
  var tab = sheet.getName();
  var m = tab.match(FIT365_MONTHLY_SHEET_NAME_RE);
  if (m && m[1]) return String(m[1]).trim();
  var cell = String(sheet.getRange(FIT365_STORE_HUB_SOURCE_STORE_CELL).getDisplayValue() || "").trim();
  return cell.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
}

function fit365OpAlertDone_() {
  SpreadsheetApp.getUi().alert("完了しました", SpreadsheetApp.getUi().ButtonSet.OK);
}

function fit365OpAlertError_() {
  SpreadsheetApp.getUi().alert("エラー", SpreadsheetApp.getUi().ButtonSet.OK);
}

function fit365AuthStep1Core_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  fit365EnsureControlSheet_(ss);
  SpreadsheetApp.openById(FIT365_STORE_HUB_SPREADSHEET_ID).getName();
}

function fit365AuthStep1() {
  try {
    fit365AuthStep1Core_();
    fit365OpAlertDone_();
  } catch (e1) {
    Logger.log(e1);
    fit365OpAlertError_();
  }
}

/** メニュー: 各店の「【店名】YY年M月」シートで 6行目→全店、月次なら SMS（DL 取込＋日付→チェック） */
function fit365SyncTestFromMenu_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var row6 = fit365StoreHubSyncRow6All_(sh);
  var sms = fit365StoreSheetIsMonthlyTab_(sh)
    ? fit365StoreSmsRefreshDlFromHub_(sh) + "\n" + fit365StoreHubSyncSmsToHub_(sh)
    : "SMS: 月次タブ以外のため未実行";
  ui.alert("同期テスト（各店→全店）", row6 + "\n" + sms, ui.ButtonSet.OK);
}

function fit365AuthStep2Core_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  fit365EnsureControlSheet_(ss);
  removeFit365MonthlyGenponTrigger();
  installFit365MonthlyGenponTriggerSilent_();
  installFit365StoreHubCheckboxSyncTrigger();
}

function fit365AuthStep2() {
  try {
    fit365AuthStep2Core_();
    fit365OpAlertDone_();
  } catch (e2) {
    Logger.log(e2);
    fit365OpAlertError_();
  }
}

/**
 * セル内の改行・全角スペースを揃えて見出し比較用にする（完全一致に近い運用）
 */
function normalizeFit365HeaderLabel(v) {
  return String(v || "")
    .replace(/^\uFEFF/, "")
    .replace(/\r\n|\n|\r/g, " ")
    .replace(/[ \u3000]+/g, " ")
    .trim();
}

/**
 * 1行ぶんのセル値から、正規化ラベル → 列番号(1始まり) のマップを作る（先勝ち）
 */
function buildFit365HeaderColumnMap(headerValues) {
  var map = {};
  for (var c = 0; c < headerValues.length; c++) {
    var key = normalizeFit365HeaderLabel(headerValues[c]);
    if (!key || map.hasOwnProperty(key)) continue;
    map[key] = c + 1;
  }
  return map;
}

/**
 * 必須見出しが揃っている行をヘッダー行とみなし、列番号とデータ開始行を返す
 */
function resolveFit365TemplateLayout(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 26);
  var scanLastRow = Math.min(sheet.getLastRow(), 25);

  var requiredNorm = {};
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.STATUS)] = true;
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.PREMIUM)] = true;
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.MEMBER_ID)] = true;
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.NAME)] = true;
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.PHONE)] = true;
  requiredNorm[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.EMAIL)] = true;

  for (var r = 1; r <= scanLastRow; r++) {
    var rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var colMap = buildFit365HeaderColumnMap(rowVals);
    var ok = true;
    for (var k in requiredNorm) {
      if (!requiredNorm.hasOwnProperty(k)) continue;
      if (!colMap.hasOwnProperty(k)) {
        ok = false;
        break;
      }
    }
    if (!ok) continue;

    var layout = {
      headerRow: r,
      dataStartRow: r + 1,
      colStatus: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.STATUS)],
      colPremium: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.PREMIUM)],
      colMemberId: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.MEMBER_ID)],
      colName: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.NAME)],
      colPhone: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.PHONE)],
      colEmail: colMap[normalizeFit365HeaderLabel(FIT365_SHEET_HEADER_LABELS.EMAIL)]
    };

    var accountHeaderCandidates = [
      "勘定項目",
      "セキュリティ",
      "セキュリティ費",
      "セキュリティー費",
      "セキュリティー"
    ];
    layout.colAccount = 0;
    for (var ah = 0; ah < accountHeaderCandidates.length; ah++) {
      var accountKey = normalizeFit365HeaderLabel(accountHeaderCandidates[ah]);
      if (colMap.hasOwnProperty(accountKey)) {
        layout.colAccount = colMap[accountKey];
        break;
      }
    }
    if (!layout.colAccount) {
      var rScan;
      for (rScan = Math.max(1, r - 2); rScan <= Math.min(scanLastRow, r + 4); rScan++) {
        var rowScan = sheet.getRange(rScan, 1, 1, lastCol).getDisplayValues()[0];
        var ci;
        for (ci = 0; ci < rowScan.length; ci++) {
          var lab = normalizeFit365HeaderLabel(rowScan[ci]);
          if (!lab) continue;
          if (lab.indexOf("セキュリティ") !== -1 || lab === "勘定項目") {
            layout.colAccount = ci + 1;
            break;
          }
        }
        if (layout.colAccount) break;
      }
    }
    layout.colForceExit = colMap.hasOwnProperty(normalizeFit365HeaderLabel("強制退会"))
      ? colMap[normalizeFit365HeaderLabel("強制退会")]
      : 0;
    layout.colAmount = findFit365OptionalAmountColumnFromHeaderRow(rowVals);

    layout.colPaidCheck = findFit365OptionalPaidCheckColumnFromHeaderRow(rowVals) || FIT365_FALLBACK_COL_PAID_CHECK;

    return layout;
  }

  throw new Error(
    "未納シートの見出し行が見つかりません。次のラベルが同一行にすべて必要です（改行の違いは吸収します）: " +
      FIT365_SHEET_HEADER_LABELS.STATUS + " / " +
      FIT365_SHEET_HEADER_LABELS.PREMIUM + " / " +
      FIT365_SHEET_HEADER_LABELS.MEMBER_ID + " / " +
      FIT365_SHEET_HEADER_LABELS.NAME + " / " +
      FIT365_SHEET_HEADER_LABELS.PHONE + " / " +
      FIT365_SHEET_HEADER_LABELS.EMAIL
  );
}

/**
 * 入金チェック列の見出し（任意）。無ければフォールバック列を使う
 */
function findFit365OptionalPaidCheckColumnFromHeaderRow(headerValues) {
  for (var c = 0; c < headerValues.length; c++) {
    var n = normalizeFit365HeaderLabel(headerValues[c]);
    if (n === "入金済" || n === "入金済み" || n.indexOf("入金済") !== -1) return c + 1;
  }
  return 0;
}

/**
 * 金額列の見出し（任意）。無ければ 0
 */
function findFit365OptionalAmountColumnFromHeaderRow(headerValues) {
  var candidates = {
    "金額": true,
    "未納金額": true,
    "請求額": true,
    "当月請求額": true,
    "請求額合計": true,
    "会費+オプション": true
  };
  for (var c = 0; c < headerValues.length; c++) {
    var n = normalizeFit365HeaderLabel(headerValues[c]);
    if (candidates[n]) return c + 1;
  }
  // 見出し揺れ対応: 「金額」を含む列を拾う（割引金額などは除外）
  for (var i = 0; i < headerValues.length; i++) {
    var label = normalizeFit365HeaderLabel(headerValues[i]);
    if (!label) continue;
    if (label.indexOf("割引") !== -1) continue;
    if (label.indexOf("金額") !== -1) return i + 1;
  }
  return 0;
}

/**
 * 列番号の配列から、dataStartRow〜lastRow をクリア。clearValidations false の列は値のみ消す
 */
function clearFit365DataColumns(sheet, dataStartRow, lastRow, colIndices, clearValidations) {
  if (lastRow < dataStartRow) return;
  if (clearValidations === undefined) clearValidations = true;
  var rowCount = lastRow - dataStartRow + 1;
  if (rowCount < 1) return;
  var uniq = {};
  for (var i = 0; i < colIndices.length; i++) {
    var col = colIndices[i];
    if (!col || col < 1) continue;
    uniq[String(col)] = col;
  }
  for (var key in uniq) {
    if (!uniq.hasOwnProperty(key)) continue;
    var c = uniq[key];
    var rng = sheet.getRange(dataStartRow, c, rowCount, 1);
    if (clearValidations) rng.clearDataValidations();
    rng.clearContent();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("FIT365 各店（請求・同期）")
    .addItem("初回セットアップ", "fit365InitialStoreSetup")
    .addSeparator()
    .addItem("📅 当月シートを手動作成", "fit365MonthlyCreateSheetFromGenponManual")
    .addItem("📩 入金データを更新する", "updateFit365Payments")
    .addItem("📥 SMSリストをCSV保存", "downloadSmsCsv")
    .addItem("📄 未納者データ自動生成 (2CSV)", "showImportDialog")
    .addToUi();
}

/**
 * ダイアログの表示
 */
function showImportDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ss.getSheets().map(function (s) {
    return s.getName();
  });
  var template = HtmlService.createTemplateFromFile('dialog');
  template.sheetNamesJson = JSON.stringify(sheetNames);
  template.activeSheetNameJson = JSON.stringify(ss.getActiveSheet().getName());
  var html = template.evaluate().setWidth(500).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'FIT365 未納データ自動生成');
}

/**
 * シート名一覧と現在開いている(アクティブな)シート名をHTMLに渡す関数
 */
function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    activeSheet: ss.getActiveSheet().getName(),
    allSheets: ss.getSheets().map(s => s.getName())
  };
}

/**
 * 西暦年の下2桁と月から「26年4月」形式のラベル
 */
function fit365FormatYearMonthLabel(fullYear, month1to12) {
  return String(fullYear).slice(-2) + "年" + month1to12 + "月";
}

/**
 * runDate の「暦上の当月」とその前月 { year, month }（月は 1〜12）
 */
function fit365CurrentMonthParts(runDate) {
  return {
    year: runDate.getFullYear(),
    month: runDate.getMonth() + 1
  };
}

function fit365PreviousMonthPartsFrom(runDate) {
  var y = runDate.getFullYear();
  var m0 = runDate.getMonth();
  var d = new Date(y, m0 - 1, 1);
  return { year: d.getFullYear(), month: d.getMonth() + 1 };
}

function fit365YyToFullYear(yy) {
  var n = parseInt(yy, 10);
  if (n >= 70) return 1900 + n;
  return 2000 + n;
}

/**
 * 全シート名から月次シートを解析し、{ bracket, year, month, fullYear } の配列
 */
function fit365ListParsedMonthlySheets(ss) {
  var sheets = ss.getSheets();
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var m = name.match(FIT365_MONTHLY_SHEET_NAME_RE);
    if (!m) continue;
    var yy = parseInt(m[2], 10);
    var mon = parseInt(m[3], 10);
    out.push({
      name: name,
      storeRaw: m[1],
      storeBracket: "【" + m[1] + "】",
      yy: yy,
      month: mon,
      fullYear: fit365YyToFullYear(yy)
    });
  }
  return out;
}

/**
 * 先月ラベル（YY年M月）に一致するシートから店舗の【…】を取得。無ければ最新の月次シートの店舗を使う
 */
function fit365ResolveStoreBracketForNewSheet(ss, runDate) {
  var prev = fit365PreviousMonthPartsFrom(runDate);
  var prevLabel = fit365FormatYearMonthLabel(prev.year, prev.month);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.slice(-prevLabel.length) !== prevLabel) continue;
    var head = n.slice(0, n.length - prevLabel.length);
    if (!/^【[^】]+】$/.test(head)) continue;
    return head;
  }
  var parsed = fit365ListParsedMonthlySheets(ss);
  if (!parsed.length) return null;
  parsed.sort(function (a, b) {
    if (a.fullYear !== b.fullYear) return b.fullYear - a.fullYear;
    return b.month - a.month;
  });
  return parsed[0].storeBracket;
}

/**
 * 「原本」を複製し、当月の【店舗】YY年M月シートを作成（既にあればスキップ）
 * @param {Date} [runDate] 省略時は new Date()（時間トリガー・手動共通）
 * @param {{ silent: boolean }} [options] silent true のとき UI なし（トリガー用）
 * @returns {string} 結果メッセージ
 */
function fit365RunMonthlyGenponCopy(runDate, options) {
  runDate = runDate || new Date();
  options = options || {};
  var silent = !!options.silent;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var genpon = ss.getSheetByName(FIT365_GENPON_SHEET_NAME);
  if (!genpon) {
    var msg = "コピー元シート「" + FIT365_GENPON_SHEET_NAME + "」が見つかりません。";
    if (!silent) throw new Error(msg);
    return msg;
  }

  var cur = fit365CurrentMonthParts(runDate);
  var targetLabel = fit365FormatYearMonthLabel(cur.year, cur.month);
  var storeBracket = fit365ResolveStoreBracketForNewSheet(ss, runDate);
  if (!storeBracket) {
    var prevEx = fit365PreviousMonthPartsFrom(runDate);
    var msg2 =
      "店舗名を決められません。先月分のシート名「【店舗】" +
      fit365FormatYearMonthLabel(prevEx.year, prevEx.month) +
      "」に一致するものを用意するか、【店舗】YY年M月 形式のシートを1つ以上置いてください。";
    if (!silent) throw new Error(msg2);
    return msg2;
  }

  var newName = storeBracket + targetLabel;
  var existingSheet = ss.getSheetByName(newName);
  if (existingSheet) {
    var wasHidden = existingSheet.isSheetHidden();
    if (wasHidden) {
      existingSheet.showSheet();
    }
    var skipMsg = wasHidden
      ? "シート「" + newName + "」は既にありました（非表示だったため表示しました）。"
      : "シート「" + newName + "」は既にあるため作成をスキップしました。";
    if (silent) Logger.log(skipMsg);
    return skipMsg;
  }

  var newSheet = genpon.copyTo(ss);
  newSheet.setName(newName);
  // コピー元「原本」が非表示だと複製も非表示になるため、必ず表示する
  if (newSheet.isSheetHidden()) {
    newSheet.showSheet();
  }

  var allSheets = ss.getSheets();
  var genponIdx = -1;
  for (var g = 0; g < allSheets.length; g++) {
    if (allSheets[g].getName() === FIT365_GENPON_SHEET_NAME) {
      genponIdx = g;
      break;
    }
  }
  if (genponIdx >= 0) {
    try {
      ss.setSheetPosition(newSheet, genponIdx + 2);
    } catch (ePos) {
      // 並び順は手動で調整可能
    }
  }

  var okMsg = "シート「" + newName + "」を「" + FIT365_GENPON_SHEET_NAME + "」から作成しました。";
  if (silent) Logger.log(okMsg);
  return okMsg;
}

/** メニューから手動実行 */
function fit365MonthlyCreateSheetFromGenponManual() {
  try {
    fit365RunMonthlyGenponCopy(new Date(), { silent: true });
    fit365OpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365OpAlertError_();
  }
}

/**
 * 月次の時間トリガーを削除（正規＋誤って付いた Manual 用の両方）
 */
function removeFit365MonthlyGenponTrigger() {
  var handlers = [FIT365_MONTHLY_GENPON_HANDLER, FIT365_MONTHLY_GENPON_HANDLER_MANUAL];
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (handlers.indexOf(triggers[i].getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 毎月1日 6:10 に原本から当月シートを作成する時間トリガー（UI なし関数のみ）
 * @returns {boolean} 新規作成したとき true
 */
function installFit365MonthlyGenponTriggerSilent_() {
  ScriptApp.newTrigger(FIT365_MONTHLY_GENPON_HANDLER)
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .nearMinute(10)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  return true;
}

/**
 * 毎月1日 午前6時（スクリプトのタイムゾーン）に原本から当月シートを作成
 */
function installFit365MonthlyGenponTrigger() {
  removeFit365MonthlyGenponTrigger();
  installFit365MonthlyGenponTriggerSilent_();
  fit365OpAlertDone_();
}

/**
 * 時間トリガーから呼ばれる（UI なし・ポップアップなし）
 */
function fit365MonthlyCreateSheetFromGenpon() {
  try {
    return fit365RunMonthlyGenponCopy(new Date(), { silent: true });
  } catch (e) {
    Logger.log("fit365MonthlyCreateSheetFromGenpon: " + e);
    throw e;
  }
}

/**
 * Code.gs 貼り直し後: 月次（時間）と onEdit を正しい状態に直す（各店8件などに実行）
 */
function fit365ReinstallStoreTriggers() {
  try {
    removeFit365MonthlyGenponTrigger();
    removeFit365StoreHubOnEditTriggers_();
    installFit365MonthlyGenponTriggerSilent_();
    installFit365StoreHubCheckboxSyncTrigger();
    fit365OpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365OpAlertError_();
  }
}

/**
 * 【メイン】2つのCSVデータを受け取り、既存のテンプレートシートの指定列に入力する処理
 */
function processFit365Data(csvTransferData, csvMemberData, targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheetName);
  
  if (!sheet) {
    throw new Error("指定されたシート「" + targetSheetName + "」が見つかりません。");
  }

  var targetLayout = resolveFit365TemplateLayout(sheet);

  // ==========================================
  // 【1】 先月シートの特定と「未入金者」のリストアップ
  // ==========================================
  var lastMonthSheetName = "";
  var match = targetSheetName.match(/^(.*?)(\d{2})年度(\d{1,2})月(.*?)$/);
  var unpaidLastMonthSet = new Set();
  
  if (match) {
    var prefix = match[1];
    var y = parseInt(match[2], 10);
    var m = parseInt(match[3], 10);
    var suffix = match[4];
    m -= 1;
    if (m === 0) { y -= 1; m = 12; }
    lastMonthSheetName = prefix + y + "年度" + m + "月" + suffix;
  }

  if (lastMonthSheetName) {
    var lmSheet = ss.getSheetByName(lastMonthSheetName);
    if (lmSheet) {
      var lmLayout = null;
      try {
        lmLayout = resolveFit365TemplateLayout(lmSheet);
      } catch (eLm) {
        lmLayout = { dataStartRow: 13, colMemberId: 7, colPaidCheck: 22 };
      }
      var lmData = lmSheet.getDataRange().getValues();
      var lmIdIdx = lmLayout.colMemberId - 1;
      var lmPaidIdx = lmLayout.colPaidCheck - 1;
      var lmStart0 = lmLayout.dataStartRow - 1;
      for (var r = lmStart0; r < lmData.length; r++) {
        var bVal = String(lmData[r][lmIdIdx] || "").trim();
        if (bVal.startsWith("'")) bVal = bVal.substring(1);
        var vVal = lmData[r][lmPaidIdx];

        if (bVal && vVal !== true) {
          unpaidLastMonthSet.add(bVal);
        }
      }
    }
  }

  // ==========================================
  // 【2】 電話番号とメールアドレスの抽出 (条件別会員リストから)
  // ==========================================
  var memberInfoMap = new Map();
  
  if (csvMemberData && csvMemberData.length > 1) {
    for (var i = 1; i < csvMemberData.length; i++) {
      var row = csvMemberData[i];
      if (row.length > 9) {
        var mId = String(row[0]).replace(/^\uFEFF/, "").trim(); 
        mId = mId.replace(/[０-９]/g, function(s) { 
          return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); 
        });

        var rawPhone = String(row[8] || "").trim(); // I列: TEL
        var email = String(row[9] || "").trim();    // J列: EMAIL
        
        var cleanPhone = rawPhone.replace(/\D/g, ""); 
        if (cleanPhone.length >= 9) {
          if (cleanPhone.length === 10 && (cleanPhone.startsWith("9") || cleanPhone.startsWith("8") || cleanPhone.startsWith("7"))) cleanPhone = "0" + cleanPhone;
          else if (cleanPhone.length === 9 && !cleanPhone.startsWith("0")) cleanPhone = "0" + cleanPhone;
        } else {
          cleanPhone = "";
        }
        
        if (mId) {
          memberInfoMap.set(mId, { phone: cleanPhone, email: email });
        }
      }
    }
  }

  // ==========================================
  // 【2.5】 過去の同店舗シートからの連絡先（電話・メール）の補完
  // ==========================================
  var historicalInfoMap = new Map();
  var storePrefixMatch = targetSheetName.match(/^【.*?】/);
  
  if (storePrefixMatch) {
    var storePrefix = storePrefixMatch[0]; // 例: "【志木】"
    var allSheets = ss.getSheets();
    
    for (var sIdx = 0; sIdx < allSheets.length; sIdx++) {
      var historySheet = allSheets[sIdx];
      var historySheetName = historySheet.getName();
      
      // 同じ店舗名が含まれていて、現在処理中のシート以外のものを対象にする
      if (historySheetName.includes(storePrefix) && historySheetName !== targetSheetName) {
        var hLayout = null;
        try {
          hLayout = resolveFit365TemplateLayout(historySheet);
        } catch (eH) {
          hLayout = { dataStartRow: 13, colMemberId: 7, colPhone: 9, colEmail: 10 };
        }
        var hData = historySheet.getDataRange().getValues();
        var hStart0 = hLayout.dataStartRow - 1;
        var gi = hLayout.colMemberId - 1;
        var pi = hLayout.colPhone - 1;
        var ei = hLayout.colEmail - 1;

        for (var r = hStart0; r < hData.length; r++) {
          var maxIdx = Math.max(gi, pi, ei);
          if (hData[r].length <= maxIdx) continue;

          var hId = String(hData[r][gi] || "").replace(/^\uFEFF/, "").trim();
          if (hId.startsWith("'")) hId = hId.substring(1);
          if (!hId) continue;

          var hPhone = String(hData[r][pi] || "").trim();
          if (hPhone.startsWith("'")) hPhone = hPhone.substring(1);
          var hEmail = String(hData[r][ei] || "").trim();
          
          if (hPhone || hEmail) {
            if (!historicalInfoMap.has(hId)) {
              historicalInfoMap.set(hId, { phone: hPhone, email: hEmail });
            } else {
              // 既に登録されていても、空欄があれば埋める
              var existing = historicalInfoMap.get(hId);
              if (!existing.phone && hPhone) existing.phone = hPhone;
              if (!existing.email && hEmail) existing.email = hEmail;
            }
          }
        }
      }
    }
  }

  // ==========================================
  // 【3】 振替結果一覧表の解析と名寄せ (列の位置関係は維持)
  // 契約名列: 1行目ヘッダーから「契約名」を検出（無い場合は契約名の判定をスキップ）
  // ==========================================
  function findHeaderColumnIndex(headerRow, targetName) {
    if (!headerRow || headerRow.length === 0) return -1;
    for (var c = 0; c < headerRow.length; c++) {
      var h = String(headerRow[c] || "").replace(/^\uFEFF/, "").trim();
      if (h === targetName) return c;
    }
    return -1;
  }

  function premiumLabelFromContractName(contractText) {
    var t = String(contractText || "").trim();
    if (!t) return "";
    if (t.indexOf("グランプレミアム") !== -1) return "グランプレミアム";
    if (t.indexOf("プレミアム") !== -1) return "プレミアム";
    return "";
  }

  function mergePremiumLabel(existing, incoming) {
    if (incoming === "グランプレミアム") return "グランプレミアム";
    if (incoming === "プレミアム" && existing !== "グランプレミアム") return "プレミアム";
    return existing || "";
  }

  var contractNameCol = -1;
  if (csvTransferData.length > 0) {
    contractNameCol = findHeaderColumnIndex(csvTransferData[0], "契約名");
  }

  var members = new Map();
  function fit365MemberAggregateKey_(memberId, memberName) {
    var id = String(memberId || "").trim();
    if (id) return "ID:" + id;
    var nm = String(memberName || "").trim();
    if (nm) return "NAME:" + nm;
    return "";
  }
  function fit365ParseAmountSafe_(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return isNaN(v) ? 0 : Math.trunc(v);
    var s = String(v)
      .replace(/[,\s\u3000]/g, "")
      .replace(/[円¥]/g, "")
      .replace(/^'+/, "")
      .trim();
    if (!s) return 0;
    var n = Number(s);
    return isNaN(n) ? 0 : Math.trunc(n);
  }

  for (var i = 1; i < csvTransferData.length; i++) {
    var row = csvTransferData[i];
    if (row.length < 16) continue; 
    
    var memberId = String(row[3]).replace(/^\uFEFF/, "").trim();   // D列: 会員番号
    memberId = memberId.replace(/[０-９]/g, function(s) { 
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); 
    });

    var memberName = String(row[4]).trim(); // E列: 会員名

    var rowPremium = "";
    if (contractNameCol >= 0 && row.length > contractNameCol) {
      rowPremium = premiumLabelFromContractName(row[contractNameCol]);
    }
    
    var carryOver = fit365ParseAmountSafe_(row[9]);     // J列: 繰越金
    if (carryOver < 0) carryOver = 0;              
    
    var monthlyFee = fit365ParseAmountSafe_(row[10]);   // K列: 当月月会費
    var annualFee = fit365ParseAmountSafe_(row[11]);    // L列: 年会費
    var totalBilled = fit365ParseAmountSafe_(row[14]);  // O列: 請求額合計
    
    var rawStatus = String(row[15] || "").trim();  // P列: 状態
    var isWithdrawn = rawStatus.includes("退会");
    
    var memberKey = fit365MemberAggregateKey_(memberId, memberName);
    if (!memberKey) continue;
    if (!members.has(memberKey)) {
      members.set(memberKey, {
        memberId: memberId,
        name: memberName,
        carryOverTotal: carryOver,       
        monthlyFeeTotal: monthlyFee,     
        annualFeeTotal: annualFee,       
        billedTotal: totalBilled,        
        hasSecurityFee: annualFee > 0,   
        isWithdrawn: isWithdrawn,
        premiumLabel: rowPremium
      });
    } else {
      var m = members.get(memberKey);
      m.carryOverTotal += carryOver;
      m.monthlyFeeTotal += monthlyFee;
      m.annualFeeTotal += annualFee;
      m.billedTotal += totalBilled;
      if (!m.memberId && memberId) m.memberId = memberId;
      
      if (annualFee > 0) m.hasSecurityFee = true;
      if (isWithdrawn) m.isWithdrawn = true;
      m.premiumLabel = mergePremiumLabel(m.premiumLabel || "", rowPremium);
    }
  }

  // ==========================================
  // 【4】 出力データの生成
  // ==========================================
  var outputRows = [];
  
  members.forEach(function(data) {
    var status = "1カ月未納"; // デフォルト
    
    var V = data.carryOverTotal;   // V列相当: 繰越金合計
    var W = data.monthlyFeeTotal;  // W列相当: 当月月会費合計
    var X = data.annualFeeTotal;   // X列相当: 年会費合計
    var Y = data.billedTotal;      // Y列相当: 請求額合計
    
    // スプレッドシートのIF関数ロジックをそのまま適用
    if ((W + X) === Y) {
      status = "1カ月未納";
    } else if (V !== 0 && W === 0 && X === 0 && V === Y) {
      status = "貸倒予定者";
    } else if (W !== 0 && W !== Y) {
      status = "2ヶ月未納";
    }
    
    // D列用の勘定項目判定
    var itemType = data.hasSecurityFee ? "セキュリティ費" : ""; 
    
    // 連絡先情報（まずはCSVから取得）
    var memberId = String(data.memberId || "").trim();
    var info = memberInfoMap.get(memberId) || { phone: "", email: "" };
    
    // 【💡追加】CSVから連絡先が見つからない、または空欄の場合、過去シートのデータで補完する
    if (!info.phone || !info.email) {
      var histInfo = historicalInfoMap.get(memberId);
      if (histInfo) {
        if (!info.phone && histInfo.phone) info.phone = histInfo.phone;
        if (!info.email && histInfo.email) info.email = histInfo.email;
      }
    }

    var phoneStr = info.phone ? "'" + info.phone : "";
    var mIdStr = memberId ? "'" + memberId : "";
    var premiumLabel = data.premiumLabel || "";
    
    // A列(0) 〜 K列(10) までの11要素の配列を作成（任意の金額列にも対応）
    var rowData = new Array(11).fill("");
    
    rowData[3] = itemType;   // D列: 勘定項目
    rowData[5] = status;     // F列: ステータス
    // G列: 契約名がプレミアム／グランプレミアムのときはその文言。それ以外は従来どおり会員番号（C列はプレミアム系のときのみ会員番号）
    if (premiumLabel) {
      rowData[2] = mIdStr;     // C列: 会員番号（Gを契約表示に使うため）
      rowData[6] = premiumLabel; // G列: プレミアム / グランプレミアム
    } else {
      rowData[6] = mIdStr;     // G列: 会員番号
    }
    rowData[7] = data.name;  // H列: 氏名
    rowData[8] = phoneStr;   // I列: 電話番号
    rowData[9] = info.email; // J列: メールアドレス
    rowData[10] = data.billedTotal; // K列想定: 金額（会費+オプションの当月請求額合計）
    
    outputRows.push(rowData);
  });
  
  // 並び替え (貸倒予定者 -> 2ヶ月未納 -> 1カ月未納)
  outputRows.sort(function(a, b) {
    var rankA = a[5] === "貸倒予定者" ? 1 : (a[5] === "2ヶ月未納" ? 2 : 3);
    var rankB = b[5] === "貸倒予定者" ? 1 : (b[5] === "2ヶ月未納" ? 2 : 3);
    return rankA - rankB;
  });
  
  // ==========================================
  // 【5】 見出し行の直下から、見出しラベルで解決した列にのみ入力
  // ==========================================
  var layout = targetLayout;
  var startRow = layout.dataStartRow;
  var lastRow = Math.max(sheet.getLastRow(), startRow);

  var colsContentOnly = [
    layout.colPremium,
    layout.colMemberId,
    layout.colName,
    layout.colPhone,
    layout.colEmail
  ];
  if (layout.colAmount) colsContentOnly.push(layout.colAmount);
  if (layout.colAccount) colsContentOnly.push(layout.colAccount);
  clearFit365DataColumns(sheet, startRow, lastRow, colsContentOnly, false);

  var colsResetValidation = [layout.colStatus];
  if (layout.colForceExit) colsResetValidation.push(layout.colForceExit);
  clearFit365DataColumns(sheet, startRow, lastRow, colsResetValidation, true);

  function fit365MemberIdForSheet(rd) {
    if (rd[2]) return rd[2];
    var six = String(rd[6] || "");
    if (six === "プレミアム" || six === "グランプレミアム") return "";
    return rd[6];
  }

  function fit365PremiumForSheet(rd) {
    var six = String(rd[6] || "");
    if (six === "プレミアム" || six === "グランプレミアム") return six;
    return "";
  }

  function fit365ColumnVector(rows, getter) {
    var v = [];
    for (var i = 0; i < rows.length; i++) v.push([getter(rows[i])]);
    return v;
  }

  if (outputRows.length > 0) {
    var endRow = startRow + outputRows.length - 1;
    var outputRowCount = outputRows.length;
    if (sheet.getMaxRows() < endRow) {
      sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
    }

    sheet
      .getRange(startRow, layout.colStatus, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, function (rd) {
        return rd[5];
      }));
    sheet
      .getRange(startRow, layout.colPremium, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, fit365PremiumForSheet));
    sheet
      .getRange(startRow, layout.colMemberId, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, fit365MemberIdForSheet));
    sheet
      .getRange(startRow, layout.colName, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, function (rd) {
        return rd[7];
      }));
    sheet
      .getRange(startRow, layout.colPhone, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, function (rd) {
        return rd[8];
      }));
    sheet
      .getRange(startRow, layout.colEmail, outputRowCount, 1)
      .setValues(fit365ColumnVector(outputRows, function (rd) {
        return rd[9];
      }));
    if (layout.colAmount) {
      sheet
        .getRange(startRow, layout.colAmount, outputRowCount, 1)
        .setValues(
          fit365ColumnVector(outputRows, function (rd) {
            return rd[10];
          })
        );
    }

    if (layout.colAccount) {
      sheet
        .getRange(startRow, layout.colAccount, outputRowCount, 1)
        .setValues(
          fit365ColumnVector(outputRows, function (rd) {
            return rd[3];
          })
        );
    }

    var ruleF = SpreadsheetApp.newDataValidation()
      .requireValueInList(["貸倒予定者", "2ヶ月未納", "1カ月未納"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, layout.colStatus, outputRowCount, 1).setDataValidation(ruleF);

    if (layout.colForceExit) {
      var ruleE = SpreadsheetApp.newDataValidation()
        .requireValueInList(["強制退会"], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(startRow, layout.colForceExit, outputRowCount, 1).setDataValidation(ruleE);
    }
  }

  return outputRows.length;
}

/**
 * =========================================================================
 * 📥 SMSリストをCSV保存 (FIT365仕様)
 * =========================================================================
 */
function downloadSmsCsv() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Meiryo', sans-serif; padding: 20px; color: #333; }
          h3 { margin-top: 0; color: #e91e63; font-size: 16px; }
          p { font-size: 13px; color: #666; }
          select { padding: 8px; width: 100%; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; margin-bottom: 15px; }
          button { padding: 10px 20px; background-color: #e91e63; color: white; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 14px; font-weight: bold; }
          button:hover { background-color: #c2185b; }
          #msg { margin-top: 15px; font-size: 13px; text-align: center; }
        </style>
      </head>
      <body>
        <h3>SMSリストの出力</h3>
        <p>出力する対象グループを選択してください。</p>
        <select id="groupSelect">
          <option value="貸倒予定者">貸倒予定者 (退会後未納)</option>
          <option value="2ヶ月未納">2ヶ月未納</option>
          <option value="1カ月未納">1カ月未納</option>
        </select>
        <button onclick="requestDownload()">CSVを作成してダウンロード</button>
        <div id="msg"></div>
        
        <script>
          function requestDownload() {
            const group = document.getElementById('groupSelect').value;
            document.getElementById('msg').innerText = "データを作成中...";
            document.getElementById('msg').style.color = "#f4b400";
            
            if (typeof google !== 'undefined' && google.script && google.script.run) {
              google.script.run
                .withSuccessHandler(function(csvContent) {
                  if(csvContent === "") {
                    document.getElementById('msg').innerText = "未入金の対象者（電話番号あり）が見つかりません。";
                    document.getElementById('msg').style.color = "#d93025";
                    return;
                  }
                  const fileName = "sms_list_" + group + ".csv";
                  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                  const link = document.createElement("a");
                  const url = URL.createObjectURL(blob);
                  link.setAttribute("href", url);
                  link.setAttribute("download", fileName);
                  link.style.visibility = 'hidden';
                  document.body.appendChild(link);
                  link.click();
                  document.body.removeChild(link);
                  
                  document.getElementById('msg').innerText = "✅ ダウンロードを開始しました。";
                  document.getElementById('msg').style.color = "#0f9d58";
                  setTimeout(google.script.host.close, 1500);
                })
                .withFailureHandler(function(error) {
                  document.getElementById('msg').innerText = "エラー: " + error.message;
                  document.getElementById('msg').style.color = "#d93025";
                })
                .generateSmsCsvContent(group);
            } else {
              setTimeout(function() {
                document.getElementById('msg').innerText = "✅ ダウンロードを開始しました (プレビュー環境)。";
                document.getElementById('msg').style.color = "#0f9d58";
              }, 1500);
            }
          }
        </script>
      </body>
    </html>
  `).setWidth(350).setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SMS送信対象の選択');
}

/**
 * 選択されたグループのCSVコンテンツを生成する（バックエンド処理）
 */
function generateSmsCsvContent(selectedGroup) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getActiveSheet();

  let layout;
  try {
    layout = resolveFit365TemplateLayout(targetSheet);
  } catch (e) {
    throw new Error("SMS用の列が特定できません: " + e.message);
  }

  const dataSheetName = 'App入金_Data';
  const dataSheet = ss.getSheetByName(dataSheetName);
  const paidNames = new Set();
  
  if (dataSheet && dataSheet.getLastRow() > 1) {
    const values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues();
    values.forEach(row => {
      const d = row[0];
      const name = row[2];
      if (d instanceof Date) {
        const y = String(d.getFullYear()).slice(-2);
        const m = String(d.getMonth() + 1);
        const sheetName = targetSheet.getName();
        if (sheetName.includes(y) && sheetName.includes(m)) {
          paidNames.add(String(name).trim());
        }
      }
    });
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < layout.dataStartRow) return "";

  const c1 = Math.min(layout.colStatus, layout.colName, layout.colPhone);
  const c2 = Math.max(layout.colStatus, layout.colName, layout.colPhone);
  const numRows = lastRow - layout.dataStartRow + 1;
  const data = targetSheet.getRange(layout.dataStartRow, c1, numRows, c2 - c1 + 1).getValues();
  const offStatus = layout.colStatus - c1;
  const offName = layout.colName - c1;
  const offPhone = layout.colPhone - c1;
  let csvContent = "";

  for (let i = 0; i < data.length; i++) {
    const currentGroup = String(data[i][offStatus] || "").trim();
    const name = String(data[i][offName] || "").trim();
    let phone = String(data[i][offPhone] || "").trim();
    
    if (phone.startsWith("'")) phone = phone.substring(1);
    if (!name || !phone) continue;
    
    if (currentGroup === selectedGroup && !paidNames.has(name)) {
      csvContent += `"${phone}"\r\n`;
    }
  }

  return csvContent;
}

/**
 * =========================================================================
 * 📩 メールの自動集計＆月別シートへの消し込み機能 (FIT365 未納通知のみ)
 * =========================================================================
 */
function updateFit365Payments() {
  const UI_SHEET_NAME = 'App入金';        
  const DATA_SHEET_NAME = 'App入金_Data'; 
  const LABEL_NAME = 'App入金';           
  const TAX_RATE = 1.1;                   
  
  const COLOR_BG_HEADER = '#e91e63'; 
  const COLOR_TEXT_HEADER = '#ffffff'; 
  const COLOR_BG_TOTAL = '#f5f5f5';  

  const QUERY = 'from:info@fit365.jp subject:"未納金のお支払いについて" newer_than:6m';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const headers = [['対象年月', '受信日時', '氏名', '内訳項目', '金額(税込)', '金額(税抜)', 'Message-ID']];

  if (!dataSheet) {
    dataSheet = ss.insertSheet(DATA_SHEET_NAME);
    dataSheet.getRange(1, 1, 1, 7).setValues(headers);
    dataSheet.hideSheet(); 
  } else {
    dataSheet.clear(); 
    dataSheet.getRange(1, 1, 1, 7).setValues(headers);
  }

  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(LABEL_NAME);
  }

  const threads = GmailApp.search(QUERY, 0, 200); 
  const newRecords = [];

  if (threads.length > 0) {
    for (const thread of threads) {
      const messages = thread.getMessages();
      let hasFit365Mail = false;

      for (const message of messages) {
        if (message.getFrom().indexOf('fit365.jp') === -1) continue;
        hasFit365Mail = true;
        const messageId = message.getId();
        const body = message.getPlainBody();
        const date = message.getDate();
        
        const records = parseFit365Email(body, date, messageId, TAX_RATE);
        if (records.length > 0) {
          records.forEach(r => newRecords.push(r));
        }
      }
      if (hasFit365Mail) thread.addLabel(label);
    }
  }

  if (newRecords.length > 0) {
    newRecords.sort((a, b) => a[1] - b[1]); 
    const nextRow = 2;
    const lastRow = nextRow + newRecords.length - 1;
    dataSheet.getRange(nextRow, 1, lastRow, 7).setValues(newRecords);
    dataSheet.getRange("A:A").setNumberFormat("yyyy/mm/dd"); 
    dataSheet.getRange("B:B").setNumberFormat("yyyy/mm/dd HH:mm");
  }

  setupUISheetV5(ss, UI_SHEET_NAME, DATA_SHEET_NAME, COLOR_BG_HEADER, COLOR_TEXT_HEADER, COLOR_BG_TOTAL);

  let msg = `✅ 更新完了\n\nデータを最新の状態に再構築しました。\n\n・取得件数: ${newRecords.length}件`;
  Browser.msgBox(msg);
}

function setupUISheetV5(ss, uiSheetName, dataSheetName, headerBg, headerText, totalBg) {
  let uiSheet = ss.getSheetByName(uiSheetName);
  if (!uiSheet) uiSheet = ss.insertSheet(uiSheetName);
  uiSheet.clear();

  uiSheet.getRange("1:1").setBackground(totalBg).setFontColor("#000000"); 
  uiSheet.getRange("A1").setBackground("#ffffff").setFontWeight("bold").setBorder(true, true, true, true, true, true, headerBg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); 
  uiSheet.getRange("B1").setValue("【 表示分合計 】").setFontWeight("bold").setHorizontalAlignment("right");
  uiSheet.getRange("E1").setFormula("=SUBTOTAL(9, E3:E)"); 
  uiSheet.getRange("F1").setFormula("=SUBTOTAL(9, F3:F)");
  uiSheet.getRange("E1").setNumberFormat('"税込" #,##0').setFontWeight("bold").setFontColor(headerBg).setFontSize(11);
  uiSheet.getRange("F1").setNumberFormat('"税抜" #,##0').setFontWeight("bold").setFontColor("#555555").setFontSize(11);

  const headers = [['対象年月', '受信日時', '氏名', '内訳項目', '金額(税込)', '金額(税抜)']];
  uiSheet.getRange(2, 1, 2, 6).setValues(headers)
    .setFontWeight('bold').setBackground(headerBg).setFontColor(headerText) 
    .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  uiSheet.setFrozenRows(2);

  const dataSheet = ss.getSheetByName(dataSheetName);
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    const dates = dataSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const dateSet = new Set();
    const dateList = [];
    dates.forEach(d => {
      if (d instanceof Date) {
        const formatted = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy/MM");
        if (!dateSet.has(formatted)) {
          dateSet.add(formatted);
          dateList.push(formatted);
        }
      }
    });
    dateList.sort().reverse();

    if (dateList.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(dateList, true).setAllowInvalid(false).build();
      uiSheet.getRange("A1").setDataValidation(rule).setNumberFormat("@"); 
      if (uiSheet.getRange("A1").getValue() === "") uiSheet.getRange("A1").setValue(dateList[0]);
    }
  }

  const queryFormula = 
    `=QUERY('${dataSheetName}'!A2:F, "SELECT A, B, C, D, E, F WHERE A IS NOT NULL " & ` + 
    `IF(A1="", "", " AND YEAR(A) = " & VALUE(LEFT(A1, 4)) & " AND MONTH(A) = " & (VALUE(RIGHT(A1, 2)) - 1)) & ` + 
    `" ORDER BY A DESC, B DESC", 0)`;
  uiSheet.getRange("A3").setFormula(queryFormula);
  
  uiSheet.getRange("A3:A").setNumberFormat("yyyy/mm");
  uiSheet.getRange("B3:B").setNumberFormat("yyyy/mm/dd HH:mm");
  uiSheet.getRange("E3:F").setNumberFormat("#,##0");
  uiSheet.setColumnWidth(3, 150); 
  uiSheet.setColumnWidth(4, 350); 
}

function parseFit365Email(body, receiveDate, messageId, taxRate) {
  const results = [];
  const nameMatch = body.match(/^(.*?)[\s　]*様/m);
  const name = nameMatch ? nameMatch[1].trim() : "不明";

  const detailRegex = /([^\n\r]*?\(\d{4}年\d{1,2}月\).*?)[\s　]*税込([\d,]+)円/g;
  let match;
  let found = false;
  const groupedData = new Map();

  while ((match = detailRegex.exec(body)) !== null) {
    found = true;
    const fullItemName = match[1].trim(); 
    const amount = parseInt(match[2].replace(/,/g, ''), 10);
    
    let targetDateObj;
    const dateMatch = fullItemName.match(/\((\d{4}年\d{1,2}月)\)/);
    if (dateMatch) {
        targetDateObj = new Date(dateMatch[1].replace("年", "/").replace("月", "/1"));
    } else {
        targetDateObj = new Date(receiveDate);
        targetDateObj.setDate(1);
    }
    targetDateObj.setHours(0, 0, 0, 0);

    const key = targetDateObj.getTime();
    if (!groupedData.has(key)) {
      groupedData.set(key, { dateObj: targetDateObj, items: [], totalAmount: 0 });
    }
    const data = groupedData.get(key);
    data.items.push(fullItemName);
    data.totalAmount += amount;
  }

  if (found) {
    groupedData.forEach((data) => {
      results.push([
        data.dateObj, receiveDate, name, data.items.join('、'),
        data.totalAmount, Math.round(data.totalAmount / taxRate), messageId
      ]);
    });
  }

  if (!found) {
    const totalMatch = body.match(/合計.*?税込([\d,]+)円/);
    if (totalMatch) {
      const amount = parseInt(totalMatch[1].replace(/,/g, ''), 10);
      let targetDateObj;
      const dateMatchG = body.match(/\((\d{4}年\d{1,2}月)\)/);
      if (dateMatchG) targetDateObj = new Date(dateMatchG[1].replace("年", "/").replace("月", "/1"));
      else { targetDateObj = new Date(receiveDate); targetDateObj.setDate(1); }
      targetDateObj.setHours(0, 0, 0, 0);

      results.push([targetDateObj, receiveDate, name, "明細不明（合計のみ抽出）", amount, Math.round(amount / taxRate), messageId]);
    }
  }
  return results;
}

/**
 * アクティブシートの B12 等から店舗名を取り、集計ブックの対象シート・行を解決する。
 * @returns {{ hubSheet: GoogleAppsScript.Spreadsheet.Sheet, hitRow: number, storeName: string }|{ error: string }}
 */
function fit365StoreHubResolveContext_(sheet) {
  var storeName = fit365GetStoreNameFromSheet_(sheet);
  if (!storeName) {
    return {
      error:
        "店舗名を取得できません。シート名を「【店舗名】YY年M月」にするか、セル " +
        FIT365_STORE_HUB_SOURCE_STORE_CELL +
        " に店名を入力してください。"
    };
  }
  var hub = SpreadsheetApp.openById(FIT365_STORE_HUB_SPREADSHEET_ID);
  var hubSheet = fit365GetHubSheetForStoreTab_(hub, sheet.getName(), storeName);
  if (!hubSheet) {
    return {
      error:
        "集計ブックでシートが見つかりません。店舗タブ「" +
        sheet.getName() +
        "」に対応する「YY年M月」等があるか、または Code.gs の FIT365_STORE_HUB_FORCE_SHEET_NAME を設定してください。"
    };
  }
  var hitRow = fit365FindHubRowByStoreName_(hubSheet, storeName);
  if (hitRow < 1) {
    return { error: "集計シート「" + hubSheet.getName() + "」のB列に店舗名「" + storeName + "」がありません。" };
  }
  return { hubSheet: hubSheet, hitRow: hitRow, storeName: storeName };
}

/**
 * 各店 P4:P7（実施した日）→ 全店その店行の K〜N チェックを更新。
 * @returns {string}
 */
function fit365StoreHubSyncSmsToHub_(sheet) {
  var ctx = fit365StoreHubResolveContext_(sheet);
  if (ctx.error) return ctx.error;
  var hubSheet = ctx.hubSheet;
  var hitRow = ctx.hitRow;
  var hubSS = hubSheet.getParent();
  var dateBlock = sheet.getRange(fit365SmsStorePBlockA1_()).getValues();
  var hubSmsChk = hubSheet.getRange(fit365SmsHubStoreCheckboxRangeA1_(hitRow)).getValues()[0];
  var chkChanged = false;
  var i;
  for (i = 0; i < 4; i++) {
    var dayNum = fit365ParseSmsDayNumber_(dateBlock[i][0]);
    var want = dayNum !== null;
    var curCb = fit365CoerceCheckboxValue_(hubSmsChk[i]);
    var curBool = curCb === null ? false : curCb;
    if (curBool !== want) {
      hubSmsChk[i] = want;
      chkChanged = true;
    }
  }
  if (chkChanged) {
    fit365MarkStoreEchoOnHub_(hubSS, hitRow);
    var smsChkRng = hubSheet.getRange(fit365SmsHubStoreCheckboxRangeA1_(hitRow));
    smsChkRng.setValues([hubSmsChk]);
    var chkRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    smsChkRng.setDataValidation(chkRule);
  }
  return chkChanged
    ? "OK: SMS（P4:P7→全店のチェック）を反映しました。"
    : "OK: SMS（日付記入→チェック）変更なし。";
}

/**
 * 全店 2 行目 K〜N（DL）→ 各店 Q4:Q7 に上書き（全店マスター一方通行）。
 * @returns {string}
 */
function fit365StoreSmsRefreshDlFromHub_(sheet) {
  var ctx = fit365StoreHubResolveContext_(sheet);
  if (ctx.error) return ctx.error;
  var hubSheet = ctx.hubSheet;
  var rSched = FIT365_SMS_HUB_SCHEDULE_ROW;
  var sched = hubSheet.getRange(fit365SmsHubScheduleRangeA1_()).getValues()[0];
  var pv = [];
  var u;
  for (u = 0; u < 4; u++) {
    pv.push([sched[u] === "" || sched[u] === null ? "" : fit365NormalizeSmsHubScheduleCell_(sched[u])]);
  }
  sheet.getRange(fit365SmsStoreQBlockA1_()).setValues(pv);
  return "OK: 全店 " + rSched + " 行目 K〜N（DL）→ " + fit365SmsStoreQBlockA1_() + " を更新しました。";
}

/**
 * 6行目 B:E（→集計 C:F）・G:K（→集計 G:K）のチェックをまとめて集計へ書く。
 * @returns {string} 結果メッセージ（成功・失敗）
 */
function fit365StoreHubSyncRow6All_(sheet) {
  var ctx = fit365StoreHubResolveContext_(sheet);
  if (ctx.error) return ctx.error;

  var r = FIT365_STORE_HUB_SOURCE_ROW;
  var be = sheet.getRange(r, 2, r, 5).getValues()[0];
  var gk = sheet.getRange(r, 7, r, 11).getValues()[0];
  var hubSheet = ctx.hubSheet;
  var hitRow = ctx.hitRow;
  var hubSS = hubSheet.getParent();

  /** C:K の1行を A1 で指定（getRange(行,列,行,列) の4引数が「行数」と誤解釈されるのを避ける） */
  var hubCkRange = hubSheet.getRange("C" + hitRow + ":K" + hitRow);
  var hubRowVals = hubCkRange.getValues()[0];
  var changed = false;
  var j;
  for (j = 0; j < 4; j++) {
    var v = be[j];
    if (v !== true && v !== false) continue;
    if (hubRowVals[j] !== v) {
      hubRowVals[j] = v;
      changed = true;
    }
  }
  var k;
  for (k = 0; k < 5; k++) {
    var v2 = gk[k];
    if (v2 !== true && v2 !== false) continue;
    if (hubRowVals[4 + k] !== v2) {
      hubRowVals[4 + k] = v2;
      changed = true;
    }
  }
  if (changed) {
    fit365MarkStoreEchoOnHub_(hubSS, hitRow);
    hubCkRange.setValues([hubRowVals]);
    /** setValues だけだとチェックの検証が外れ TRUE/FALSE 文字になることがあるため付け直す */
    var chkRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    hubCkRange.setDataValidation(chkRule);
  }
  return "OK: 集計「" + hubSheet.getName() + "」 行" + hitRow + " を更新しました。";
}

/** メニューから: 同期フラグに関係なく、現在シートの6行目を集計へ送る（動作確認用） */
function fit365StoreHubPushRow6FromActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var msg = fit365StoreHubSyncRow6All_(sheet);
  SpreadsheetApp.getUi().alert(msg);
}

/** メニューから: 全店の DL 行→各店 Q、各店 P→全店チェック（動作確認用） */
function fit365StoreHubPushSmsFromActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "SMS 手動反映",
    fit365StoreSmsRefreshDlFromHub_(sheet) + "\n" + fit365StoreHubSyncSmsToHub_(sheet),
    ui.ButtonSet.OK
  );
}

/**
 * 【各店・初回1回】全店ブック接続の権限確認 ＋ 月次・編集のトリガー設定。
 * メニュー「初回セットアップ」から実行。
 */
function fit365InitialStoreSetup() {
  try {
    fit365AuthStep1Core_();
    fit365AuthStep2Core_();
    fit365OpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365OpAlertError_();
  }
}

/** 旧メニュー名との互換（実行ドロップダウン用） */
function fit365StoreHubSyncOneTimeSetup() {
  fit365InitialStoreSetup();
}

/**
 * 店舗シート6行目のチェック変更を、集計ブックの該当店舗行へ反映する。
 * インストール型 onEdit で実行すること。
 */
function fit365StoreHubCheckboxSyncOnEdit(e) {
  if (!FIT365_STORE_HUB_SYNC_ENABLED) return;
  if (!e || !e.range) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() === FIT365_SYNC_CONTROL_SHEET_NAME) return;
  if (fit365StoreShouldSkipEchoFromHub_(ss)) return;
  var row = range.getRow();
  var col = range.getColumn();
  if (fit365StoreSheetIsMonthlyTab_(sheet) && fit365StoreIsSmsSyncCell_(row, col)) {
    if (col === FIT365_SMS_STORE_COL_DL) {
      try {
        var rDl = fit365StoreSmsRefreshDlFromHub_(sheet);
        if (String(rDl).indexOf("OK:") !== 0) {
          fit365StoreHubNotifySyncIssue_(ss, "FIT365 SMS", rDl);
        } else {
          fit365SyncToast_(
            ss,
            "FIT365",
            "DL（Q列）は全店の " +
              FIT365_SMS_HUB_SCHEDULE_ROW +
              " 行目 K〜N の半角数字でのみ変更できます。実施した日は P4〜P7 に入力してください（DL は Q4〜Q7）。"
          );
        }
      } catch (errDl) {
        fit365StoreHubNotifySyncIssue_(ss, "FIT365 SMS", errDl);
      }
      return;
    }
    try {
      var msgSms = fit365StoreHubSyncSmsToHub_(sheet);
      if (String(msgSms).indexOf("OK:") !== 0) {
        fit365StoreHubNotifySyncIssue_(ss, "各店→全店 SMS", msgSms);
      }
    } catch (errSms) {
      fit365StoreHubNotifySyncIssue_(ss, "各店→全店 SMS エラー", errSms);
    }
    return;
  }
  if (row !== FIT365_STORE_HUB_SOURCE_ROW) return;
  if (!((col >= 2 && col <= 5) || (col >= 7 && col <= 11))) return;
  var vRaw = e.value !== undefined ? e.value : range.getValue();
  if (!fit365IsLikelyCheckboxEdit_(vRaw)) return;
  try {
    var msg = fit365StoreHubSyncRow6All_(sheet);
    if (String(msg).indexOf("OK:") !== 0) fit365StoreHubNotifySyncIssue_(ss, "各店→全店", msg);
  } catch (err) {
    fit365StoreHubNotifySyncIssue_(ss, "各店→全店 エラー", err);
  }
}

/**
 * 集計シートの店舗名列で、表示テキストが一致する行（先頭）を返す。無ければ -1。
 */
function fit365FindHubRowByStoreName_(hubSheet, storeName) {
  var last = hubSheet.getLastRow();
  if (last < FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW) return -1;
  var c = FIT365_STORE_HUB_TARGET_STORE_COL;
  var vals = hubSheet.getRange(FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW, c, last, c).getDisplayValues();
  var want = String(storeName || "").trim();
  for (var i = 0; i < vals.length; i++) {
    var cell = String(vals[i][0] || "").trim();
    var plain = cell.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
    if (plain === want || cell === want) return FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW + i;
  }
  return -1;
}

/** onEdit 同期トリガーをすべて削除（重複解消用） */
function removeFit365StoreHubOnEditTriggers_() {
  var fn = "fit365StoreHubCheckboxSyncOnEdit";
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * fit365StoreHubCheckboxSyncOnEdit 用の onEdit トリガーを1本だけ作成する
 * @returns {boolean} 常に新規作成したとき true
 */
function installFit365StoreHubCheckboxSyncTrigger() {
  var fn = "fit365StoreHubCheckboxSyncOnEdit";
  removeFit365StoreHubOnEditTriggers_();
  ScriptApp.newTrigger(fn)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  return true;
}