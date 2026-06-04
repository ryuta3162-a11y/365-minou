/**
 * FIT365 全店実施管理ブック専用
 *
 * 【貼る場所】全店・実施管理ブック 1冊だけ → 拡張機能 → Apps Script → Code.gs 全文置換
 * 【ツールバー】シート上段のカスタムメニュー「FIT365 全店」（onOpen で作成）
 *   成功時: 初回セットアップ / 更新 / リマインド送信 / リマインド自動設定
 *   古い表示「FIT365 全店→各店」→ このファイルが Google 上で未保存、または別ブックを編集
 *
 * - 店舗行のチェック → 各店6行目へ同期
 * - 色: チェックON→白 / 未チェック+2日前・前日・当日→黄 / 未チェック+超過→赤（メールなし）
 * - メール: 黄と同じく 2日前・前日・当日の未チェックのみ
 * - 2 行目 K〜N（SMS・DL の半角日付）→ 各店 Q4:Q7（一方通行）
 * - 店舗行 K〜N チェック → 各店 P4:P7（日付記入）
 *
 * 各店は FIT365/Code.gs。ループ防止: _FIT365_SYNC（A1=店舗行エコー、B1=DL行エコー予備）
 */

/** 貼付確認用（メニュー「スクリプト版確認」に表示） */
var FIT365_HUB_TOOLBAR_MENU_VERSION = "2026-06-hub-yellow3-red-v1";

var FIT365_STORE_HUB_SPREADSHEET_ID = "1jyeP8hZYLICEuZEQktwwDUqGgaatuQok69gC3vd0KDc";
var FIT365_HUB_TO_STORE_ROW6_SYNC_ENABLED = true;
var FIT365_HUB_TO_STORE_TARGET_SPREADSHEET_ID = "1VbEkmSibULcMWQtMrv9c7IItsu0aSh8H9oVu6PZswSg";
var FIT365_STORE_BOOK_MAP_SHEET_NAME = "店舗ブック";
var FIT365_SYNC_CONTROL_SHEET_NAME = "_FIT365_SYNC";
var FIT365_SYNC_ECHO_MS = 8000;
var FIT365_STORE_HUB_SOURCE_ROW = 6;
var FIT365_STORE_HUB_TARGET_STORE_COL = 2;
var FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW = 4;

/** SMS: 全店 2 行目 K〜N（DL）→各店 Q4:Q7。全店店舗行 K〜N チェック →各店 P4:P7（実施日） */
var FIT365_SMS_HUB_SCHEDULE_ROW = 2;
var FIT365_SMS_HUB_COL_FIRST = 11;
var FIT365_SMS_HUB_COL_LAST = 14;
var FIT365_SMS_STORE_ROWS_FIRST = 4;
var FIT365_SMS_STORE_ROWS_LAST = 7;
var FIT365_SMS_STORE_COL_DATE = 16; // P列: 実施日
var FIT365_SMS_STORE_COL_DL = 17;   // Q列: DL

/** リマインドメール設定 */
var FIT365_REMINDER_ENABLED = true;
var FIT365_REMINDER_CONTACT_SHEET_NAME_TEST = "連絡先サンプル";
var FIT365_REMINDER_LOG_SHEET_NAME = "_FIT365_REMINDER_LOG";
var FIT365_REMINDER_DAY_ROW = 2;
var FIT365_REMINDER_TASK_ROW = 3;
var FIT365_REMINDER_STORE_COL = 2;
var FIT365_REMINDER_STORE_START_ROW = 4;
var FIT365_REMINDER_CONTACT_STORE_COL = 1;
var FIT365_REMINDER_CONTACT_MAIL_COL_FIRST = 3;
var FIT365_REMINDER_CONTACT_MAIL_COL_LAST = 6;
var FIT365_REMINDER_COLOR_2DAYS = "#FB8C00"; // オレンジ
var FIT365_REMINDER_COLOR_1DAY = "#FFEB3B";  // 黄色
var FIT365_REMINDER_COLOR_TODAY = "#43A047"; // 緑
/** 色分け（チェック＋締切日。メール送信では色を変えない） */
var FIT365_HUB_COLOR_REMINDER_YELLOW = "#FFEB3B"; // 黄: 2日前・前日・当日（＝リマインド対象）
var FIT365_HUB_COLOR_OVERDUE_RED = "#FF5252"; // 赤: 超過（メール対象外）

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

var FIT365_MONTHLY_SHEET_NAME_RE = /^【([^】]+)】(\d{2})年(\d{1,2})月$/;

function fit365GetStoreSheetForHubTab_(storeSS, hubTabName, storeNamePlain) {
  var h = String(hubTabName || "").trim();
  var plain = String(storeNamePlain || "").trim();
  var candidates = [];
  if (plain && h) {
    candidates.push("【" + plain + "】" + h);
    var m = h.match(/^(\d{2})年(\d{1,2})月$/);
    if (m) {
      var normalized = m[1] + "年" + parseInt(m[2], 10) + "月";
      var br = "【" + plain + "】" + normalized;
      if (candidates.indexOf(br) === -1) candidates.push(br);
    }
  }
  var stripped = h.replace(/^【[^】]+】\s*/, "").trim();
  if (plain && stripped && stripped !== h) {
    var c2 = "【" + plain + "】" + stripped;
    if (candidates.indexOf(c2) === -1) candidates.push(c2);
  }
  if (h && candidates.indexOf(h) === -1) candidates.push(h);
  for (var i = 0; i < candidates.length; i++) {
    var sh = storeSS.getSheetByName(candidates[i]);
    if (sh) return sh;
  }
  return null;
}

function fit365ParseSpreadsheetId_(raw) {
  var s = String(raw || "").trim();
  if (!s) return "";
  var hy = s.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (hy) s = hy[1].trim();
  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{30,}$/.test(s)) return s;
  return "";
}

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

function fit365ParseSmsDayNumber_(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "boolean") return null;
  var n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^\d.]/g, ""));
  if (isNaN(n)) return null;
  n = Math.floor(Math.abs(n));
  if (n < 1 || n > 31) return null;
  return n;
}

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

function fit365HubSmsScheduleCellEdit_(vRaw) {
  if (vRaw === "" || vRaw === null || vRaw === undefined) return true;
  if (typeof vRaw === "number" && !isNaN(vRaw)) return true;
  return fit365ParseSmsDayNumber_(vRaw) !== null;
}

function fit365SyncToast_(ss, title, msg) {
  try {
    ss.toast(String(msg).slice(0, 280), String(title || "FIT365全店"), 12);
  } catch (ignore) {
    Logger.log(title + ": " + msg);
  }
}

function fit365FindStoreBookMapSheet_(hubSS) {
  var byName = hubSS.getSheetByName(FIT365_STORE_BOOK_MAP_SHEET_NAME);
  if (byName) return byName;
  var sheets = hubSS.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getName() === FIT365_SYNC_CONTROL_SHEET_NAME) continue;
    var a1 = String(sh.getRange(1, 1).getDisplayValue() || "");
    var b1 = String(sh.getRange(1, 2).getDisplayValue() || "");
    if (a1.indexOf("店舗名") !== -1 && (b1.indexOf("スプレッドシート") !== -1 || b1.indexOf("URL") !== -1 || b1.indexOf("ＩＤ") !== -1)) {
      return sh;
    }
  }
  return null;
}

function fit365EnsureControlSheet_(ss) {
  var name = FIT365_SYNC_CONTROL_SHEET_NAME;
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.hideSheet();
  return sh;
}

function fit365HubShouldSkipEchoFromStore_(hubSS, hubRow) {
  fit365EnsureControlSheet_(hubSS);
  var sh = hubSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME);
  if (hubRow === FIT365_SMS_HUB_SCHEDULE_ROW) {
    var bParts = String(sh.getRange(1, 2).getDisplayValue() || "").split("|");
    if (bParts[0] === "STORE_ECHO_ROW2") {
      var tsB = parseInt(bParts[1], 10);
      if (tsB && Date.now() - tsB <= FIT365_SYNC_ECHO_MS) {
        sh.getRange(1, 2).clearContent();
        return true;
      }
    }
    return false;
  }
  var parts = String(sh.getRange(1, 1).getDisplayValue() || "").split("|");
  if (parts[0] !== "STORE_ECHO") return false;
  if (parseInt(parts[1], 10) !== hubRow) return false;
  var ts = parseInt(parts[2], 10);
  if (!ts || Date.now() - ts > FIT365_SYNC_ECHO_MS) return false;
  sh.getRange(1, 1).clearContent();
  return true;
}

function fit365MarkHubEchoOnStore_(storeSS) {
  fit365EnsureControlSheet_(storeSS);
  storeSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME).getRange(1, 1).setValue("HUB_ECHO|" + Date.now());
}

/** 各店→全店の書き込み直後（全店 onEdit のループ防止） */
function fit365MarkStoreEchoOnHub_(hubSS, hubRow) {
  fit365EnsureControlSheet_(hubSS);
  hubSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME).getRange(1, 1).setValue("STORE_ECHO|" + hubRow + "|" + Date.now());
}

/** 更新バッチ後など：全店での手動編集が8秒間ブロックされないようエコーを解除 */
function fit365HubClearStoreEchoFlags_(hubSS) {
  fit365EnsureControlSheet_(hubSS);
  var sh = hubSS.getSheetByName(FIT365_SYNC_CONTROL_SHEET_NAME);
  sh.getRange(1, 1).clearContent();
  sh.getRange(1, 2).clearContent();
}

function fit365HubOpAlertDone_() {
  SpreadsheetApp.getUi().alert("完了しました", SpreadsheetApp.getUi().ButtonSet.OK);
}

function fit365HubOpAlertError_() {
  SpreadsheetApp.getUi().alert("エラー", SpreadsheetApp.getUi().ButtonSet.OK);
}

function fit365LoadStoreIdMap_(hubSS) {
  var out = {};
  var sheet = fit365FindStoreBookMapSheet_(hubSS);
  if (!sheet) return out;
  var last = sheet.getLastRow();
  if (last < 2) return out;
  var rowCount = last - 1;
  var rng = sheet.getRange(2, 1, rowCount, 2);
  var disp = rng.getDisplayValues();
  var vals = rng.getValues();
  var forms = rng.getFormulas();
  var rich = rng.getRichTextValues();
  for (var i = 0; i < disp.length; i++) {
    var nm = String(disp[i][0] || "").trim();
    nm = nm.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
    var id = fit365ParseSpreadsheetId_(vals[i][1]);
    if (!id && forms[i][1]) id = fit365ParseSpreadsheetId_(forms[i][1]);
    if (!id) id = fit365ParseSpreadsheetId_(disp[i][1]);
    if (!id && rich[i][1]) {
      var linkUrl = rich[i][1].getLinkUrl();
      if (linkUrl) id = fit365ParseSpreadsheetId_(linkUrl);
    }
    if (nm && id) out[nm] = id;
  }
  return out;
}

function fit365EnsureStoreBookMapTemplate_(hubSS) {
  if (fit365FindStoreBookMapSheet_(hubSS)) return;
  var sh = hubSS.insertSheet(FIT365_STORE_BOOK_MAP_SHEET_NAME);
  sh.getRange(1, 1, 1, 2).setValues([["店舗名", "スプレッドシートID（またはURL）"]]);
  sh.setColumnWidth(2, 420);
}

function fit365ResolveStoreSpreadsheetId_(hubSS, storeNamePlain) {
  var plain = String(storeNamePlain || "").trim();
  plain = plain.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
  var map = fit365LoadStoreIdMap_(hubSS);
  if (plain && map[plain]) return map[plain];
  if (!plain && FIT365_HUB_TO_STORE_TARGET_SPREADSHEET_ID) {
    return String(FIT365_HUB_TO_STORE_TARGET_SPREADSHEET_ID).trim();
  }
  return "";
}

function fit365WarmOpenStoreMapsForAuth_(hubSS) {
  var map = fit365LoadStoreIdMap_(hubSS);
  var n = 0;
  for (var k in map) {
    if (!map.hasOwnProperty(k)) continue;
    try {
      SpreadsheetApp.openById(map[k]).getName();
      n++;
      if (n >= 15) break;
    } catch (ignore) {}
  }
}

function onOpen() {
  fit365HubBuildToolbarMenu_();
}

/** スプレッドシート上段ツールバーに出すメニュー（全店ブック専用） */
function fit365HubBuildToolbarMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("FIT365 全店")
    .addItem("初回セットアップ", "fit365HubInitialSetup")
    .addItem("更新", "fit365HubDataUpdateMenu_")
    .addSeparator()
    .addItem("リマインド送信", "fit365ReminderRunManualToday")
    .addItem("リマインド自動設定", "fit365HubReminderSetupMenu_")
    .addSeparator()
    .addItem("スクリプト版確認", "fit365HubShowToolbarVersion_")
    .addToUi();
}

/**
 * メニューが古いままのとき: Apps Script でこの関数を選び「実行」→ シートを F5 で再読み込み
 * （onOpen はシートを開き直したときだけ動くため）
 */
function fit365HubRebuildToolbarMenu() {
  fit365HubBuildToolbarMenu_();
}

function fit365HubShowToolbarVersion_() {
  SpreadsheetApp.getUi().alert(
    "FIT365 全店（ツールバー）\n\n" +
      "スクリプト版: " + FIT365_HUB_TOOLBAR_MENU_VERSION + "\n\n" +
      "メニュー名が「FIT365 全店」なら最新です。\n" +
      "「FIT365 全店→各店」のままなら、Code.gs がこの版に置き換わっていません。"
  );
}

function fit365HubAuthCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) {
    throw new Error("全店ブック専用");
  }
  fit365EnsureControlSheet_(ss);
  fit365EnsureStoreBookMapTemplate_(ss);
  fit365WarmOpenStoreMapsForAuth_(ss);
  removeFit365HubOnEditTriggers_();
  installFit365HubToStoreRow6CheckboxSyncTrigger();
}

/** 初回1回：権限・店舗ブック接続・編集時トリガー */
function fit365HubInitialSetup() {
  try {
    fit365HubAuthCore_();
    fit365HubOpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365HubOpAlertError_();
  }
}

function fit365HubSyncOneTimeSetup() {
  fit365HubInitialSetup();
}

/**
 * 1店舗分：各店シートの6行目チェック・SMS実施日 → 全店の該当行（C〜K, K〜N）
 * @returns {boolean} 反映できたとき true
 */
function fit365HubPullFromStoreForRow_(hubSheet, hitRow) {
  var storeNameCell = hubSheet.getRange(hitRow, FIT365_STORE_HUB_TARGET_STORE_COL);
  var storeName = String(storeNameCell.getDisplayValue() || "").trim();
  storeName = storeName.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
  if (!storeName) return false;

  var hubSS = hubSheet.getParent();
  var storeId = fit365ResolveStoreSpreadsheetId_(hubSS, storeName);
  if (!storeId) {
    Logger.log("店舗ブック未登録: " + storeName);
    return false;
  }
  var storeSS;
  try {
    storeSS = SpreadsheetApp.openById(storeId);
  } catch (eOpen) {
    Logger.log("各店ブックを開けません: " + storeName + " " + eOpen);
    return false;
  }
  var storeSh = fit365GetStoreSheetForHubTab_(storeSS, hubSheet.getName(), storeName);
  if (!storeSh) {
    Logger.log("各店シートなし: " + storeName + " / " + hubSheet.getName());
    return false;
  }

  var r = FIT365_STORE_HUB_SOURCE_ROW;
  var be = storeSh.getRange(r, 2, r, 5).getValues()[0];
  var gk = storeSh.getRange(r, 7, r, 11).getValues()[0];
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
    var chkRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    hubCkRange.setDataValidation(chkRule);
  }

  var dateBlock = storeSh.getRange(fit365SmsStorePBlockA1_()).getValues();
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
    var chkRule2 = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    smsChkRng.setDataValidation(chkRule2);
  }

  return changed || chkChanged;
}

/**
 * 月初作業用：表示中の月次シート（26年6月 等）で
 * ①各店→全店（チェック・SMS実施状況の取り込み）
 * ②全店→各店（2行目のSMS日程 DL を各店へ配信）
 */
function fit365HubDataUpdateMenu_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) {
      throw new Error("全店ブック専用");
    }
    var hubSheet = ss.getActiveSheet();
    if (!fit365ReminderParseMonthSheetName_(hubSheet.getName())) {
      throw new Error("月次シート（YY年M月）を開いて実行してください");
    }
    var last = hubSheet.getLastRow();
    if (last < FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW) {
      throw new Error("店舗行がありません");
    }
    var ok = 0;
    var row;
    for (row = FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW; row <= last; row++) {
      var nm = String(hubSheet.getRange(row, FIT365_STORE_HUB_TARGET_STORE_COL).getDisplayValue() || "").trim();
      if (!nm) continue;
      if (fit365HubPullFromStoreForRow_(hubSheet, row)) ok++;
    }
    var smsMsg = fit365HubSmsScheduleRowToAllStores_(hubSheet);
    Logger.log("データ更新: 取込 " + ok + " 店 / SMS配信 " + smsMsg);
    if (String(smsMsg).indexOf("OK:") !== 0) {
      throw new Error(smsMsg);
    }
    fit365HubClearStoreEchoFlags_(ss);
    fit365HubRefreshAllTaskColors_(hubSheet, new Date());
    fit365HubOpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365HubOpAlertError_();
  }
}

/** 旧メニュー名との互換 */
function fit365HubPullAllStoresMenu_() {
  fit365HubDataUpdateMenu_();
}

/** リマインド担当者が1回：送信権限の確認＋毎朝の自動送信トリガー（メールは送らない） */
function fit365HubReminderSetupMenu_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) {
      throw new Error("全店ブック専用");
    }
    var contact = fit365ReminderFindContactSheet_(ss);
    if (contact) contact.getLastRow();
    MailApp.getRemainingDailyQuota();
    removeFit365HubReminderTriggers_();
    installFit365ReminderDailyTrigger();
    fit365HubOpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365HubOpAlertError_();
  }
}

function fit365HubInstallReminderTriggerMenu_() {
  fit365HubReminderSetupMenu_();
}

function removeFit365HubReminderTriggers_() {
  var fn = "fit365ReminderDailyTrigger";
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** 全店 2 行目 K〜N（DL）→店舗ブックに登録された全店の Q4:Q7 */
function fit365HubSmsScheduleRowToAllStores_(hubSheet) {
  var hubSS = hubSheet.getParent();
  var rSched = FIT365_SMS_HUB_SCHEDULE_ROW;
  var sched = hubSheet.getRange(fit365SmsHubScheduleRangeA1_()).getValues()[0];
  var map = fit365LoadStoreIdMap_(hubSS);
  var tabName = hubSheet.getName();
  var qBlock = fit365SmsStoreQBlockA1_();
  var cnt = 0;
  var miss = [];
  for (var nm in map) {
    if (!map.hasOwnProperty(nm)) continue;
    try {
      var st = SpreadsheetApp.openById(map[nm]);
      var sh = fit365GetStoreSheetForHubTab_(st, tabName, nm);
      if (!sh) {
        miss.push(nm);
        continue;
      }
      fit365MarkHubEchoOnStore_(st);
      var pv = [];
      for (var u = 0; u < 4; u++) {
        pv.push([sched[u] === "" || sched[u] === null ? "" : fit365NormalizeSmsHubScheduleCell_(sched[u])]);
      }
      sh.getRange(qBlock).setValues(pv);
      cnt++;
    } catch (ex) {
      miss.push(nm + "(" + ex + ")");
    }
  }
  var tail = miss.length ? "\n未反映: " + miss.join(", ") : "";
  return "OK: 全店 " + rSched + " 行目 K〜N（DL）を " + cnt + " 店の " + qBlock + " に反映しました。" + tail;
}

/** 全店 店舗行の K〜N チェック → 該当各店の P4:P7（実施日記入） */
function fit365HubSmsCheckboxesToStore_(hubSheet, hitRow) {
  if (hitRow < FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW) {
    return "データ行（B列に店名）を選んでください。";
  }
  var storeNameCell = hubSheet.getRange(hitRow, FIT365_STORE_HUB_TARGET_STORE_COL);
  var storeName = String(storeNameCell.getDisplayValue() || "").trim();
  storeName = storeName.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
  if (!storeName) return "B" + hitRow + " に店舗名がありません。";
  var hubSS = hubSheet.getParent();
  var storeId = fit365ResolveStoreSpreadsheetId_(hubSS, storeName);
  if (!storeId) return "「店舗ブック」に「" + storeName + "」のURLを登録してください。";
  var storeSS;
  try {
    storeSS = SpreadsheetApp.openById(storeId);
  } catch (e1) {
    return "各店ブックを開けません: " + e1;
  }
  var storeSh = fit365GetStoreSheetForHubTab_(storeSS, hubSheet.getName(), storeName);
  if (!storeSh) {
    return "各店に「【" + storeName + "】" + hubSheet.getName() + "」シートがありません。";
  }
  var sched = hubSheet.getRange(fit365SmsHubScheduleRangeA1_()).getValues()[0];
  var checksRow = hubSheet.getRange(fit365SmsHubStoreCheckboxRangeA1_(hitRow)).getValues()[0];
  var pBlock = fit365SmsStorePBlockA1_();
  var existingP = storeSh.getRange(pBlock).getValues();
  fit365MarkHubEchoOnStore_(storeSS);
  var outP = [];
  var j;
  for (j = 0; j < 4; j++) {
    var c = fit365CoerceCheckboxValue_(checksRow[j]);
    var isOn = c === true;
    if (!isOn) {
      outP.push([""]);
      continue;
    }
    var d = fit365NormalizeSmsHubScheduleCell_(sched[j]);
    if (d !== "" && d !== null) {
      outP.push([d]);
    } else {
      outP.push([existingP[j][0]]);
    }
  }
  storeSh.getRange(pBlock).setValues(outP);
  return "OK: 「" + storeName + "」の " + pBlock + " を更新しました。";
}

function fit365HubToStoreSyncRow6AllForHubRow_(hubSheet, hitRow) {
  if (hitRow < FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW) {
    return "データ行（B列に店名）を選んでください。";
  }
  var storeNameCell = hubSheet.getRange(hitRow, FIT365_STORE_HUB_TARGET_STORE_COL);
  var storeName = String(storeNameCell.getDisplayValue() || "").trim();
  storeName = storeName.replace(/^[「『【](.+)[」』】]$/, "$1").trim();
  if (!storeName) return "B" + hitRow + " に店舗名がありません。";
  var hubSS = hubSheet.getParent();
  var storeId = fit365ResolveStoreSpreadsheetId_(hubSS, storeName);
  if (!storeId) return "「店舗ブック」に「" + storeName + "」のURLを登録してください。";
  var storeSS;
  try {
    storeSS = SpreadsheetApp.openById(storeId);
  } catch (e1) {
    return "各店ブックを開けません: " + e1;
  }
  var storeSh = fit365GetStoreSheetForHubTab_(storeSS, hubSheet.getName(), storeName);
  if (!storeSh) {
    return "各店に「【" + storeName + "】" + hubSheet.getName() + "」シートがありません。";
  }
  var hubRng = hubSheet.getRange("C" + hitRow + ":K" + hitRow);
  var hubVals = hubRng.getValues()[0];
  var r = FIT365_STORE_HUB_SOURCE_ROW;
  var storeBk = storeSh.getRange("B" + r + ":K" + r);
  var storeRow = storeBk.getValues()[0];
  var j;
  for (j = 0; j < 4; j++) {
    var v = hubVals[j];
    if (v === true || v === false) storeRow[j] = v;
  }
  for (var k = 0; k < 5; k++) {
    var v2 = hubVals[4 + k];
    if (v2 === true || v2 === false) storeRow[5 + k] = v2;
  }
  fit365MarkHubEchoOnStore_(storeSS);
  storeBk.setValues([storeRow]);
  var chk = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  storeSh.getRange("B" + r + ":E" + r).setDataValidation(chk);
  storeSh.getRange("G" + r + ":K" + r).setDataValidation(chk);
  return "OK: 「" + storeName + "」" + storeSh.getName() + " の" + r + "行目を更新しました。";
}

function fit365HubToStoreRow6CheckboxSyncOnEdit(e) {
  if (!FIT365_HUB_TO_STORE_ROW6_SYNC_ENABLED) return;
  if (!e || !e.range) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) return;
  var hubSheet = e.range.getSheet();
  if (hubSheet.getName() === FIT365_SYNC_CONTROL_SHEET_NAME) return;
  if (hubSheet.getName() === FIT365_STORE_BOOK_MAP_SHEET_NAME) return;
  var mapSh = fit365FindStoreBookMapSheet_(ss);
  if (mapSh && hubSheet.getSheetId() === mapSh.getSheetId()) return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var vRaw = e.value !== undefined ? e.value : e.range.getValue();

  if (row === FIT365_SMS_HUB_SCHEDULE_ROW && col >= FIT365_SMS_HUB_COL_FIRST && col <= FIT365_SMS_HUB_COL_LAST) {
    if (!fit365HubSmsScheduleCellEdit_(vRaw)) return;
    if (fit365HubShouldSkipEchoFromStore_(ss, FIT365_SMS_HUB_SCHEDULE_ROW)) return;
    try {
      var msgSched = fit365HubSmsScheduleRowToAllStores_(hubSheet);
      if (String(msgSched).indexOf("OK:") !== 0) fit365SyncToast_(ss, "全店→各店 SMS予定", msgSched);
    } catch (errS) {
      fit365SyncToast_(ss, "全店→各店 SMS予定 エラー", String(errS));
    }
    return;
  }

  if (row >= FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW && fit365HubIsTaskColumn_(hubSheet, col)) {
    fit365HubOnStoreRowCheckboxEdit_(hubSheet, row, col, vRaw);
    return;
  }
}

function fit365HubToStorePushFromActiveHubRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveSheet().getActiveRange().getRow();
  SpreadsheetApp.getUi().alert(fit365HubToStoreSyncRow6AllForHubRow_(ss.getActiveSheet(), row));
}

function removeFit365HubOnEditTriggers_() {
  var fn = "fit365HubToStoreRow6CheckboxSyncOnEdit";
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function installFit365HubToStoreRow6CheckboxSyncTrigger() {
  var fn = "fit365HubToStoreRow6CheckboxSyncOnEdit";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hubId = FIT365_STORE_HUB_SPREADSHEET_ID;
  if (ss.getId() !== hubId) return false;
  removeFit365HubOnEditTriggers_();
  ScriptApp.newTrigger(fn).forSpreadsheet(ss).onEdit().create();
  return true;
}

function fit365NormalizeStoreName_(name) {
  return String(name || "")
    .trim()
    .replace(/^[「『【](.+)[」』】]$/, "$1")
    .replace(/[ \u3000]+/g, " ")
    .trim();
}

function fit365ReminderParseMonthSheetName_(sheetName) {
  var m = String(sheetName || "").match(/^(\d{2})年(\d{1,2})月$/);
  if (!m) return null;
  var yy = parseInt(m[1], 10);
  var mm = parseInt(m[2], 10);
  if (isNaN(yy) || isNaN(mm) || mm < 1 || mm > 12) return null;
  return { fullYear: 2000 + yy, month: mm };
}

function fit365ReminderResolveTargetSheet_(ss) {
  var now = new Date();
  var curLabel = String(now.getFullYear()).slice(-2) + "年" + (now.getMonth() + 1) + "月";
  var direct = ss.getSheetByName(curLabel);
  if (direct) return direct;
  var sheets = ss.getSheets();
  var latest = null;
  for (var i = 0; i < sheets.length; i++) {
    var info = fit365ReminderParseMonthSheetName_(sheets[i].getName());
    if (!info) continue;
    if (!latest) {
      latest = { sheet: sheets[i], y: info.fullYear, m: info.month };
      continue;
    }
    if (info.fullYear > latest.y || (info.fullYear === latest.y && info.month > latest.m)) {
      latest = { sheet: sheets[i], y: info.fullYear, m: info.month };
    }
  }
  return latest ? latest.sheet : null;
}

/** メニュー操作は表示中タブ、時間トリガーは当月タブを優先 */
function fit365HubResolveMonthSheetForOps_(ss, useActiveSheet) {
  if (useActiveSheet) {
    var active = ss.getActiveSheet();
    if (active && fit365ReminderParseMonthSheetName_(active.getName())) return active;
    return null;
  }
  return fit365ReminderResolveTargetSheet_(ss);
}

function fit365HubNormalizeBgHex_(color) {
  var s = String(color || "").trim().toLowerCase();
  if (!s) return "";
  if (s.charAt(0) !== "#") s = "#" + s;
  if (s.length === 4) {
    s = "#" + s.charAt(1) + s.charAt(1) + s.charAt(2) + s.charAt(2) + s.charAt(3) + s.charAt(3);
  }
  return s;
}

function fit365HubIsOurOverdueRed_(bg) {
  return fit365HubNormalizeBgHex_(bg) === fit365HubNormalizeBgHex_(FIT365_HUB_OVERDUE_RED);
}

function fit365HubIsTaskColumn_(hubSheet, col) {
  var tasks = fit365ReminderExtractTaskColumns_(hubSheet);
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].col === col) return true;
  }
  return false;
}

/** 2行目の日付＋シート名の年月から締切日（日付のみ） */
function fit365HubTaskDueDate_(hubSheet, sheetInfo, task) {
  if (task.sameDayOnly) {
    var header = String(hubSheet.getRange(FIT365_REMINDER_DAY_ROW, task.col).getDisplayValue() || "");
    if (header.indexOf("翌月") !== -1) {
      var nm = sheetInfo.month + 1;
      var ny = sheetInfo.fullYear;
      if (nm > 12) {
        nm = 1;
        ny++;
      }
      return new Date(ny, nm - 1, 1);
    }
    return new Date(sheetInfo.fullYear, sheetInfo.month - 1, task.day);
  }
  return new Date(sheetInfo.fullYear, sheetInfo.month - 1, task.day);
}

/** 締切日の翌日以降＝超過 */
function fit365HubIsTaskOverdue_(hubSheet, sheetInfo, task, today) {
  var due = fit365HubTaskDueDate_(hubSheet, sheetInfo, task);
  var baseToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return baseToday.getTime() > due.getTime();
}

/**
 * 背景色: チェックON→白 / 未チェック+2日前・前日・当日→黄 / 未チェック+超過→赤 / それ以外→白
 * @returns {number} 黄または赤にしたセル数
 */
function fit365HubPaintTaskCellColor_(hubSheet, sheetInfo, row, task, todayDate) {
  var cell = hubSheet.getRange(row, task.col);
  var val = cell.getValue();
  var isDone = val === true || String(val).toUpperCase() === "TRUE";
  if (isDone) {
    cell.setBackground(null);
    return 0;
  }
  var timing = fit365ReminderTimingForTask_(hubSheet, sheetInfo, task, todayDate);
  if (timing === "超過") {
    cell.setBackground(FIT365_HUB_COLOR_OVERDUE_RED);
    return 1;
  }
  if (timing === "2日前" || timing === "前日" || timing === "当日") {
    cell.setBackground(FIT365_HUB_COLOR_REMINDER_YELLOW);
    return 1;
  }
  cell.setBackground(null);
  return 0;
}

/** 1店舗行のタスク列だけ色を再計算（チェック変更直後用） */
function fit365HubRefreshTaskColorsForRow_(hubSheet, row, today) {
  var sheetInfo = fit365ReminderParseMonthSheetName_(hubSheet.getName());
  if (!sheetInfo) return 0;
  if (row < FIT365_REMINDER_STORE_START_ROW) return 0;
  var storeName = String(hubSheet.getRange(row, FIT365_REMINDER_STORE_COL).getDisplayValue() || "").trim();
  if (!storeName) return 0;
  var tasks = fit365ReminderExtractTaskColumns_(hubSheet);
  if (!tasks.length) return 0;
  var todayDate = today instanceof Date ? today : new Date();
  var painted = 0;
  var t;
  for (t = 0; t < tasks.length; t++) {
    painted += fit365HubPaintTaskCellColor_(hubSheet, sheetInfo, row, tasks[t], todayDate);
  }
  return painted;
}

/** 表示中タブの全店舗行を、チェック＋超過ルールで再着色（更新・リマインド後に実行） */
function fit365HubRefreshAllTaskColors_(hubSheet, today) {
  var sheetInfo = fit365ReminderParseMonthSheetName_(hubSheet.getName());
  if (!sheetInfo) return 0;
  var tasks = fit365ReminderExtractTaskColumns_(hubSheet);
  if (!tasks.length) return 0;
  var lastRow = hubSheet.getLastRow();
  if (lastRow < FIT365_REMINDER_STORE_START_ROW) return 0;
  var todayDate = today instanceof Date ? today : new Date();
  var painted = 0;
  var row;
  for (row = FIT365_REMINDER_STORE_START_ROW; row <= lastRow; row++) {
    painted += fit365HubRefreshTaskColorsForRow_(hubSheet, row, todayDate);
  }
  return painted;
}

/** 旧名との互換 */
function fit365HubApplyOverdueRedHighlight_(hubSheet, today) {
  return fit365HubRefreshAllTaskColors_(hubSheet, today);
}

/**
 * 全店シートでチェックを付け外ししたとき：各店へ反映＋その行の色を更新
 */
function fit365HubOnStoreRowCheckboxEdit_(hubSheet, hitRow, col, vRaw) {
  if (hitRow < FIT365_STORE_HUB_TARGET_SCAN_MIN_ROW) return;
  if (!fit365IsLikelyCheckboxEdit_(vRaw)) return;
  if (!fit365ReminderParseMonthSheetName_(hubSheet.getName())) return;

  var hubSS = hubSheet.getParent();
  var echoSkip = fit365HubShouldSkipEchoFromStore_(hubSS, hitRow);
  if (!echoSkip) {
    if (col >= 3 && col <= 11) {
      try {
        var msgRow = fit365HubToStoreSyncRow6AllForHubRow_(hubSheet, hitRow);
        if (String(msgRow).indexOf("OK:") !== 0) fit365SyncToast_(hubSS, "全店→各店", msgRow);
      } catch (errRow) {
        fit365SyncToast_(hubSS, "全店→各店 エラー", String(errRow));
      }
    }
    if (col >= FIT365_SMS_HUB_COL_FIRST && col <= FIT365_SMS_HUB_COL_LAST) {
      try {
        var msgSms = fit365HubSmsCheckboxesToStore_(hubSheet, hitRow);
        if (String(msgSms).indexOf("OK:") !== 0) fit365SyncToast_(hubSS, "全店→各店 SMS", msgSms);
      } catch (errSms) {
        fit365SyncToast_(hubSS, "全店→各店 SMS エラー", String(errSms));
      }
    }
  }
  fit365HubRefreshTaskColorsForRow_(hubSheet, hitRow, new Date());
}

function fit365ReminderParseDayCell_(vRaw) {
  if (vRaw === null || vRaw === undefined || vRaw === "") return null;
  if (typeof vRaw === "number" && !isNaN(vRaw)) {
    var dayNum = Math.floor(Math.abs(vRaw));
    if (dayNum >= 1 && dayNum <= 31) return { day: dayNum, sameDayOnly: false };
    return null;
  }
  var s = String(vRaw).trim();
  if (!s) return null;
  var sameDayOnly = s.indexOf("翌月1日") !== -1;
  var m = s.match(/(\d{1,2})/g);
  if (!m || m.length === 0) return null;
  var pick = sameDayOnly ? 1 : parseInt(m[m.length - 1], 10);
  if (isNaN(pick) || pick < 1 || pick > 31) return null;
  return { day: pick, sameDayOnly: sameDayOnly };
}

function fit365ReminderExtractTaskColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < FIT365_REMINDER_STORE_COL + 1) return [];
  var dayVals = sheet.getRange(FIT365_REMINDER_DAY_ROW, 1, 1, lastCol).getDisplayValues()[0];
  var taskVals = sheet.getRange(FIT365_REMINDER_TASK_ROW, 1, 1, lastCol).getDisplayValues()[0];
  var tasks = [];
  for (var c = FIT365_REMINDER_STORE_COL + 1; c <= lastCol; c++) {
    var dayInfo = fit365ReminderParseDayCell_(dayVals[c - 1]);
    if (!dayInfo) continue;
    var taskName = String(taskVals[c - 1] || "").trim();
    if (!taskName) continue;
    tasks.push({
      col: c,
      day: dayInfo.day,
      sameDayOnly: dayInfo.sameDayOnly,
      taskName: taskName
    });
  }
  return tasks;
}

/** タスク1件の送信タイミング（null=今日は送らない） */
function fit365ReminderTimingForTask_(hubSheet, sheetInfo, task, today) {
  var baseToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var due = fit365HubTaskDueDate_(hubSheet, sheetInfo, task);
  var deltaDays = Math.round((due.getTime() - baseToday.getTime()) / 86400000);
  if (deltaDays < 0) return "超過";
  if (task.sameDayOnly) {
    if (deltaDays === 0) return "当日";
    return null;
  }
  if (deltaDays === 2) return "2日前";
  if (deltaDays === 1) return "前日";
  if (deltaDays === 0) return "当日";
  return null;
}

function fit365ReminderSelectDueTasks_(hubSheet, sheetInfo, tasks, today) {
  var out = [];
  for (var i = 0; i < tasks.length; i++) {
    var timing = fit365ReminderTimingForTask_(hubSheet, sheetInfo, tasks[i], today);
    if (timing && timing !== "超過") out.push({ task: tasks[i], timing: timing });
  }
  return out;
}

/**
 * 1店舗行の未チェックタスクのうち、送信対象を集める
 * @param {string[]} allowedTimings 例: ["2日前","前日","当日"] または 手動時は "超過" も含む
 */
function fit365ReminderCollectPendingForRow_(tableRow, tasks, hubSheet, sheetInfo, today, ctx) {
  var pending = [];
  var sheetName = ctx.sheetName;
  var storeName = ctx.storeName;
  var logged = ctx.logged || {};
  var allowed = ctx.allowedTimings || [];
  var allowMap = {};
  var a;
  for (a = 0; a < allowed.length; a++) allowMap[allowed[a]] = true;

  for (var i = 0; i < tasks.length; i++) {
    var task = tasks[i];
    var timing = fit365ReminderTimingForTask_(hubSheet, sheetInfo, task, today);
    if (!timing || !allowMap[timing]) continue;
    var cIdx = task.col - 1;
    if (cIdx < 0 || cIdx >= tableRow.length) continue;
    var val = tableRow[cIdx];
    var isDone = val === true || String(val).toUpperCase() === "TRUE";
    if (isDone) continue;
    var key = [sheetName, storeName, task.taskName, timing].join("|");
    if (logged[key]) continue;
    pending.push({
      taskName: task.taskName,
      day: task.day,
      timing: timing,
      logKey: key,
      col: task.col
    });
  }
  return pending;
}

function fit365ReminderTimingLabel_(timing) {
  if (timing === "超過") return "期限超過";
  return "DL" + timing;
}

function fit365ReminderSelectForcedTasks_(tasks, forcedTimings) {
  var out = [];
  if (!forcedTimings || !forcedTimings.length) return out;
  var allow = {};
  for (var i = 0; i < forcedTimings.length; i++) allow[String(forcedTimings[i])] = true;
  for (var t = 0; t < tasks.length; t++) {
    if (allow["2日前"]) out.push({ task: tasks[t], timing: "2日前" });
    if (allow["前日"]) out.push({ task: tasks[t], timing: "前日" });
    if (allow["当日"]) out.push({ task: tasks[t], timing: "当日" });
  }
  return out;
}

function fit365ReminderFindContactSheet_(ss) {
  return ss.getSheetByName("連絡先") || ss.getSheetByName(FIT365_REMINDER_CONTACT_SHEET_NAME_TEST);
}

function fit365ReminderParseEmails_(raw) {
  var s = String(raw || "").trim();
  if (!s) return [];
  var parts = s.split(/[,\n;、，\s]+/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = String(parts[i] || "").trim();
    if (!p) continue;
    if (p.indexOf("@") === -1) continue;
    out.push(p);
  }
  return out;
}

function fit365ReminderLoadContactMap_(ss) {
  var map = {};
  var sh = fit365ReminderFindContactSheet_(ss);
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var width = FIT365_REMINDER_CONTACT_MAIL_COL_LAST;
  var vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var store = fit365NormalizeStoreName_(vals[i][FIT365_REMINDER_CONTACT_STORE_COL - 1]);
    if (!store) continue;
    var emails = [];
    for (var c = FIT365_REMINDER_CONTACT_MAIL_COL_FIRST; c <= FIT365_REMINDER_CONTACT_MAIL_COL_LAST; c++) {
      var parsed = fit365ReminderParseEmails_(vals[i][c - 1]);
      for (var k = 0; k < parsed.length; k++) emails.push(parsed[k]);
    }
    if (!emails.length) continue;
    var uniq = {};
    var arr = [];
    for (var e = 0; e < emails.length; e++) {
      var key = emails[e].toLowerCase();
      if (uniq[key]) continue;
      uniq[key] = true;
      arr.push(emails[e]);
    }
    map[store] = arr;
  }
  return map;
}

function fit365ReminderEnsureLogSheet_(ss) {
  var sh = ss.getSheetByName(FIT365_REMINDER_LOG_SHEET_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(FIT365_REMINDER_LOG_SHEET_NAME);
  sh.hideSheet();
  sh.getRange(1, 1, 1, 6).setValues([["timestamp", "date", "sheet", "store", "task", "timing"]]);
  return sh;
}

function fit365ReminderLoadTodayLogKeys_(ss, dateLabel) {
  var sh = fit365ReminderEnsureLogSheet_(ss);
  var last = sh.getLastRow();
  var set = {};
  if (last < 2) return set;
  var vals = sh.getRange(2, 2, last - 1, 5).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "") !== dateLabel) continue;
    var key = [vals[i][1], vals[i][2], vals[i][3], vals[i][4]].join("|");
    set[key] = true;
  }
  return set;
}

function fit365ReminderAppendLogRows_(ss, rows) {
  if (!rows || !rows.length) return;
  var sh = fit365ReminderEnsureLogSheet_(ss);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function fit365ReminderBuildMessage_(storeName, sheetName, pending, url) {
  var lines = [];
  lines.push(storeName + " ご担当者様");
  lines.push("");
  lines.push("リマインドです。下記スプレッドシートからご対応をお願いします。");
  lines.push("");
  for (var i = 0; i < pending.length; i++) {
    lines.push(
      "・" + pending[i].taskName + "（" + fit365ReminderTimingLabel_(pending[i].timing) + " / 締切" + pending[i].day + "日）"
    );
  }
  lines.push("");
  lines.push("確認URL: " + url);
  lines.push("");
  lines.push("未実施タスク実施後、チェックをお願いいたします。");
  return lines.join("\n");
}

function fit365ReminderEscapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function fit365ReminderBuildHtmlMessage_(storeName, sheetName, pending, url) {
  var rows = "";
  for (var i = 0; i < pending.length; i++) {
    rows +=
      "<tr>" +
      "<td style='padding:10px 12px;border-bottom:1px solid #f3d5e1;'>" + fit365ReminderEscapeHtml_(pending[i].taskName) + "</td>" +
      "<td style='padding:10px 12px;border-bottom:1px solid #f3d5e1;text-align:center;white-space:nowrap;'>" +
      fit365ReminderEscapeHtml_(fit365ReminderTimingLabel_(pending[i].timing)) +
      "</td>" +
      "<td style='padding:10px 12px;border-bottom:1px solid #f3d5e1;text-align:center;white-space:nowrap;'>" + fit365ReminderEscapeHtml_(pending[i].day + "日") + "</td>" +
      "</tr>";
  }
  return (
    "<div style='background:#f8f9fb;padding:24px 0;font-family:Meiryo,Arial,sans-serif;color:#333;'>" +
    "<div style='max-width:680px;margin:0 auto;background:#ffffff;border:1px solid #f1d6e2;border-radius:10px;overflow:hidden;'>" +
    "<div style='background:linear-gradient(90deg,#e91e63 0%,#c2185b 100%);padding:16px 20px;color:#fff;font-size:18px;font-weight:bold;'>" +
    "FIT365 リマインド" +
    "</div>" +
    "<div style='padding:20px;line-height:1.7;'>" +
    "<p style='margin:0 0 8px 0;'>" + fit365ReminderEscapeHtml_(storeName) + " ご担当者様</p>" +
    "<p style='margin:0 0 16px 0;'>リマインドです。下記スプレッドシートからご対応をお願いします。</p>" +
    "<table style='width:100%;border-collapse:collapse;border:1px solid #f3d5e1;border-radius:8px;overflow:hidden;'>" +
    "<thead><tr style='background:#fff3f8;color:#7a1b43;'>" +
    "<th style='padding:10px 12px;text-align:left;font-weight:bold;'>未実施タスク</th>" +
    "<th style='padding:10px 12px;text-align:center;font-weight:bold;'>DL</th>" +
    "<th style='padding:10px 12px;text-align:center;font-weight:bold;'>実施日</th>" +
    "</tr></thead>" +
    "<tbody>" + rows + "</tbody>" +
    "</table>" +
    "<div style='margin-top:16px;'>" +
    "<a href='" + fit365ReminderEscapeHtml_(url) + "' style='display:inline-block;background:#7c4dff;color:#fff;text-decoration:none;padding:10px 16px;border-radius:6px;font-weight:bold;'>スプレッドシートを開く</a>" +
    "</div>" +
    "<p style='margin:16px 0 0 0;'>未実施タスク実施後、チェックをお願いいたします。</p>" +
    "<p style='margin:10px 0 0 0;color:#777;font-size:12px;'>対象シート: " + fit365ReminderEscapeHtml_(sheetName) + "</p>" +
    "<p style='margin:4px 0 0 0;color:#777;font-size:12px;'>このメールは自動送信です。</p>" +
    "</div></div></div>"
  );
}

function fit365ReminderResolveStoreSheetUrl_(hubSS, hubSheetName, storeName) {
  var sid = fit365ResolveStoreSpreadsheetId_(hubSS, storeName);
  if (!sid) return "";
  try {
    var storeSS = SpreadsheetApp.openById(sid);
    var sh = fit365GetStoreSheetForHubTab_(storeSS, hubSheetName, storeName);
    if (sh) return storeSS.getUrl() + "#gid=" + sh.getSheetId();
    return storeSS.getUrl();
  } catch (e) {
    return "";
  }
}

function fit365ReminderColorForTiming_(timing) {
  return "";
}

function fit365SendReminderEmailsForToday_(opt) {
  if (!FIT365_REMINDER_ENABLED) return "リマインド設定が無効です。";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) {
    return "この処理は全店ブックで実行してください。";
  }

  var options = {};
  if (opt instanceof Date) {
    options.today = opt;
  } else if (opt && typeof opt === "object") {
    options = opt;
  }
  var today = options.today instanceof Date ? options.today : new Date();
  var useActive = !!options.useActiveSheet;
  var targetSheet = fit365HubResolveMonthSheetForOps_(ss, useActive);
  if (!targetSheet) {
    return useActive
      ? "月次シート（YY年M月）を開いて実行してください。"
      : "月次タブ（YY年M月）が見つかりません。";
  }
  var sheetInfo = fit365ReminderParseMonthSheetName_(targetSheet.getName());
  if (!sheetInfo) return "対象シート名が YY年M月 形式ではありません: " + targetSheet.getName();

  var dateLabel = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var tasks = fit365ReminderExtractTaskColumns_(targetSheet);
  if (!tasks.length) return "リマインド対象タスク（日付行+内容行）が見つかりません。";

  var contactSh = fit365ReminderFindContactSheet_(ss);
  if (!contactSh) {
    return "「連絡先」または「" + FIT365_REMINDER_CONTACT_SHEET_NAME_TEST + "」シートがありません。";
  }

  var manualSend = !!options.manualSend;
  /** メール＝黄と同じ（2日前・前日・当日の未チェック。超過は赤のみで送信しない） */
  var allowedTimings = options.forceTimings ? options.forceTimings : ["2日前", "前日", "当日"];

  var contactMap = fit365ReminderLoadContactMap_(ss);
  var lastRow = targetSheet.getLastRow();
  if (lastRow < FIT365_REMINDER_STORE_START_ROW) return "店舗行がありません。";
  var table = targetSheet.getRange(FIT365_REMINDER_STORE_START_ROW, 1, lastRow - FIT365_REMINDER_STORE_START_ROW + 1, targetSheet.getLastColumn()).getValues();
  var logged = options.ignoreLog ? {} : fit365ReminderLoadTodayLogKeys_(ss, dateLabel);
  var logsToAdd = [];
  var sentStores = 0;
  var sentTasks = 0;
  var skippedNoMail = [];
  var skippedAlreadySent = 0;
  var fallbackUrl = ss.getUrl() + "#gid=" + targetSheet.getSheetId();
  var sheetLabel = targetSheet.getName();

  for (var r = 0; r < table.length; r++) {
    var storeRaw = table[r][FIT365_REMINDER_STORE_COL - 1];
    var storeName = fit365NormalizeStoreName_(storeRaw);
    if (!storeName) continue;
    var recipients = contactMap[storeName];
    if (!recipients || !recipients.length) {
      skippedNoMail.push(storeName);
      continue;
    }

    var pending;
    if (options.forceTimings) {
      pending = [];
      var dueForced = fit365ReminderSelectForcedTasks_(tasks, options.forceTimings);
      var d;
      for (d = 0; d < dueForced.length; d++) {
        var td = dueForced[d];
        var cIdxF = td.task.col - 1;
        if (cIdxF < 0 || cIdxF >= table[r].length) continue;
        var valF = table[r][cIdxF];
        if (valF === true || String(valF).toUpperCase() === "TRUE") continue;
        var keyF = [sheetLabel, storeName, td.task.taskName, td.timing].join("|");
        if (logged[keyF]) {
          skippedAlreadySent++;
          continue;
        }
        pending.push({
          taskName: td.task.taskName,
          day: td.task.day,
          timing: td.timing,
          logKey: keyF,
          col: td.task.col
        });
      }
    } else {
      pending = fit365ReminderCollectPendingForRow_(table[r], tasks, targetSheet, sheetInfo, today, {
        sheetName: sheetLabel,
        storeName: storeName,
        logged: logged,
        allowedTimings: allowedTimings
      });
    }

    if (!pending.length) continue;

    var url = fit365ReminderResolveStoreSheetUrl_(ss, sheetLabel, storeName) || fallbackUrl;
    var subjectPrefix = options.subjectPrefix ? String(options.subjectPrefix) : "";
    var subject = subjectPrefix + "【FIT365】リマインド通知 | " + storeName + " | " + sheetLabel;
    var body = fit365ReminderBuildMessage_(storeName, sheetLabel, pending, url);
    var htmlBody = fit365ReminderBuildHtmlMessage_(storeName, sheetLabel, pending, url);
    MailApp.sendEmail({
      to: recipients.join(","),
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
    sentStores++;
    sentTasks += pending.length;
    for (var p = 0; p < pending.length; p++) {
      logged[pending[p].logKey] = true;
      logsToAdd.push([new Date(), dateLabel, sheetLabel, storeName, pending[p].taskName, pending[p].timing]);
    }
  }

  if (!options.ignoreLog) fit365ReminderAppendLogRows_(ss, logsToAdd);
  var colorCnt = fit365HubRefreshAllTaskColors_(targetSheet, today);

  var lines = [];
  lines.push("対象シート: " + sheetLabel);
  lines.push("連絡先: " + contactSh.getName());
  if (manualSend) {
    lines.push("手動送信（2日前・前日・当日の未チェック＝黄色のセル）");
  } else {
    lines.push("自動送信（2日前・前日・当日）");
  }
  if (sentStores > 0) {
    lines.push("送信: " + sentStores + "店舗 / " + sentTasks + "タスク");
  } else {
    lines.push("送信: 0件（未チェックの該当タスクなし、または本日送信済み）");
  }
  if (skippedNoMail.length) {
    lines.push("宛先なし（連絡先に店舗名・メール未登録）: " + skippedNoMail.join(", "));
  }
  if (skippedAlreadySent > 0) {
    lines.push("本日送信済みでスキップ: " + skippedAlreadySent + "件");
  }
  lines.push(
    "色分け再計算: " +
      colorCnt +
      "セル（黄=2日前・前日・当日 / 赤=超過 / 白=チェック済・それ以外）"
  );
  return lines.join("\n");
}

function fit365ReminderRunManualToday() {
  try {
    var msg = fit365SendReminderEmailsForToday_({
      today: new Date(),
      useActiveSheet: true,
      manualSend: true
    });
    SpreadsheetApp.getUi().alert("リマインド送信", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    fit365HubOpAlertError_();
  }
}

function fit365ReminderRunManualTestAllTimings() {
  try {
    fit365SendReminderEmailsForToday_({
      today: new Date(),
      useActiveSheet: true,
      forceTimings: ["2日前", "前日", "当日"],
      ignoreLog: true,
      subjectPrefix: "【テスト送信】"
    });
    fit365HubOpAlertDone_();
  } catch (e) {
    Logger.log(e);
    fit365HubOpAlertError_();
  }
}

function fit365ReminderDailyTrigger() {
  var msg = fit365SendReminderEmailsForToday_(new Date());
  Logger.log(msg);
}

function installFit365ReminderDailyTrigger() {
  var fn = "fit365ReminderDailyTrigger";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() !== FIT365_STORE_HUB_SPREADSHEET_ID) return false;
  removeFit365HubReminderTriggers_();
  ScriptApp.newTrigger(fn)
    .forSpreadsheet(ss)
    .timeBased()
    .atHour(9)
    .nearMinute(5)
    .everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  return true;
}
