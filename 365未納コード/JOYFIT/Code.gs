/**
 * JOYFIT未納金回収メール自動集計スクリプト (Ver 5.6) ＋ Excel(CSV)取り込み機能統合版
 *
 * ローカル配置: Code.gs と dialog.html を JOYFIT 用スプレッドシートの Apps Script にコピーして使う。
 * Ver 8.9.1: 入金反映完了メッセージを簡略化（今回新規取込分の月・氏名のみ表示）。
 * Ver 8.9: 入金反映—Gmailは先月1日〜翌月1日、対象年月はメール本文の(YYYY年MM月)。先月・当月シートを消込。
 * Ver 8.8: 入金メール反映の高速化—消込・人数突合・月次行リンクをバッチ処理。
 * Ver 8.7: 入金メール反映—実行日の「当月」1タブのみ（→8.9で先月+当月に戻す。6月初旬の5月分取りこぼし対策）。
 * Ver 8.6: 入金メール反映—当月＋前月のみ（→8.7で当月1タブに変更）。
 * Ver 8.5: メニュー整理—日常3項目（CSV取込・入金反映・SMS）＋「その他」にメンテ3項目。
 * Ver 8.4: App入金シートの見た目（見出し・ゼブラ）復元。月次シートの入金済みグレー・取り消し線（B〜L）+V列を復活。
 * Ver 8.3: メニュー「月次シートを原本体裁に整える」—AA〜AI削除・J/K/L数式・区分色(B〜L)・中央揃え（CSV不要）。
 * Ver 8.2: 月次シートはB〜Iのみ書込（中央揃え）。J/Kは数式復元。消込はV列のみ。M〜T・U・W・Xは触らない。
 * Ver 8.1: 月次シートは見た目を変更しない（背景色・枠線・行高・列幅・行削除なし）。AA〜AI列の照合欄廃止。
 * Ver 8.0.3: App入金表示—QUERY廃止（gid誤参照エラー回避）、スクリプトで一覧出力。
 * Ver 8.0.2: App入金シート合計—税込=D列・税抜=E列のSUBTOTAL（F列は月次シートのため0になっていた）。
 * Ver 8.0.1: App入金_Data 書込—メール6列＋月次行等3列は別途反映（列数不一致エラー修正）。
 * Ver 8.0: 月次シート27列目〜にApp入金照合（A行/P区分/P会員番号/B会員名/B電話/Y年月/U税抜/K税込/E経路）。
 *         App入金シートはA=月次行・B=対象年月…、受信日時は非表示。
 * Ver 7.9: 入金消込—App入金にいない行の誤グレーを自動解除。人数突合サマリを表示。
 * Ver 7.8: 【推奨】消込一から見直し—誤グレー・取り消し線を全員分クリア後、App入金で再照合＋ログシート。
 * Ver 7.7: 入金消込—5月シートは5月分のみ照合（F列の複数月表記は使わない）。
 * Ver 7.6: 入金消込—氏名＋対象月一致のみ（他の中村さんへの誤マッチ防止）。
 * Ver 7.5: 入金消込—氏名は空白正規化で照合。App入金済みは全月次シート上の該当行にグレーアウト。
 * Ver 7.4: App入金メール解析を行単位に改善（複数月分を月ごとに正しく集計・合計照合）。
 * Ver 7.3: App入金シート—内訳列削除（5列表示）、UIをモノトーン化。
 * Ver 7.2: 入金消込—背景グレーはB〜U、取り消し線・白文字はB〜Lのみ（W=チェック、X=アプリ入金）。
 * Ver 7.1: CSVの2/1ヶ月から貸倒候補を除外（先月青ゾーンと会員番号・氏名一致、CSV退会者も2/1に載せない）。
 * Ver 7.0: 貸倒＝先月シートの青ゾーンをそのまま転記のみ。2/1ヶ月＝CSVのみ（全過去シート走査・CSV貸倒廃止）。
 * Ver 6.12: CSV取込後、データ行の N〜T・W〜Z にチェックボックスを配置。
 * Ver 6.11: 列ずれ過去シート対応（会員名がC列など）—緩和レイアウト＋データ行から列推定。
 * Ver 6.10: 貸倒の初回月はF列ではなく「青ゾーンに載った最古の月次シート名」で判定。
 * Ver 6.9: 色・枠線などの装飾範囲を A〜Z 列まで拡張（データ入力は従来どおり B〜L）。
 * Ver 6.8: 見た目整理—見出し行・金額書式・A列ラベル中央揃え。
 * Ver 6.7: 3区分それぞれを太枠（SOLID_THICK）で独立に囲んで区切り。
 * Ver 6.6: 貸倒F列は初回貸倒月（全月次シート・F列タグのうち最も古い月）。
 * Ver 6.5: 取込後に見出し＋3区分を太枠で区切り（既定枠線をクリアして再描画）。
 * Ver 6.4: 出力並び—貸倒は○月貸倒を古い月順・CSV貸倒は最後、2/1ヶ月は会員番号順。
 * Ver 6.3: 過去シート貸倒取込時、会員名の取り消し線（入金済み）行を除外。
 * Ver 6.2: G列未納理由を入力規則の7択に正規化（空欄はその他）。
 * Ver 6.1: 列ずれシートはスキップ。B列はJACCS/クレカのみ。見出し不一致シートは貸倒参照しない。
 * Ver 6.0: 貸倒は全「○年○月」シートの青ゾーンを自動バックアップ（参照シート選択不要）。
 * Ver 5.9: 割引後単価で月数算出。F列は当月含む逆算。貸倒はCSV貸倒/○月貸倒。
 * Ver 5.8: 未納月数は max(N÷100,月会費) で算出。
 * Ver 5.6: CSV取込の退会判定・2/1ヶ月区分を CASIO列(N÷100・O累計)とナショナル会員U行基準に変更。
 * （HTML はファイル名を dialog にすること。）
 * FIT365 用の別プロジェクトは同階層の FIT365/ フォルダ。
 */

/**
 * スプレッドシートを開いたときにメニューを追加（シンプルトリガー）
 * ※スクリプトエディタから onOpen を手動実行すると getUi() が使えずエラーになるため、
 *   その場合はスプレッドシートを開き直すか installJoyfitMenu を実行してください。
 */
function onOpen(e) {
  try {
    addJoyfitBillingMenu_();
  } catch (err) {
    Logger.log("onOpen: メニュー追加をスキップ — " + err);
  }
}

/**
 * メニュー「請求データ管理」
 * 日常: CSV取込 → 入金メール反映 → SMS出力（この順が基本）
 * その他: 見た目修復・消込やり直し・App入金一覧だけ更新
 */
function addJoyfitBillingMenu_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("請求データ管理")
    .addItem("未納CSVを取り込む", "showImportDialog")
    .addItem("入金メールを反映（先月・当月）", "updateJoyfitPayments")
    .addSeparator()
    .addItem("SMS用CSVを出力", "downloadSmsCsv")
    .addSubMenu(
      ui
        .createMenu("その他")
        .addItem("月次シートの見た目を整える", "restoreJoyfitMonthlySheetsFromMenu")
        .addItem("入金消込をやり直す", "rebuildJoyfitReconcileFromAppData")
        .addItem("App入金一覧だけ更新", "refreshAppPaymentUiFromMenu")
    )
    .addToUi();
}

/**
 * メニューを手動で再表示したいとき用（対象スプレッドシートを開いた状態で実行）
 */
function installJoyfitMenu() {
  addJoyfitBillingMenu_();
}

/**
 * =========================================================================
 * 【追加機能】 スプレッドシート内のシート名一覧を取得する関数（ダイアログ用）
 * =========================================================================
 */
function getSheetNames() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}

/**
 * ダイアログ初期表示用: シート一覧と現在アクティブシート名を返す
 */
function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    activeSheet: ss.getActiveSheet().getName(),
    allSheets: ss.getSheets().map(sheet => sheet.getName())
  };
}

/**
 * =========================================================================
 * 【追加機能】 ドラッグ＆ドロップ用のダイアログを表示する関数
 * =========================================================================
 */
function showImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('dialog')
      .setWidth(450)
      .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '未納者データの取り込み');
}

/**
 * 見出し比較用の正規化（改行/全角空白を吸収）
 */
function normalizeJoyfitHeaderLabel_(v) {
  return String(v || "")
    .replace(/^\uFEFF/, "")
    .replace(/\r\n|\n|\r/g, " ")
    .replace(/[ \u3000]+/g, " ")
    .trim();
}

function buildJoyfitHeaderColumnMap_(rowVals) {
  var map = {};
  for (var c = 0; c < rowVals.length; c++) {
    var k = normalizeJoyfitHeaderLabel_(rowVals[c]);
    if (!k || map.hasOwnProperty(k)) continue;
    map[k] = c + 1;
  }
  return map;
}

/**
 * 見出し行から「JACCS / クレカ」列（支払区分）を推定
 */
function findJoyfitPayTypeColumnFromRowVals_(rowVals) {
  for (var c = 0; c < rowVals.length; c++) {
    var k = normalizeJoyfitHeaderLabel_(rowVals[c]);
    if (!k) continue;
    var low = k.toLowerCase();
    if (low.indexOf("jaccs") !== -1 || k.indexOf("ジャックス") !== -1) return c + 1;
    if (k.indexOf("クレカ") !== -1) return c + 1;
  }
  return 0;
}

/** 書き込み直前に B列・G列を入力規則に合わせる */
function sanitizeOutputRowsForValidation_(rows) {
  for (var wi = 0; wi < rows.length; wi++) {
    rows[wi][0] = normalizePayTypeForOutput_(rows[wi][0]);
    rows[wi][5] = normalizeUnpaidReasonForOutput_(rows[wi][5]);
  }
}

/** B列入力規則用: JACCS または クレカ のみ返す */
function normalizePayTypeForOutput_(raw) {
  var v = String(raw || "").replace(/\s+/g, "").trim();
  if (!v) return "クレカ";
  var low = v.toLowerCase();
  if (v === "JACCS" || low === "jaccs" || v.indexOf("ジャックス") !== -1) return "JACCS";
  return "クレカ";
}

function looksLikeMemberId_(raw) {
  var s = String(raw || "").replace(/^'/, "").trim();
  return /^\d{8,12}$/.test(s);
}

function looksLikeMemberName_(raw) {
  var s = String(raw || "").trim();
  if (!s || s === "会員名" || s === "氏名") return false;
  if (parseJoyfitSheetYearMonth_(s)) return false;
  if (/^\d+$/.test(s.replace(/\s/g, ""))) return false;
  if (
    s === "その他" ||
    s.indexOf("使用不可") !== -1 ||
    s.indexOf("有効期") !== -1 ||
    s.indexOf("依頼書") !== -1 ||
    s.indexOf("限度") !== -1
  ) {
    return false;
  }
  if (s.indexOf(" ") === -1 && s.indexOf("　") === -1 && s.length < 4) return false;
  return true;
}

/** 氏名の重複判定用（空白除去） */
function normalizeMemberNameKey_(name) {
  return String(name || "").replace(/\s+/g, "").trim();
}

/** 会員番号の昇順比較（数値IDは数値順、それ以外は文字列順） */
function compareMemberId_(idA, idB) {
  var a = String(idA || "").replace(/^'/, "").trim();
  var b = String(idB || "").replace(/^'/, "").trim();
  if (a === b) return 0;
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    var na = parseInt(a, 10);
    var nb = parseInt(b, 10);
    if (na !== nb) return na - nb;
  }
  return a < b ? -1 : 1;
}

/** F列「26年4月貸倒」などから月ソートキー（CSV貸倒は最大＝末尾） */
function parseBadDebtMonthTagSortKey_(unpaidMonthLabel, badDebtSource) {
  if (badDebtSource === "csv") return 999999;
  var s = String(unpaidMonthLabel || "").trim();
  if (!s || s === "CSV貸倒" || s.indexOf("CSV") !== -1) return 999999;
  var ym = parseJoyfitSheetYearMonth_(s.replace(/貸倒/g, "").trim());
  if (ym) return ym.yy * 12 + ym.mm;
  var m = s.match(/(\d+)\s*年\s*(\d+)\s*月/);
  if (m) return parseInt(m[1], 10) * 12 + parseInt(m[2], 10);
  return 0;
}

/**
 * 取込結果の並び替え
 * - 退会後未納貸倒: F列の月タグ順（古い月→新しい月）、同一月内は会員番号順
 * - 2ヶ月未納・1ヶ月未納: 各グループ内で会員番号順
 */
function sortUnpaidMembersForOutput_(membersArray) {
  membersArray.sort(function(a, b) {
    if (a.group !== b.group) return a.group - b.group;

    if (a.group === 1) {
      var keyA = parseBadDebtMonthTagSortKey_(a.unpaidMonthLabel, a.badDebtSource);
      var keyB = parseBadDebtMonthTagSortKey_(b.unpaidMonthLabel, b.badDebtSource);
      if (keyA !== keyB) return keyA - keyB;
      return compareMemberId_(a.id, b.id);
    }

    return compareMemberId_(a.id, b.id);
  });
}

/**
 * CSV取込で自動装飾する範囲（列・行）
 * - 列: A(区分ラベル) 〜 Z 列（色・枠線・見出し体裁）
 * - 値の書き込み: B〜I（テキスト）のみ。J/Kは数式、L以降は触らない
 * - 行: 見出し行(headerRow) 〜 データ最終行（最低65行目まで枠線クリア対象）
 */
var JOYFIT_UNPAID_STYLE_LAST_COL_ = 26;
var JOYFIT_UNPAID_BORDER_LAST_COL_ = 26;
var JOYFIT_UNPAID_END_ROW_MIN_ = 65;

/** CSV取込で値を書き込む列（B=2 〜 I=9） */
var JOYFIT_UNPAID_DATA_COL_START_ = 2;
var JOYFIT_UNPAID_DATA_COL_COUNT_ = 8;
/** J=支払額(=H+I)、K=手数料(IFS) — 原本の数式 */
var JOYFIT_COL_PAYMENT_ = 10;
var JOYFIT_COL_FEE_ = 11;
var JOYFIT_COL_TOTAL_ = 12;

/** 区分ごとの背景色（交互のゼブラ） */
var JOYFIT_GROUP_COLORS_ = {
  badDebt: ["#c9daf8", "#a4c2f4"],
  twoMonth: ["#fff2cc", "#ffe599"],
  oneMonth: ["#d9ead3", "#b6d7a8"],
  labelBadDebt: "#a4c2f4",
  labelTwoMonth: "#ffe599",
  labelOneMonth: "#b6d7a8",
  header: "#e8eaed"
};

function joyfitGroupRowBgColor_(group, indexInGroup) {
  var pal =
    group === 1
      ? JOYFIT_GROUP_COLORS_.badDebt
      : group === 2
        ? JOYFIT_GROUP_COLORS_.twoMonth
        : JOYFIT_GROUP_COLORS_.oneMonth;
  return pal[(indexInGroup - 1) % 2];
}

/** 範囲を太枠で囲む（内側の格子線は付けない） */
function applyJoyfitThickBoxBorder_(range, borderStyle) {
  var black = "#000000";
  var style = borderStyle || SpreadsheetApp.BorderStyle.SOLID_THICK;
  range.setBorder(true, true, true, true, false, false, black, style);
}

/**
 * 見出し行＋3区分をそれぞれ太枠で区切る（既定枠線はクリアしてから描画）
 * @param {number} headerRow 見出し行（例: 2）
 * @param {number} dataStartRow データ先頭行（例: 3）
 * @param {{1:number,2:number,3:number}} groupCounts 各区分の行数
 */
function applyJoyfitUnpaidSectionBorders_(sheet, headerRow, dataStartRow, groupCounts, lastCol) {
  lastCol = lastCol || JOYFIT_UNPAID_BORDER_LAST_COL_;
  var g1 = groupCounts[1] || 0;
  var g2 = groupCounts[2] || 0;
  var g3 = groupCounts[3] || 0;
  var dataRows = g1 + g2 + g3;
  if (dataRows <= 0 && !headerRow) return;

  headerRow = parseInt(headerRow, 10) || dataStartRow - 1;
  dataStartRow = parseInt(dataStartRow, 10) || headerRow + 1;
  var bottomRow = dataStartRow + dataRows - 1;
  var topRow = headerRow > 0 ? headerRow : dataStartRow;
  var clearBottom = Math.max(bottomRow, JOYFIT_UNPAID_END_ROW_MIN_, sheet.getLastRow());

  var outerThick = SpreadsheetApp.BorderStyle.SOLID_THICK;
  var headerThick = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  var thin = SpreadsheetApp.BorderStyle.SOLID;
  var gray = "#666666";

  var clearRange = sheet.getRange(topRow, 1, clearBottom - topRow + 1, lastCol);
  clearRange.setBorder(false, false, false, false, false, false);

  if (headerRow > 0 && headerRow < dataStartRow) {
    var headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
    applyJoyfitThickBoxBorder_(headerRange, headerThick);
    if (lastCol >= 2) {
      sheet
        .getRange(headerRow, 2, 1, lastCol - 1)
        .setBorder(null, null, null, null, true, false, gray, thin);
    }
  }

  var cursor = dataStartRow;
  var groupSizes = [g1, g2, g3];
  for (var gi = 0; gi < groupSizes.length; gi++) {
    var cnt = groupSizes[gi];
    if (cnt <= 0) continue;
    var block = sheet.getRange(cursor, 1, cnt, lastCol);
    applyJoyfitThickBoxBorder_(block, outerThick);
    if (lastCol >= 2 && cnt > 0) {
      sheet
        .getRange(cursor, 2, cnt, lastCol - 1)
        .setBorder(null, null, null, null, true, true, gray, thin);
    }
    cursor += cnt;
  }
}

/** 見出し行（A〜Z）の体裁 */
function applyJoyfitUnpaidHeaderStyle_(sheet, headerRow, lastCol) {
  lastCol = lastCol || JOYFIT_UNPAID_STYLE_LAST_COL_;
  if (!headerRow || headerRow < 1) return;
  var headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
  headerRange
    .setBackground(JOYFIT_GROUP_COLORS_.header)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setRowHeight(headerRow, 32);
}

/** データ行 A〜Z の体裁（金額書式は H〜L のみ） */
function applyJoyfitUnpaidDataStyle_(sheet, dataStartRow, rowCount, lastCol) {
  lastCol = lastCol || JOYFIT_UNPAID_STYLE_LAST_COL_;
  if (rowCount <= 0) return;
  sheet.getRange(dataStartRow, 1, rowCount, lastCol).setVerticalAlignment("middle").setFontSize(10);
  sheet.getRange(dataStartRow, 8, rowCount, 4).setNumberFormat("¥#,##0"); // H〜K
  sheet.getRange(dataStartRow, 12, rowCount, 1).setNumberFormat("¥#,##0"); // L
  for (var r = 0; r < rowCount; r++) {
    sheet.setRowHeight(dataStartRow + r, 22);
  }
}

/** データ行の N〜T・W〜Z にチェックボックスを設置（未チェックで初期化） */
function applyJoyfitUnpaidCheckboxColumns_(sheet, dataStartRow, rowCount) {
  if (!sheet || rowCount <= 0 || dataStartRow < 1) return;

  var rangeNT = sheet.getRange(
    dataStartRow,
    JOYFIT_CHECKBOX_COL_NT_START_,
    rowCount,
    JOYFIT_CHECKBOX_COL_NT_COUNT_
  );
  var rangeWZ = sheet.getRange(
    dataStartRow,
    JOYFIT_CHECKBOX_COL_WZ_START_,
    rowCount,
    JOYFIT_CHECKBOX_COL_WZ_COUNT_
  );

  rangeNT.clearDataValidations();
  rangeWZ.clearDataValidations();
  rangeNT.clearContent();
  rangeWZ.clearContent();
  rangeNT.insertCheckboxes();
  rangeWZ.insertCheckboxes();
  rangeNT.setHorizontalAlignment("center").setVerticalAlignment("middle");
  rangeWZ.setHorizontalAlignment("center").setVerticalAlignment("middle");
}

/** A列の区分ラベル（結合セル）の体裁 */
function applyJoyfitGroupLabelStyle_(labelRange, bgColor) {
  labelRange
    .setBackground(bgColor)
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

/**
 * 見出しからレイアウトを解決（見出しが揃わないシートは null＝スキップ）
 * 必須: 会員番号・電話番号・支払区分(クレカ/JACCS)の見出し
 */
function tryResolveJoyfitSheetLayout_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var scanLastRow = Math.min(sheet.getLastRow(), 10);
  for (var r = 1; r <= scanLastRow; r++) {
    var rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var map = buildJoyfitHeaderColumnMap_(rowVals);
    var idCol = map[normalizeJoyfitHeaderLabel_("会員番号")] || 0;
    var phoneCol = map[normalizeJoyfitHeaderLabel_("電話番号")] || 0;
    var payCol = findJoyfitPayTypeColumnFromRowVals_(rowVals);
    if (!idCol || !phoneCol || !payCol) continue;
    var nameCol =
      map[normalizeJoyfitHeaderLabel_("会員名")] ||
      map[normalizeJoyfitHeaderLabel_("氏名")] ||
      0;
    return {
      headerRow: r,
      dataStartRow: r + 1,
      colGroup: 1,
      colPayType: payCol,
      colMemberId: idCol,
      colMemberName: nameCol || Math.min(idCol + 1, lastCol),
      colPhone: phoneCol
    };
  }
  return null;
}

function pickBestScoredColumn_(scores, minScore) {
  minScore = minScore || 2;
  var best = 0;
  var bestK = 0;
  for (var k in scores) {
    if (!scores.hasOwnProperty(k)) continue;
    var n = scores[k];
    if (n > best) {
      best = n;
      bestK = parseInt(k, 10);
    }
  }
  return best >= minScore ? bestK : 0;
}

function looksLikePhoneCell_(raw) {
  var d = String(raw || "").replace(/\D/g, "");
  return d.length >= 9 && d.length <= 11;
}

/**
 * データ行を走査して会員番号・氏名・電話の列を推定（列ずれシート用）
 */
function inferJoyfitMemberColumnsFromData_(sheet, dataStartRow, lastCol, hintIdCol, hintNameCol, hintPhoneCol) {
  var lastRow = Math.min(sheet.getLastRow(), dataStartRow + 100);
  if (lastRow < dataStartRow) return null;

  var numRows = lastRow - dataStartRow + 1;
  var data = sheet.getRange(dataStartRow, 1, numRows, lastCol).getValues();
  var maxCol = Math.min(lastCol, 15);
  var idScores = {};
  var nameScores = {};
  var phoneScores = {};

  for (var ri = 0; ri < data.length; ri++) {
    for (var c = 1; c <= maxCol; c++) {
      var s = String(data[ri][c - 1] || "").replace(/^'/, "").trim();
      if (!s) continue;
      if (looksLikeMemberId_(s)) idScores[c] = (idScores[c] || 0) + 1;
      if (looksLikeMemberName_(s)) nameScores[c] = (nameScores[c] || 0) + 1;
      if (looksLikePhoneCell_(s)) phoneScores[c] = (phoneScores[c] || 0) + 1;
    }
  }

  var idCol = hintIdCol || pickBestScoredColumn_(idScores, 2);
  var nameCol = hintNameCol || pickBestScoredColumn_(nameScores, 2);
  var phoneCol = hintPhoneCol || pickBestScoredColumn_(phoneScores, 2);

  if (nameCol && !idCol) idCol = nameCol + 1;
  if (idCol && !nameCol) nameCol = idCol - 1;
  if (idCol === nameCol) {
    if (nameScores[3] && idScores[4]) {
      nameCol = 3;
      idCol = 4;
    } else if (nameScores[3] && idScores[2]) {
      nameCol = 3;
      idCol = 2;
    } else {
      return null;
    }
  }
  if (!idCol || !nameCol) return null;

  if (!phoneCol) {
    phoneCol = nameCol + 1;
    if (phoneCol === idCol) phoneCol = idCol + 1;
  }

  var payCol = findJoyfitPayTypeColumnFromRowVals_(
    sheet.getRange(dataStartRow - 1, 1, 1, lastCol).getValues()[0]
  );
  if (!payCol) payCol = Math.min(2, nameCol - 1);

  return {
    idCol: idCol,
    nameCol: nameCol,
    phoneCol: phoneCol,
    payCol: payCol
  };
}

/** 過去月次シート用：見出しが一部欠け・列ずれでもレイアウト推定 */
function tryResolveJoyfitSheetLayoutLenient_(sheet) {
  var strict = tryResolveJoyfitSheetLayout_(sheet);
  if (strict) return strict;

  if (!parseJoyfitSheetYearMonth_(sheet.getName())) return null;

  var lastCol = Math.max(sheet.getLastColumn(), 15);
  var scanLastRow = Math.min(sheet.getLastRow(), 15);

  for (var r = 1; r <= scanLastRow; r++) {
    var rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var map = buildJoyfitHeaderColumnMap_(rowVals);
    var idCol = map[normalizeJoyfitHeaderLabel_("会員番号")] || 0;
    var nameCol =
      map[normalizeJoyfitHeaderLabel_("会員名")] ||
      map[normalizeJoyfitHeaderLabel_("氏名")] ||
      0;
    var phoneCol = map[normalizeJoyfitHeaderLabel_("電話番号")] || 0;
    var payCol = findJoyfitPayTypeColumnFromRowVals_(rowVals);

    for (var c = 0; c < rowVals.length; c++) {
      var k = normalizeJoyfitHeaderLabel_(rowVals[c]);
      if (!k) continue;
      if (!nameCol && (k.indexOf("会員名") !== -1 || k.indexOf("氏名") !== -1)) nameCol = c + 1;
      if (!idCol && k.indexOf("会員番号") !== -1) idCol = c + 1;
      if (!phoneCol && k.indexOf("電話") !== -1) phoneCol = c + 1;
    }

    if (!idCol && !nameCol) continue;

    var dataStartRow = r + 1;
    var inferred = inferJoyfitMemberColumnsFromData_(sheet, dataStartRow, lastCol, idCol, nameCol, phoneCol);
    if (!inferred) continue;

    return {
      headerRow: r,
      dataStartRow: dataStartRow,
      colGroup: 1,
      colPayType: payCol || inferred.payCol,
      colMemberId: inferred.idCol,
      colMemberName: inferred.nameCol,
      colPhone: inferred.phoneCol,
      lenient: true
    };
  }

  var fallback = inferJoyfitMemberColumnsFromData_(sheet, 3, lastCol, 0, 3, 0);
  if (!fallback) return null;

  return {
    headerRow: 2,
    dataStartRow: 3,
    colGroup: 1,
    colPayType: fallback.payCol,
    colMemberId: fallback.idCol,
    colMemberName: fallback.nameCol,
    colPhone: fallback.phoneCol,
    lenient: true
  };
}

/** 過去シート走査用（厳密→緩和の順で解決） */
function resolveJoyfitSheetLayoutForScan_(sheet) {
  return tryResolveJoyfitSheetLayout_(sheet) || tryResolveJoyfitSheetLayoutLenient_(sheet);
}

/** 行内の近傍列から会員番号・氏名を取得（1列ズレ対策） */
function extractMemberIdFromRow_(row, preferredCol0) {
  var offsets = [0, 1, -1, 2, -2];
  for (var oi = 0; oi < offsets.length; oi++) {
    var c = preferredCol0 + offsets[oi];
    if (c < 0 || c >= row.length) continue;
    var s = String(row[c] || "").replace(/^'/, "").trim();
    if (looksLikeMemberId_(s)) return s;
  }
  return "";
}

function extractMemberNameFromRow_(row, preferredCol0) {
  var offsets = [0, 1, -1, 2, -2];
  for (var oi = 0; oi < offsets.length; oi++) {
    var c = preferredCol0 + offsets[oi];
    if (c < 0 || c >= row.length) continue;
    var s = String(row[c] || "").trim();
    if (looksLikeMemberName_(s)) return s;
  }
  return "";
}

/** 金額列（見出し優先、なければ会員番号列からの相対位置） */
function resolveJoyfitFinancialColumnMap_(headerMap, layout) {
  var id = layout.colMemberId;
  var shift = layout.lenient ? 3 : 4;
  return {
    reason: headerMap[normalizeJoyfitHeaderLabel_("未納理由")] || id + shift,
    carry: headerMap[normalizeJoyfitHeaderLabel_("繰越額")] || id + shift + 1,
    monthly: headerMap[normalizeJoyfitHeaderLabel_("当月分")] || id + shift + 2,
    payment: headerMap[normalizeJoyfitHeaderLabel_("支払額")] || id + shift + 3
  };
}

/** 書き込み先など、見出し不明時のみ旧固定列にフォールバック */
function resolveJoyfitSheetLayout_(sheet) {
  var layout = tryResolveJoyfitSheetLayout_(sheet);
  if (layout) return layout;
  return {
    headerRow: 2,
    dataStartRow: 3,
    colGroup: 1,
    colPayType: 2,
    colMemberId: 3,
    colMemberName: 4,
    colPhone: 5
  };
}

/**
 * 過去シートから会員番号 -> 電話番号を収集（見出し解決優先、失敗時は旧固定）
 */
function buildJoyfitHistoricalPhoneMap_(ss, excludeSheetName) {
  var phoneMap = new Map();
  ss.getSheets().forEach(function(s) {
    var sName = s.getName();
    if (sName === excludeSheetName || !sName.includes("年") || !sName.includes("月")) return;
    var layout = resolveJoyfitSheetLayoutForScan_(s);
    if (!layout) return;
    var data = s.getDataRange().getValues();
    var idIdx = layout.colMemberId - 1;
    var phoneIdx = layout.colPhone - 1;
    var start0 = layout.dataStartRow - 1;
    for (var r = start0; r < data.length; r++) {
      var idNew = String(data[r][idIdx] || "").replace(/^'/, "").trim();
      var idOld = String(data[r][1] || "").replace(/^'/, "").trim();
      var phoneNew = String(data[r][phoneIdx] || "").replace(/^'/, "").trim();
      var phoneOld = String(data[r][3] || "").replace(/^'/, "").trim();
      var memberId = idNew || idOld;
      var rawPhone = phoneNew || phoneOld;
      if (!memberId || !rawPhone) continue;
      var cleanPhone = rawPhone.replace(/\D/g, "");
      if (cleanPhone.length >= 9) {
        if (cleanPhone.length === 10 && (cleanPhone.startsWith("9") || cleanPhone.startsWith("8") || cleanPhone.startsWith("7"))) {
          cleanPhone = "0" + cleanPhone;
        } else if (cleanPhone.length === 9 && !cleanPhone.startsWith("0")) {
          cleanPhone = "0" + cleanPhone;
        }
        phoneMap.set(memberId, cleanPhone);
      }
    }
  });
  return phoneMap;
}

/**
 * CASIO CSV: ナショナル会員U（本体月会費）行か（オプション行は除外）
 */
function isCasioNationalMemberURow_(row) {
  var cat = String(row[5] || "").trim();
  var contract = String(row[6] || "").trim();
  if (cat.indexOf("オプション") !== -1) return false;
  return contract.indexOf("ナショナル会員U") !== -1;
}

/**
 * CASIO CSV: 本体契約行（ナショナル会員Uが無い会員向けフォールバック）
 */
function isCasioMainMembershipRow_(row) {
  if (isCasioNationalMemberURow_(row)) return true;
  var cat = String(row[5] || "").trim();
  if (cat.indexOf("オプション") !== -1) return false;
  return cat.indexOf("フィットネス") !== -1 || cat.indexOf("法人個人") !== -1;
}

/** 当月売上(N列相当): 表示は末尾00多め → ÷100 が1ヶ月分の基準額 */
function casioSalesToMonthlyBase_(salesRaw) {
  var n = parseInt(salesRaw, 10) || 0;
  if (n <= 0) return 0;
  return Math.round(n / 100);
}

/** 本体行の1ヶ月基準額: 月会費+割引（例:9350-4675=4675）。無効時はN÷100と月会費の大きい方 */
function casioNetMonthlyFee_(monthlyFeeRaw, discountRaw, salesRaw) {
  var fee = parseInt(monthlyFeeRaw, 10) || 0;
  var disc = parseInt(discountRaw, 10) || 0;
  var net = fee + disc;
  if (net > 0) return net;
  var fromSales = casioSalesToMonthlyBase_(salesRaw);
  return Math.max(fromSales, fee);
}

/** G列入力規則で許可される未納理由（スプレッドシートのプルダウンと一致させる） */
var JOYFIT_ALLOWED_UNPAID_REASONS_ = [
  "使用不可",
  "取扱不可",
  "有効期限切れ",
  "利用限度額超過",
  "資金不足",
  "依頼書なし",
  "その他"
];

/** CSV・過去シートの文言を G列用の未納理由に正規化（空・不明は「その他」） */
function normalizeUnpaidReasonForOutput_(rawReason) {
  var raw = String(rawReason || "").replace(/\r\n|\n|\r/g, " ").trim();
  if (!raw) return "その他";

  var i;
  for (i = 0; i < JOYFIT_ALLOWED_UNPAID_REASONS_.length; i++) {
    if (raw === JOYFIT_ALLOWED_UNPAID_REASONS_[i]) return JOYFIT_ALLOWED_UNPAID_REASONS_[i];
  }

  if (raw.indexOf("使用不可") !== -1) return "使用不可";
  if (raw.indexOf("取扱不可") !== -1 || raw.indexOf("取扱不") !== -1) return "取扱不可";
  if (raw.indexOf("有効期限") !== -1 || raw.indexOf("有効期") !== -1) return "有効期限切れ";
  if (raw.indexOf("限度額") !== -1 || raw.indexOf("限度超過") !== -1) return "利用限度額超過";
  if (raw.indexOf("資金不足") !== -1) return "資金不足";
  if (raw.indexOf("依頼書") !== -1) return "依頼書なし";

  return "その他";
}

function normalizeUnpaidReasonFromCsv_(rawReason) {
  return normalizeUnpaidReasonForOutput_(rawReason);
}

/**
 * 本体行の O÷(N÷100) で未納月数を推定（四捨五入、最低1）
 */
function estimateUnpaidMonths_(mainBillingO, mainSalesBase, totalO) {
  var o = parseInt(mainBillingO, 10) || 0;
  var base = parseInt(mainSalesBase, 10) || 0;
  if (base > 0 && o > 0) {
    return Math.max(1, Math.round(o / base));
  }
  var total = parseInt(totalO, 10) || 0;
  if (base > 0 && total > 0) {
    return Math.max(1, Math.round(total / base));
  }
  return total > 0 ? 1 : 0;
}

/** シート名「26年5月」などから年月を取得 */
function parseJoyfitSheetYearMonth_(sheetName) {
  var m = String(sheetName || "").match(/(\d{1,2})\s*年\s*(\d{1,2})\s*月/);
  if (!m) return null;
  return { yy: parseInt(m[1], 10), mm: parseInt(m[2], 10) };
}

function shiftJoyfitYearMonth_(yy, mm, deltaMonths) {
  var total = yy * 12 + (mm - 1) + deltaMonths;
  var ny = Math.floor(total / 12);
  var nm = (total % 12) + 1;
  return { yy: ny, mm: nm };
}

/**
 * 未納対象月ラベル（書き込み先シート月＝当月を含めてさかのぼる）
 * 例: 26年5月シート・2か月 → 26年4・5月 / 1か月 → 26年5月
 */
function formatJoyfitMonthRangeLabel_(monthList) {
  var byYear = {};
  for (var i = 0; i < monthList.length; i++) {
    var y = monthList[i].yy;
    if (!byYear[y]) byYear[y] = [];
    byYear[y].push(monthList[i].mm);
  }
  var years = Object.keys(byYear).map(function(k) { return parseInt(k, 10); }).sort(function(a, b) { return a - b; });
  var parts = [];
  for (var yi = 0; yi < years.length; yi++) {
    var y = years[yi];
    var ms = byYear[y];
    if (ms.length === 1) parts.push(y + "年" + ms[0] + "月");
    else parts.push(y + "年" + ms.join("・") + "月");
  }
  return parts.join("、");
}

function buildUnpaidTargetMonthLabel_(anchorYy, anchorMm, unpaidMonths) {
  var n = Math.max(1, parseInt(unpaidMonths, 10) || 1);
  var list = [];
  for (var k = n - 1; k >= 0; k--) {
    list.push(shiftJoyfitYearMonth_(anchorYy, anchorMm, -k));
  }
  return formatJoyfitMonthRangeLabel_(list);
}

/** 2ヶ月未納・1ヶ月未納のF列（貸倒は別タグ） */
function buildUnpaidTargetMonthLabelForGroup_(anchorYy, anchorMm, group, calculatedMonths) {
  if (group === 1) return "";
  if (group === 2) {
    var m2 = Math.max(2, parseInt(calculatedMonths, 10) || 2);
    return buildUnpaidTargetMonthLabel_(anchorYy, anchorMm, m2);
  }
  return buildUnpaidTargetMonthLabel_(anchorYy, anchorMm, 1);
}

function buildPastBadDebtMonthTag_(referenceSheetName) {
  var ym = parseJoyfitSheetYearMonth_(referenceSheetName);
  if (ym) return ym.yy + "年" + ym.mm + "月貸倒";
  return String(referenceSheetName || "").trim() + "貸倒";
}

/**
 * F列「25年9月貸倒」などからタグ文字列を取得（該当しなければ空）
 * ※貸倒の初回月判定には使わない（F列は前月からの繰越表示のため、全員同じ月に偏る）
 */
function parsePastBadDebtTagFromFCell_(fCellValue) {
  var s = String(fCellValue || "").trim();
  if (!s || s === "CSV貸倒" || s.indexOf("CSV") !== -1) return "";
  var m = s.match(/(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*貸倒/);
  if (m) return parseInt(m[1], 10) + "年" + parseInt(m[2], 10) + "月貸倒";
  return "";
}

/** 「○年○月貸倒」→ ソート用キー（yy*12+mm） */
function badDebtMonthTagToSortKey_(tag) {
  var ym = parseJoyfitSheetYearMonth_(String(tag || "").replace(/貸倒/g, "").trim());
  if (ym) return ym.yy * 12 + ym.mm;
  return 999999;
}

/** 2つの貸倒タグのうち、より古い（初回）方を返す */
function pickEarlierBadDebtMonthTag_(tagA, tagB) {
  var a = String(tagA || "").trim();
  var b = String(tagB || "").trim();
  if (!a) return b;
  if (!b) return a;
  return badDebtMonthTagToSortKey_(a) <= badDebtMonthTagToSortKey_(b) ? a : b;
}

function isBadDebtGroupLabel_(groupLabel) {
  var g = String(groupLabel || "").replace(/\r\n|\n|\r/g, " ").trim();
  if (!g) return false;
  return (
    g.indexOf("退会後未納貸倒候補") !== -1 ||
    (g.indexOf("退会後未納") !== -1 && g.indexOf("貸倒候補") !== -1)
  );
}

/** 会員名列に取り消し線が付いている = 入金済み（手動・レジ入金・アプリ入金の照合後） */
function isJoyfitNameCellPaidByStrikethrough_(fontLineValue) {
  var s = String(fontLineValue || "").toLowerCase();
  return s.indexOf("line-through") !== -1;
}

/** データ行の会員名列の FontLine を一括取得（dataStartRow から） */
function readJoyfitMemberNameFontLines_(sheet, layout, lastRow) {
  if (!sheet || !layout || lastRow < layout.dataStartRow) return [];
  var numRows = lastRow - layout.dataStartRow + 1;
  return sheet.getRange(layout.dataStartRow, layout.colMemberName, numRows, 1).getFontLines();
}

/** 対象月次シートでアプリ入金済みの氏名 */
function getPaidMemberNamesForSheet_(ss, targetSheetName) {
  var paid = new Set();
  var dataSheet = ss.getSheetByName("App入金_Data");
  if (!dataSheet || dataSheet.getLastRow() < 2) return paid;
  var values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues();
  for (var vi = 0; vi < values.length; vi++) {
    var d = values[vi][0];
    var name = values[vi][2];
    if (d instanceof Date && name) {
      var y = String(d.getFullYear()).slice(-2);
      var m = String(d.getMonth() + 1);
      if (y + "年" + m + "月" === targetSheetName) paid.add(String(name).trim());
    }
  }
  return paid;
}

function parseNumCell_(v) {
  if (v === "" || v == null) return 0;
  if (typeof v === "number") return v;
  return parseInt(String(v).replace(/[^\d-]/g, ""), 10) || 0;
}

/** 書き込み先シートのひとつ前の「○年○月」シート名 */
function getPreviousMonthSheetName_(targetSheetName) {
  var ym = parseJoyfitSheetYearMonth_(targetSheetName);
  if (!ym) return "";
  var prev = shiftJoyfitYearMonth_(ym.yy, ym.mm, -1);
  return prev.yy + "年" + prev.mm + "月";
}

/**
 * 月次シートの青ゾーン（退会後未納貸倒候補）行を取得
 * @param {boolean} copyFColumnAsIs true=先月転記（F列の値をそのまま使用）
 */
function collectBadDebtRowsFromSheet_(ss, referenceSheetName, targetSheetName, phoneMap, copyFColumnAsIs) {
  var list = [];
  referenceSheetName = String(referenceSheetName || "").trim();
  if (!referenceSheetName) return list;

  var refSheet = ss.getSheetByName(referenceSheetName);
  if (!refSheet) return list;

  var layout = resolveJoyfitSheetLayoutForScan_(refSheet);
  if (!layout) return list;

  var fallbackTag = buildPastBadDebtMonthTag_(referenceSheetName);
  var paidNamesOnRef = getPaidMemberNamesForSheet_(ss, referenceSheetName);
  var paidNamesOnTarget = getPaidMemberNamesForSheet_(ss, targetSheetName);
  var lastRow = refSheet.getLastRow();
  if (lastRow < layout.dataStartRow) return list;

  var lastCol = Math.max(refSheet.getLastColumn(), 12);
  var headerMap = buildJoyfitHeaderColumnMap_(refSheet.getRange(layout.headerRow, 1, 1, lastCol).getValues()[0]);
  var colUnpaidMonth = headerMap[normalizeJoyfitHeaderLabel_("未納対象月")] || 6;
  var finCols = resolveJoyfitFinancialColumnMap_(headerMap, layout);
  var colReason = finCols.reason;
  var colCarry = finCols.carry;
  var colMonthly = finCols.monthly;
  var colPayment = finCols.payment;

  var data = refSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var nameFontLines = readJoyfitMemberNameFontLines_(refSheet, layout, lastRow);
  var g0 = layout.colGroup - 1;
  var p0 = layout.colPayType - 1;
  var id0 = layout.colMemberId - 1;
  var n0 = layout.colMemberName - 1;
  var ph0 = layout.colPhone - 1;
  var start0 = layout.dataStartRow - 1;
  var currentGroup = "";

  for (var i = start0; i < data.length; i++) {
    var colA = String(data[i][g0] || "").trim();
    if (colA !== "" && colA !== "合計") currentGroup = colA;
    if (!isBadDebtGroupLabel_(currentGroup)) continue;

    var memberId = extractMemberIdFromRow_(data[i], id0);
    var name = extractMemberNameFromRow_(data[i], n0);
    if (!memberId || !name) continue;

    var fontIdx = i - start0;
    if (
      fontIdx >= 0 &&
      fontIdx < nameFontLines.length &&
      isJoyfitNameCellPaidByStrikethrough_(nameFontLines[fontIdx][0])
    ) {
      continue;
    }
    if (paidNamesOnRef.has(name) || paidNamesOnTarget.has(name)) continue;

    var payType = normalizePayTypeForOutput_(data[i][p0]);
    var phone = String(data[i][ph0] || "").trim() || phoneMap.get(memberId) || "";
    var reason = normalizeUnpaidReasonForOutput_(data[i][colReason - 1]);
    var carry = parseNumCell_(data[i][colCarry - 1]);
    var monthly = parseNumCell_(data[i][colMonthly - 1]);
    var payment = parseNumCell_(data[i][colPayment - 1]);
    if (!payment) payment = carry + monthly;
    var total = payment;

    var unpaidMonthLabel = fallbackTag;
    if (copyFColumnAsIs) {
      var fRaw = String(data[i][colUnpaidMonth - 1] || "").trim();
      if (fRaw) unpaidMonthLabel = fRaw;
    }

    list.push({
      id: memberId,
      name: name,
      reason: reason,
      carry: carry,
      monthly: monthly,
      totalPayment: total,
      group: 1,
      unpaidMonths: 0,
      unpaidMonthLabel: unpaidMonthLabel,
      badDebtSource: copyFColumnAsIs ? "prevMonth" : "past",
      payType: payType,
      phone: phone,
      _refSortKey: badDebtMonthTagToSortKey_(unpaidMonthLabel)
    });
  }
  return list;
}

/** 先月シートの「退会後未納貸倒候補」青ゾーンをそのまま転記用に取得 */
function collectBadDebtFromPreviousMonthSheet_(ss, targetSheetName, phoneMap) {
  var prevName = getPreviousMonthSheetName_(targetSheetName);
  if (!prevName) return { rows: [], prevSheetName: "" };
  var rows = collectBadDebtRowsFromSheet_(ss, prevName, targetSheetName, phoneMap, true);
  return { rows: rows, prevSheetName: prevName };
}

/** @deprecated Ver7.0以降は未使用（全過去シートからの自動貸倒は廃止） */
function collectPastBadDebtMembersFromSheet_(ss, referenceSheetName, targetSheetName, phoneMap) {
  return collectBadDebtRowsFromSheet_(ss, referenceSheetName, targetSheetName, phoneMap, false);
}

/** 見出し名から列番号（1始まり）を取得。無ければ0 */
function findJoyfitColumnByHeader_(sheet, headerLabel) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var scanLastRow = Math.min(sheet.getLastRow(), 10);
  var key = normalizeJoyfitHeaderLabel_(headerLabel);
  for (var r = 1; r <= scanLastRow; r++) {
    var rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var map = buildJoyfitHeaderColumnMap_(rowVals);
    if (map[key]) return map[key];
  }
  return 0;
}

/** シート名の年月でソート用キー（大きいほど新しい月） */
function joyfitSheetYearMonthSortKey_(sheetName) {
  var ym = parseJoyfitSheetYearMonth_(sheetName);
  if (!ym) return 0;
  return ym.yy * 12 + ym.mm;
}

/** スプレッドシート内の月次シート名一覧（書き込み先以外の「○年○月」） */
function listJoyfitMonthlySheetNames_(ss, excludeSheetName) {
  excludeSheetName = String(excludeSheetName || "").trim();
  var names = [];
  ss.getSheets().forEach(function(s) {
    var n = s.getName();
    if (n === excludeSheetName) return;
    if (!parseJoyfitSheetYearMonth_(n)) return;
    if (!resolveJoyfitSheetLayoutForScan_(s)) return;
    names.push(n);
  });
  names.sort(function(a, b) {
    return joyfitSheetYearMonthSortKey_(a) - joyfitSheetYearMonthSortKey_(b);
  });
  return names;
}

/** 全月次シートの青ゾーンから貸倒候補を収集（25年11月〜26年4月などすべて） */
function buildPastBadDebtMapFromAllMonthlySheets_(ss, targetSheetName, phoneMap) {
  var sheetNames = listJoyfitMonthlySheetNames_(ss, targetSheetName);
  return buildPastBadDebtMapFromSheets_(ss, sheetNames, targetSheetName, phoneMap);
}

/** 複数シートの青ゾーンを統合（同一IDは青ゾーンに載った最古の月次シート＝初回貸倒月） */
function buildPastBadDebtMapFromSheets_(ss, referenceSheetNames, targetSheetName, phoneMap) {
  var map = new Map();
  var names = referenceSheetNames || [];
  for (var si = 0; si < names.length; si++) {
    var refName = String(names[si] || "").trim();
    if (!refName) continue;
    var refKey = joyfitSheetYearMonthSortKey_(refName);
    var list = collectPastBadDebtMembersFromSheet_(ss, refName, targetSheetName, phoneMap);
    for (var li = 0; li < list.length; li++) {
      var p = list[li];
      var prev = map.get(p.id);
      var candidateKey = p._refSortKey != null ? p._refSortKey : refKey;
      var prevKey = prev && prev._refSortKey != null ? prev._refSortKey : 999999;
      if (!prev || candidateKey < prevKey) {
        p.unpaidMonthLabel = buildPastBadDebtMonthTag_(refName);
        p._refSortKey = candidateKey;
        map.set(p.id, p);
      }
    }
  }
  return map;
}

/**
 * =========================================================================
 * 【追加機能】 ダイアログからデータを受け取り、既存シートへ書き込む処理
 * =========================================================================
 */
function processUnpaidData(csvData, targetSheetName, targetMonth, referenceSheetName, referenceSheetName2) {
  // referenceSheetName / 2 は未使用（後方互換のため引数のみ残す）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  targetSheetName = String(targetSheetName || "").trim();
  if (!targetSheetName) {
    targetSheetName = ss.getActiveSheet().getName();
  }
  var sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("指定されたシート「" + targetSheetName + "」が見つかりません。");
  }
  var targetLayout = tryResolveJoyfitSheetLayout_(sheet);
  if (!targetLayout) {
    throw new Error(
      "シート「" + targetSheetName + "」の見出し（会員番号・電話番号・クレカ/JACCS）が標準と違うためスキップしました。列ずれシートは取り込み対象外です。"
    );
  }
  var anchor = parseJoyfitSheetYearMonth_(targetSheetName);
  if (!anchor) {
    var now = new Date();
    anchor = {
      yy: parseInt(String(now.getFullYear()).slice(-2), 10),
      mm: now.getMonth() + 1
    };
  }
  if (!targetMonth) {
    targetMonth = anchor.yy + "年" + anchor.mm + "月";
  }

  // 過去シートから会員番号 -> 電話番号を収集（値のみ利用）
  var phoneMap = buildJoyfitHistoricalPhoneMap_(ss, targetSheetName);

  // 貸倒（青）: 先月シートの「退会後未納貸倒候補」行をそのまま転記（半年未納は別管理）
  var prevBadDebt = collectBadDebtFromPreviousMonthSheet_(ss, targetSheetName, phoneMap);
  var badDebtRows = prevBadDebt.rows;
  var badDebtIdSet = new Set();
  var badDebtNameSet = new Set();
  for (var bi = 0; bi < badDebtRows.length; bi++) {
    badDebtIdSet.add(badDebtRows[bi].id);
    badDebtNameSet.add(normalizeMemberNameKey_(badDebtRows[bi].name));
  }

  // 名寄せ: 会員番号単位。CSVは2ヶ月/1ヶ月未納のみ（退会でも貸倒には入れない）
  var members = new Map();
  for (var i = 1; i < csvData.length; i++) {
    var row = csvData[i];
    if (row.length < 15) continue;

    var memberId = String(row[3] || "").replace(/^'/, "").trim();
    var memberName = String(row[4] || "").trim();
    if (!memberId || !memberName) continue;

    var billingO = parseInt(row[14], 10) || 0;
    var netMonthly = casioNetMonthlyFee_(row[10], row[12], row[13]);
    var rawStatus = String(row[15] || "").trim();
    var rawReason = String(row[18] || row[row.length - 1] || "").trim();
    var reason = normalizeUnpaidReasonFromCsv_(rawReason);

    if (!members.has(memberId)) {
      members.set(memberId, {
        name: memberName,
        reason: reason,
        totalO: billingO,
        mainNetMonthly: 0,
        mainBillingO: 0,
        hasNational: false,
        isWithdrawn: false
      });
    } else {
      members.get(memberId).totalO += billingO;
    }

    var m = members.get(memberId);

    if (isCasioNationalMemberURow_(row)) {
      m.hasNational = true;
      m.mainNetMonthly = netMonthly;
      m.mainBillingO = billingO;
      if (rawStatus.indexOf("退会") !== -1) m.isWithdrawn = true;
      if (reason !== "") m.reason = reason;
    } else if (isCasioMainMembershipRow_(row) && !m.hasNational) {
      m.mainNetMonthly = netMonthly;
      m.mainBillingO = billingO;
      if (rawStatus.indexOf("退会") !== -1) m.isWithdrawn = true;
      if (reason !== "") m.reason = reason;
    }
  }

  var membersArray = [];
  for (var di = 0; di < badDebtRows.length; di++) {
    membersArray.push(badDebtRows[di]);
  }

  members.forEach(function(data, id) {
    // 貸倒（先月青ゾーン）と重複させない：会員番号 or 氏名が一致したら2/1ヶ月に載せない
    if (badDebtIdSet.has(id)) return;
    if (badDebtNameSet.has(normalizeMemberNameKey_(data.name))) return;
    // CSV上の退会者は2/1ヶ月に載せない（貸倒は先月転記のみ。半年未満は別管理）
    if (data.isWithdrawn) return;

    var months = estimateUnpaidMonths_(data.mainBillingO, data.mainNetMonthly, data.totalO);
    var group = months >= 2 ? 2 : 3;

    var total = data.totalO;
    var carry = 0;
    var monthly = 0;
    if (group === 2) {
      monthly = Math.floor(total / 2);
      carry = total - monthly;
    } else {
      carry = 0;
      monthly = total;
    }

    var unpaidMonthLabel = buildUnpaidTargetMonthLabelForGroup_(anchor.yy, anchor.mm, group, months);

    membersArray.push({
      id: id,
      name: data.name,
      reason: data.reason,
      carry: carry,
      monthly: monthly,
      totalPayment: total,
      group: group,
      unpaidMonths: months,
      unpaidMonthLabel: unpaidMonthLabel,
      badDebtSource: "csv",
      payType: "クレカ",
      phone: ""
    });
  });

  sortUnpaidMembersForOutput_(membersArray);

  var startRow = targetLayout.dataStartRow;
  var headerRow = targetLayout.headerRow;
  var outputData = []; // B〜I の8列のみ（J/Kは数式、L以降は原本のまま）
  var groupCounts = { 1: 0, 2: 0, 3: 0 };

  membersArray.forEach(function(item) {
    groupCounts[item.group]++;
    var matchedPhone = phoneMap.get(item.id) || "";

    var fLabel = item.unpaidMonthLabel || targetMonth;
    if (item.group === 1) {
      fLabel = item.unpaidMonthLabel || "";
    }
    var payType = normalizePayTypeForOutput_(item.payType);
    var phoneOut = item.phone || matchedPhone;
    var reasonOut = normalizeUnpaidReasonForOutput_(item.reason);

    outputData.push([
      payType,            // B: 区分（JACCS / クレカ のみ）
      item.id,            // C: 会員番号
      item.name,          // D: 会員名
      phoneOut,           // E: 電話番号
      fLabel,             // F: 未納対象月
      reasonOut,          // G: 未納理由
      item.carry || "",   // H: 繰越額
      item.monthly || ""  // I: 当月分
    ]);
  });

  // B〜I のみクリア・再入力（J/K/L・M以降の数式・チェックボックスは触らない）
  var clearLastRow = Math.max(sheet.getLastRow(), startRow);
  if (clearLastRow >= startRow) {
    var clearHeight = clearLastRow - startRow + 1;
    sheet
      .getRange(startRow, JOYFIT_UNPAID_DATA_COL_START_, clearHeight, JOYFIT_UNPAID_DATA_COL_COUNT_)
      .clearContent();
  }

  if (outputData.length > 0) {
    if (sheet.getMaxRows() < (startRow + outputData.length)) {
      sheet.insertRowsAfter(sheet.getMaxRows(), startRow + outputData.length - sheet.getMaxRows());
    }
    var inputRange = sheet.getRange(
      startRow,
      JOYFIT_UNPAID_DATA_COL_START_,
      outputData.length,
      JOYFIT_UNPAID_DATA_COL_COUNT_
    );
    sanitizeOutputRowsForValidation_(outputData);
    try {
      inputRange.setValues(outputData);
    } catch (writeErr) {
      sanitizeOutputRowsForValidation_(outputData);
      inputRange.setValues(outputData);
    }
    inputRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    applyJoyfitMonthlyRowFormulasJK_(sheet, startRow, outputData.length);
  }

  clearLegacyAppPaymentLogColumnsOnSheet_(sheet, headerRow, startRow);

  // A列の区分ラベル（値のみ・色・枠線は変更しない）
  var aLastRow = Math.max(sheet.getLastRow(), startRow + outputData.length - 1);
  if (aLastRow >= startRow) {
    var aRange = sheet.getRange(startRow, 1, aLastRow - startRow + 1, 1);
    try {
      var mergedRanges = aRange.getMergedRanges();
      for (var m = 0; m < mergedRanges.length; m++) mergedRanges[m].breakApart();
    } catch (e) {}
    aRange.clearContent();
  }

  var cursor = startRow;
  if (groupCounts[1] > 0) {
    var g1 = sheet.getRange(cursor, 1, groupCounts[1], 1);
    g1.setValue("退会後未納貸倒候補");
    if (groupCounts[1] > 1) g1.merge();
    g1.setHorizontalAlignment("center").setVerticalAlignment("middle");
    cursor += groupCounts[1];
  }
  if (groupCounts[2] > 0) {
    var g2 = sheet.getRange(cursor, 1, groupCounts[2], 1);
    g2.setValue("2ヶ月未納");
    if (groupCounts[2] > 1) g2.merge();
    g2.setHorizontalAlignment("center").setVerticalAlignment("middle");
    cursor += groupCounts[2];
  }
  if (groupCounts[3] > 0) {
    var g3 = sheet.getRange(cursor, 1, groupCounts[3], 1);
    g3.setValue("1カ月未納");
    if (groupCounts[3] > 1) g3.merge();
    g3.setHorizontalAlignment("center").setVerticalAlignment("middle");
  }

  var prevName = prevBadDebt.prevSheetName;
  if (prevName && !ss.getSheetByName(prevName)) {
    Logger.log("先月シート「" + prevName + "」が見つからないため、貸倒ゾーンは空です。");
  }

  return outputData.length;
}

/**
 * =========================================================================
 * 📥 新・SMSリストをCSV保存 (ダイアログでグループ選択)
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
          h3 { margin-top: 0; color: #1a73e8; font-size: 16px; }
          p { font-size: 13px; color: #666; }
          label { display: block; font-size: 12px; font-weight: bold; margin: 10px 0 4px; color: #444; }
          select { padding: 8px; width: 100%; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
          button { margin-top: 16px; padding: 10px 20px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 14px; font-weight: bold; }
          button:hover { background-color: #1557b0; }
          #msg { margin-top: 15px; font-size: 13px; text-align: center; }
        </style>
      </head>
      <body>
        <h3>SMSリストの出力</h3>
        <p>未納者詳細シートの見出し（会員名・電話番号など）から列を自動検出します。</p>
        <label for="paySelect">支払区分（B列相当）</label>
        <select id="paySelect">
          <option value="クレカ">クレカ</option>
          <option value="JACCS">JACCS</option>
        </select>
        <label for="segSelect">抽出グループ（A列の区分ラベル）</label>
        <select id="segSelect">
          <option value="退会後未納貸倒候補">退会後未納貸倒候補</option>
          <option value="2ヶ月未納">2ヶ月未納</option>
          <option value="1カ月未納">1カ月未納</option>
          <option value="クレカ全員">クレカ全員（区分のみ一致・グループは問わない）</option>
        </select>
        <button onclick="requestDownload()">CSVを作成してダウンロード</button>
        <div id="msg"></div>
        <script>
          function requestDownload() {
            var pay = document.getElementById('paySelect').value;
            var seg = document.getElementById('segSelect').value;
            document.getElementById('msg').innerText = "データを作成中...";
            document.getElementById('msg').style.color = "#f4b400";
            google.script.run
              .withSuccessHandler(function(csvContent) {
                if (csvContent === "") {
                  document.getElementById('msg').innerText = "条件に合う電話番号がありません。";
                  document.getElementById('msg').style.color = "#d93025";
                  return;
                }
                var safe = (pay + "_" + seg).replace(/\\s+/g, "_");
                safe = safe.split(String.fromCharCode(92)).join("_");
                safe = safe.split("/").join("_").split(":").join("_").split("*").join("_").split("?").join("_");
                safe = safe.split('"').join("_").split("<").join("_").split(">").join("_").split("|").join("_");
                var fileName = "sms_joyfit_" + safe + ".csv";
                var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                var link = document.createElement("a");
                var url = URL.createObjectURL(blob);
                link.setAttribute("href", url);
                link.setAttribute("download", fileName);
                link.style.visibility = "hidden";
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                document.getElementById('msg').innerText = "ダウンロードを開始しました。";
                document.getElementById('msg').style.color = "#0f9d58";
                setTimeout(google.script.host.close, 1500);
              })
              .withFailureHandler(function(error) {
                document.getElementById('msg').innerText = "エラー: " + error.message;
                document.getElementById('msg').style.color = "#d93025";
              })
              .generateSmsCsvContent(pay, seg);
          }
        </script>
      </body>
    </html>
  `).setWidth(380).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SMS送信対象の選択');
}

/**
 * SMS用: 電話番号として十分か（数字10桁以上目安）
 */
function joyfitPhoneLooksValidForSms_(raw) {
  var s = String(raw || "").replace(/^'/, "").trim();
  if (!s || s === "電話番号") return false;
  var d = s.replace(/\D/g, "");
  return d.length >= 10;
}

/**
 * B列の表示値が選択した支払区分と一致するか
 */
function joyfitPayTypeCellMatches_(cellVal, payKind) {
  var v = String(cellVal || "").replace(/\s+/g, "").trim();
  var p = String(payKind || "").trim();
  if (!v || !p) return false;
  if (p === "クレカ") return v === "クレカ" || v.indexOf("クレカ") !== -1;
  if (p === "JACCS") return v === "JACCS" || v.toUpperCase() === "JACCS" || v.indexOf("ジャックス") !== -1;
  return v === p;
}

/**
 * A列の区分ラベルが選択グループに該当するか
 */
function joyfitSmsSegmentMatches_(segment, groupLabel) {
  var g = String(groupLabel || "").replace(/\r\n|\n|\r/g, " ").trim();
  var s = String(segment || "").trim();
  if (!s) return false;
  if (s === "クレカ全員") return true;
  if (s === "退会後未納貸倒候補") {
    return (
      g.indexOf("退会後未納貸倒候補") !== -1 ||
      (g.indexOf("退会後未納") !== -1 && g.indexOf("貸倒候補") !== -1)
    );
  }
  if (s === "2ヶ月未納") return g.indexOf("2ヶ月未納") !== -1 || g.indexOf("2カ月未納") !== -1;
  if (s === "1ヶ月未納" || s === "1カ月未納") {
    return (
      g.indexOf("1ヶ月未納") !== -1 ||
      g.indexOf("1カ月未納") !== -1 ||
      g.indexOf("1ヵ月未納") !== -1
    );
  }
  return false;
}

/**
 * 選択された支払区分・グループのCSV（電話番号1列）を生成
 * @param {string} payKind 「クレカ」または「JACCS」
 * @param {string} segment 退会後未納貸倒候補 / 貸倒候補 / 2ヶ月未納 / 1ヶ月未納 / クレカ全員
 *
 * 後方互換: 第1引数だけ渡され、先頭が JACCS_ の旧形式なら解釈する
 */
function generateSmsCsvContent(payKind, segment) {
  var pk = "クレカ";
  var sg = "クレカ全員";
  if (arguments.length === 1) {
    var one = String(payKind || "");
    if (one.indexOf("JACCS_") === 0) {
      pk = "JACCS";
      sg = one.replace("JACCS_", "");
    } else {
      pk = "クレカ";
      sg = one || "2ヶ月未納";
    }
  } else {
    pk = String(payKind || "クレカ").trim() || "クレカ";
    sg = String(segment || "クレカ全員").trim() || "クレカ全員";
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getActiveSheet();
  var targetSheetName = targetSheet.getName();
  if (!targetSheet) throw new Error("アクティブシートが取得できません。");

  var dataSheetName = "App入金_Data";
  var dataSheet = ss.getSheetByName(dataSheetName);
  var paidNames = new Set();
  if (dataSheet && dataSheet.getLastRow() > 1) {
    var values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues();
    for (var vi = 0; vi < values.length; vi++) {
      var row = values[vi];
      var d = row[0];
      var name = row[2];
      if (d instanceof Date) {
        var y = String(d.getFullYear()).slice(-2);
        var m = String(d.getMonth() + 1);
        var sheetName = y + "年" + m + "月";
        if (sheetName === targetSheetName) paidNames.add(String(name).trim());
      }
    }
  }

  var layout = resolveJoyfitSheetLayout_(targetSheet);
  var lastRow = targetSheet.getLastRow();
  if (lastRow < layout.dataStartRow) return "";

  var maxCol = Math.max(
    layout.colGroup,
    layout.colPayType,
    layout.colMemberId,
    layout.colMemberName,
    layout.colPhone,
    1
  );
  var data = targetSheet.getRange(1, 1, lastRow, maxCol).getValues();

  var g0 = layout.colGroup - 1;
  var p0 = layout.colPayType - 1;
  var n0 = layout.colMemberName - 1;
  var ph0 = layout.colPhone - 1;

  var csvContent = "";
  var currentGroup = "";
  var start0 = layout.dataStartRow - 1;

  for (var i = start0; i < data.length; i++) {
    var colA = String(data[i][g0] || "").trim();
    if (colA !== "" && colA !== "合計") currentGroup = colA;

    var colB = data[i][p0];
    if (!joyfitPayTypeCellMatches_(colB, pk)) continue;

    var name = String(data[i][n0] || "").trim();
    var phone = String(data[i][ph0] || "").trim();
    if (!name || name === "氏名" || name === "会員名") continue;
    if (!joyfitPhoneLooksValidForSms_(phone)) continue;

    if (!joyfitSmsSegmentMatches_(sg, currentGroup)) continue;
    if (paidNames.has(name)) continue;

    var digits = phone.replace(/^'/, "").replace(/\D/g, "");
    csvContent += '"' + digits + '"\r\n';
  }

  return csvContent;
}

/** Date → 照合キー "26-6"（App入金_Data・月次シート共通） */
function joyfitDateToMonthKey_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return "";
  return String(date.getFullYear()).slice(-2) + "-" + String(date.getMonth() + 1);
}

function joyfitFormatGmailQueryDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/M/d");
}

/**
 * 実行日時点の「先月＋当月」
 * 例: 2026/6/6 → 月次 26年5月・26年6月 / Gmail after:2026/5/1 before:2026/7/1
 * 対象年月（App入金_DataのA列・消込先シート）はメール本文の (2026年05月) で決める（受信日ではない）
 */
function joyfitBuildCurrentMonthWindowFromDate_(anchorDate) {
  var d = anchorDate instanceof Date ? anchorDate : new Date();
  var curFull = d.getFullYear();
  var curMm = d.getMonth() + 1;
  var prevFull = curMm === 1 ? curFull - 1 : curFull;
  var prevMm = curMm === 1 ? 12 : curMm - 1;
  var curYy = parseInt(String(curFull).slice(-2), 10);
  var prevYy = parseInt(String(prevFull).slice(-2), 10);

  function buildMonth(yy, mm) {
    var m = {
      yy: yy,
      mm: mm,
      sheetName: joyfitFormatSheetNameFromYm_(yy, mm)
    };
    m.monthKey = joyfitSheetYearMonthKey_(m);
    return m;
  }

  var months = [buildMonth(prevYy, prevMm), buildMonth(curYy, curMm)];
  var monthKeySet = new Set(months.map(function(m) {
    return m.monthKey;
  }));
  var sheetNameSet = new Set(months.map(function(m) {
    return m.sheetName;
  }));

  var gmailAfter = new Date(prevFull, prevMm - 1, 1);
  var gmailBefore = new Date(curFull, curMm, 1);

  return {
    month: months[1],
    months: months,
    monthKeySet: monthKeySet,
    sheetNameSet: sheetNameSet,
    gmailAfter: gmailAfter,
    gmailBefore: gmailBefore,
    gmailAfterLabel: joyfitFormatGmailQueryDate_(gmailAfter),
    gmailBeforeLabel: joyfitFormatGmailQueryDate_(gmailBefore),
    anchorLabel: joyfitFormatGmailQueryDate_(d)
  };
}

function joyfitFormatSheetNameFromYm_(yy, mm) {
  return String(yy) + "年" + String(mm) + "月";
}

function joyfitBuildCurrentMonthGmailQuery_(window) {
  return (
    'from:info@joyfit-service.jp subject:"未納金のお支払いについて" after:' +
    window.gmailAfterLabel +
    " before:" +
    window.gmailBeforeLabel
  );
}

/** App入金メール行の同一判定（Message-ID＋対象月＋氏名） */
function joyfitPaymentRowDedupeKey_(row) {
  var monthKey = joyfitDateToMonthKey_(row[0]);
  var nameKey = normalizeMemberNameKey_(String(row[2] || "").trim());
  var msgId = String(row[5] || "").trim();
  return msgId + "|" + monthKey + "|" + nameKey;
}

/** App入金_Data: 対象月分だけ差し替え（それ以外の行は保持） */
function joyfitMergeAppPaymentMailRows_(dataSheet, incomingRows, monthKeySet, mailDataColCount) {
  var kept = [];
  var existingKeys = new Set();
  if (dataSheet && dataSheet.getLastRow() >= 2) {
    var existing = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, mailDataColCount).getValues();
    for (var i = 0; i < existing.length; i++) {
      var key = joyfitDateToMonthKey_(existing[i][0]);
      if (!monthKeySet.has(key)) {
        kept.push(existing[i]);
      } else {
        existingKeys.add(joyfitPaymentRowDedupeKey_(existing[i]));
      }
    }
  }
  var added = [];
  var newRows = [];
  for (var j = 0; j < incomingRows.length; j++) {
    var k2 = joyfitDateToMonthKey_(incomingRows[j][0]);
    if (!monthKeySet.has(k2)) continue;
    added.push(incomingRows[j]);
    if (!existingKeys.has(joyfitPaymentRowDedupeKey_(incomingRows[j]))) {
      newRows.push(incomingRows[j]);
    }
  }
  var merged = kept.concat(added);
  merged.sort(function(a, b) {
    return a[1] - b[1];
  });
  return { merged: merged, addedCount: added.length, newRows: newRows, newCount: newRows.length };
}

/** 今回新規取込分を「月次タブ名＋氏名」だけの短文に整形 */
function joyfitFormatNewPaymentNamesSummary_(newRows, monthWindow) {
  if (!newRows || newRows.length === 0) return "新しい入金はありませんでした。";

  var byMonth = {};
  for (var i = 0; i < newRows.length; i++) {
    var d = newRows[i][0];
    var name = String(newRows[i][2] || "").trim();
    if (!(d instanceof Date) || !name) continue;
    var yy = parseInt(String(d.getFullYear()).slice(-2), 10);
    var mm = d.getMonth() + 1;
    var sheetName = joyfitFormatSheetNameFromYm_(yy, mm);
    if (!byMonth[sheetName]) byMonth[sheetName] = [];
    byMonth[sheetName].push(name);
  }

  var lines = [];
  var months = monthWindow && monthWindow.months ? monthWindow.months : [];
  for (var mi = 0; mi < months.length; mi++) {
    var sn = months[mi].sheetName;
    var names = byMonth[sn] || [];
    var seen = {};
    var unique = [];
    for (var ni = 0; ni < names.length; ni++) {
      if (seen[names[ni]]) continue;
      seen[names[ni]] = true;
      unique.push(names[ni]);
    }
    if (unique.length === 0) continue;
    lines.push(sn);
    for (var ui = 0; ui < unique.length; ui++) {
      lines.push("・" + unique[ui]);
    }
    lines.push("");
  }
  if (lines.length === 0) return "新しい入金はありませんでした。";
  return lines.join("\n").trim();
}

function joyfitDescribeMonthWindow_(window) {
  if (!window || !window.months || window.months.length === 0) return "";
  return window.months
    .map(function(m) {
      return m.sheetName;
    })
    .join("、");
}

/**
 * =========================================================================
 * 以下、既存のJOYFIT入金管理スクリプト
 * =========================================================================
 */
function updateJoyfitPayments() {
  const UI_SHEET_NAME = 'App入金';
  const DATA_SHEET_NAME = 'App入金_Data';
  const LABEL_NAME = 'App入金';
  const TAX_RATE = 1.1;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthWindow = joyfitBuildCurrentMonthWindowFromDate_(new Date());
  const QUERY = joyfitBuildCurrentMonthGmailQuery_(monthWindow);

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const headers = [
    [
      "対象年月",
      "受信日時",
      "氏名",
      "金額(税込)",
      "金額(税抜)",
      "Message-ID",
      "月次行",
      "月次シート",
      "経路"
    ]
  ];
  const dataColCount = 9;
  const mailDataColCount = 6;

  if (!dataSheet) {
    dataSheet = ss.insertSheet(DATA_SHEET_NAME);
    dataSheet.getRange(1, 1, 1, dataColCount).setValues(headers);
    dataSheet.hideSheet();
  } else if (dataSheet.getLastRow() < 1) {
    dataSheet.getRange(1, 1, 1, dataColCount).setValues(headers);
  }

  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) label = GmailApp.createLabel(LABEL_NAME);

  const threads = GmailApp.search(QUERY, 0, 200);
  const parsedRecords = [];

  if (threads.length > 0) {
    for (const thread of threads) {
      const messages = thread.getMessages();
      let hasJoyfitMail = false;

      for (const message of messages) {
        if (message.getFrom().indexOf('joyfit-service.jp') === -1) continue;
        hasJoyfitMail = true;
        const records = parseJoyfitEmailV5(message.getPlainBody(), message.getDate(), message.getId(), TAX_RATE);
        if (records.length > 0) records.forEach(r => parsedRecords.push(r));
      }
      if (hasJoyfitMail) thread.addLabel(label);
    }
  }

  var mergeResult = joyfitMergeAppPaymentMailRows_(
    dataSheet,
    parsedRecords,
    monthWindow.monthKeySet,
    mailDataColCount
  );
  var mergedRows = mergeResult.merged;
  var newRows = mergeResult.newRows || [];
  var oldLastRow = dataSheet.getLastRow();

  dataSheet.getRange(1, 1, 1, dataColCount).setValues(headers);
  if (mergedRows.length > 0) {
    dataSheet.getRange(2, 1, mergedRows.length, mailDataColCount).setValues(mergedRows);
    dataSheet.getRange("A:A").setNumberFormat("yyyy/mm/dd");
    dataSheet.getRange("B:B").setNumberFormat("yyyy/mm/dd HH:mm");
    dataSheet.getRange("D:E").setNumberFormat("#,##0");
  }
  if (oldLastRow > mergedRows.length + 1) {
    dataSheet
      .getRange(mergedRows.length + 2, 1, oldLastRow - mergedRows.length - 1, dataColCount)
      .clearContent();
  }

  clearLegacyAppPaymentLogColumns_(ss, monthWindow.sheetNameSet);
  reconcileMonthlySheets(ss, DATA_SHEET_NAME, {
    sheetNameSet: monthWindow.sheetNameSet
  });
  backfillAppPaymentSheetLinks_(ss, dataSheet, { monthKeySet: monthWindow.monthKeySet });
  var uiSheet = ss.getSheetByName(UI_SHEET_NAME);
  if (uiSheet) {
    refreshAppPaymentUiDisplay_(ss, uiSheet, dataSheet);
  } else {
    setupUISheetV5(ss, UI_SHEET_NAME, DATA_SHEET_NAME);
  }
  Browser.msgBox(joyfitFormatNewPaymentNamesSummary_(newRows, monthWindow));
}

var JOYFIT_APP_DATA_SHEET_NAME_ = "App入金_Data";
var JOYFIT_RECONCILE_AUDIT_SHEET_NAME_ = "消込見直しログ";

/** 旧バージョンが書き込んでいた AA〜AI 列（27〜35）—実行時に内容のみクリア */
var JOYFIT_LEGACY_APP_LOG_COL_START_ = 27;
var JOYFIT_LEGACY_APP_LOG_COL_COUNT_ = 9;

/** 入金方法（App入金消込は V 列のみ。W/X 等は触らない） */
var JOYFIT_RECONCILE_METHOD_COL_ = 22; // V

/** 月次シート・入金済み行のグレーアウト（B〜Lのみ。M列以降は触らない） */
var JOYFIT_PAID_ROW_BG_COLOR_ = "#7f7f7f";
var JOYFIT_PAID_ROW_TEXT_COLOR_ = "#ffffff";
var JOYFIT_RECONCILE_VISUAL_COL_START_ = 2; // B
var JOYFIT_RECONCILE_VISUAL_COL_END_ = 12; // L

/** App入金_Data: 氏名キー → 入金済み対象月の Set（"yy-mm"） */
function buildAppPaidMemberMonthMap_(dataSheet) {
  var map = new Map();
  if (!dataSheet || dataSheet.getLastRow() < 2) return map;
  var values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues();
  for (var vi = 0; vi < values.length; vi++) {
    var date = values[vi][0];
    var name = values[vi][2];
    if (!(date instanceof Date) || !name) continue;
    var nameKey = normalizeMemberNameKey_(String(name).trim());
    if (!nameKey) continue;
    var monthKey = String(date.getFullYear()).slice(-2) + "-" + String(date.getMonth() + 1);
    if (!map.has(nameKey)) map.set(nameKey, new Set());
    map.get(nameKey).add(monthKey);
  }
  return map;
}

/**
 * 入金消込: 氏名が一致し、App入金の対象年月がその月次シートの月と同じか
 * 例: 「26年5月」シート → App入金の対象年月が5月の行だけ（4月分は4月シートで照合）
 */
function isMemberPaidForMonthlyRow_(nameKey, sheetYm, paidMonthMap) {
  if (!nameKey || !sheetYm || !paidMonthMap.has(nameKey)) return false;
  var sheetKey = sheetYm.yy + "-" + sheetYm.mm;
  return paidMonthMap.get(nameKey).has(sheetKey);
}

function joyfitSheetYearMonthKey_(sheetYm) {
  return sheetYm.yy + "-" + sheetYm.mm;
}

/** App入金_Data のうち、指定月次シート月と一致する入金者 {key, name} 一覧 */
function listAppPaidEntriesForSheetMonth_(dataSheet, sheetYm) {
  var list = [];
  if (!dataSheet || !sheetYm || dataSheet.getLastRow() < 2) return list;
  var sheetKey = joyfitSheetYearMonthKey_(sheetYm);
  var values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues();
  for (var vi = 0; vi < values.length; vi++) {
    var date = values[vi][0];
    var name = String(values[vi][2] || "").trim();
    if (!(date instanceof Date) || !name) continue;
    var monthKey = String(date.getFullYear()).slice(-2) + "-" + String(date.getMonth() + 1);
    if (monthKey !== sheetKey) continue;
    var nameKey = normalizeMemberNameKey_(name);
    if (nameKey) list.push({ key: nameKey, name: name });
  }
  return list;
}

function rowHasJoyfitReconcileStyle_(payTypeBg, memberBg, memberFontLine) {
  return (
    isJoyfitPaidReconcileBackground_(payTypeBg) ||
    isJoyfitPaidReconcileBackground_(memberBg) ||
    isJoyfitNameCellPaidByStrikethrough_(memberFontLine)
  );
}

function joyfitAppendReconcileStatsLines_(lines, stats) {
  if (!stats) return;
  lines.push(
    stats.sheetName +
      ": App入金 " +
      stats.appPaidCount +
      "名 / シート消込 " +
      stats.reconciledOnSheet +
      "名 / 一致 " +
      stats.matchedCount +
      "名"
  );
  if (stats.appPaidCount !== stats.reconciledOnSheet) {
    lines.push("  → 人数がずれています（下記を確認）");
  }
  if (stats.extraGrayNames.length > 0) {
    lines.push("  ※誤消込: " + stats.extraGrayNames.slice(0, 8).join("、") + (stats.extraGrayNames.length > 8 ? " 他" : ""));
  }
  if (stats.missingGrayNames.length > 0) {
    lines.push("  ※未消込: " + stats.missingGrayNames.slice(0, 8).join("、") + (stats.missingGrayNames.length > 8 ? " 他" : ""));
  }
  if (stats.appOnlyNames.length > 0) {
    lines.push("  ※未掲載: " + stats.appOnlyNames.slice(0, 8).join("、") + (stats.appOnlyNames.length > 8 ? " 他" : ""));
  }
}

/** 月次シートの人数突合（App入金当月件数 vs グレー行数） */
function buildJoyfitReconcileSummaryText_(ss, dataSheetName, opt) {
  opt = opt || {};
  var dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet || dataSheet.getLastRow() < 2) return "";

  var paidMonthMap = buildAppPaidMemberMonthMap_(dataSheet);
  var lines = [];

  if (opt.monthWindow && opt.monthWindow.months && opt.monthWindow.months.length) {
    for (var mi = 0; mi < opt.monthWindow.months.length; mi++) {
      var sheetName = opt.monthWindow.months[mi].sheetName;
      if (!ss.getSheetByName(sheetName)) {
        lines.push(sheetName + ": 月次シートなし");
        continue;
      }
      joyfitAppendReconcileStatsLines_(
        lines,
        computeJoyfitReconcileStatsForSheet_(ss, sheetName, paidMonthMap, dataSheet)
      );
    }
    return lines.join("\n");
  }

  var activeName = "";
  try {
    activeName = ss.getActiveSheet().getName();
  } catch (e) {}
  if (!parseJoyfitSheetYearMonth_(activeName)) return "";

  joyfitAppendReconcileStatsLines_(
    lines,
    computeJoyfitReconcileStatsForSheet_(ss, activeName, paidMonthMap, dataSheet)
  );
  return lines.join("\n");
}

function computeJoyfitReconcileStatsForSheet_(ss, sheetName, paidMonthMap, dataSheet) {
  var sheetYm = parseJoyfitSheetYearMonth_(sheetName);
  if (!sheetYm) return null;

  var targetSheet = ss.getSheetByName(sheetName);
  var layout = targetSheet ? resolveJoyfitSheetLayoutForScan_(targetSheet) : null;
  if (!targetSheet || !layout) return null;

  var appEntries = listAppPaidEntriesForSheetMonth_(dataSheet, sheetYm);
  var appKeySet = new Set();
  for (var ei = 0; ei < appEntries.length; ei++) appKeySet.add(appEntries[ei].key);
  var appPaidCount = appEntries.length;

  var dataStart = layout.dataStartRow;
  var targetLastRow = targetSheet.getLastRow();
  if (targetLastRow < dataStart) {
    return {
      sheetName: sheetName,
      appPaidCount: appPaidCount,
      reconciledOnSheet: 0,
      matchedCount: 0,
      extraGrayNames: [],
      missingGrayNames: [],
      appOnlyNames: []
    };
  }

  var numRows = targetLastRow - dataStart + 1;
  var lastCol = Math.max(targetSheet.getLastColumn(), layout.colMemberName);
  var rows = targetSheet.getRange(dataStart, 1, numRows, lastCol).getValues();
  var vValues = targetSheet.getRange(dataStart, JOYFIT_RECONCILE_METHOD_COL_, numRows, 1).getValues();
  var bBackgrounds = targetSheet.getRange(dataStart, JOYFIT_RECONCILE_VISUAL_COL_START_, numRows, 1).getBackgrounds();
  var reconciledOnSheet = 0;
  var matchedCount = 0;
  var extraGrayNames = [];
  var foundOnSheet = new Set();
  var reconciledByKey = {};

  for (var i = 0; i < rows.length; i++) {
    var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
    if (!rowName) continue;
    var nameKey = normalizeMemberNameKey_(rowName);
    var isReconciled = isJoyfitAppReconciledFromSnapshot_(vValues[i][0], bBackgrounds[i][0]);
    if (isReconciled) {
      reconciledOnSheet++;
      if (appKeySet.has(nameKey)) matchedCount++;
      else extraGrayNames.push(rowName);
    }
    reconciledByKey[nameKey] = reconciledByKey[nameKey] || isReconciled;
    foundOnSheet.add(nameKey);
  }

  var missingGrayNames = [];
  var appOnlyNames = [];
  for (var ai = 0; ai < appEntries.length; ai++) {
    var entry = appEntries[ai];
    if (!foundOnSheet.has(entry.key)) {
      appOnlyNames.push(entry.name);
      continue;
    }
    if (!reconciledByKey[entry.key]) missingGrayNames.push(entry.name);
  }

  return {
    sheetName: sheetName,
    appPaidCount: appPaidCount,
    reconciledOnSheet: reconciledOnSheet,
    matchedCount: matchedCount,
    extraGrayNames: extraGrayNames,
    missingGrayNames: missingGrayNames,
    appOnlyNames: appOnlyNames
  };
}

function formatJoyfitSheetNameFromDate_(date) {
  if (!(date instanceof Date)) return "";
  var yy = parseInt(String(date.getFullYear()).slice(-2), 10);
  var mm = date.getMonth() + 1;
  return yy + "年" + mm + "月";
}

/** 月次シート名または gid（URLの#gid=）からシートを取得 */
function resolveMonthlySheet_(ss, sheetNameOrGid) {
  var s = String(sheetNameOrGid || "").trim();
  if (!s || !ss) return null;
  var sheet = ss.getSheetByName(s);
  if (sheet) return sheet;
  if (/^\d{5,}$/.test(s)) {
    var gid = parseInt(s, 10);
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getSheetId() === gid) return all[i];
    }
  }
  return null;
}

/** 対象年月から存在する月次シート名を返す（なければ空） */
function getValidMonthlySheetName_(ss, date) {
  var name = formatJoyfitSheetNameFromDate_(date);
  if (!name || !ss) return "";
  return resolveMonthlySheet_(ss, name) ? name : "";
}

/** 月次シート上で氏名が一致する行番号（先頭一致） */
function findMemberRowOnMonthlySheet_(ss, sheetName, nameKey) {
  if (!ss || !sheetName || !nameKey) return 0;
  var sheet = resolveMonthlySheet_(ss, sheetName);
  if (!sheet) return 0;
  sheetName = sheet.getName();
  var layout = resolveJoyfitSheetLayoutForScan_(sheet);
  if (!layout) return 0;
  var dataStart = layout.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStart) return 0;
  var lastCol = Math.max(sheet.getLastColumn(), layout.colMemberName);
  var rows = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getValues();
  for (var i = 0; i < rows.length; i++) {
    var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
    if (!rowName) continue;
    if (normalizeMemberNameKey_(rowName) === nameKey) return dataStart + i;
  }
  return 0;
}

function readMemberSnapshotFromSheetRow_(row, layout) {
  var snap = { payType: "", memberId: "", memberName: "", phone: "" };
  if (!row || !layout) return snap;
  if (layout.colPayType > 0 && layout.colPayType <= row.length) {
    snap.payType = String(row[layout.colPayType - 1] || "").trim();
  }
  if (layout.colMemberId > 0 && layout.colMemberId <= row.length) {
    snap.memberId = String(row[layout.colMemberId - 1] || "").trim();
  }
  if (layout.colMemberName > 0 && layout.colMemberName <= row.length) {
    snap.memberName = String(row[layout.colMemberName - 1] || "").trim();
  }
  if (layout.colPhone > 0 && layout.colPhone <= row.length) {
    snap.phone = String(row[layout.colPhone - 1] || "").trim();
  }
  return snap;
}

function findAppPaymentRecordForMonth_(dataSheet, nameKey, sheetYm) {
  if (!dataSheet || !nameKey || !sheetYm || dataSheet.getLastRow() < 2) return null;
  var sheetKey = joyfitSheetYearMonthKey_(sheetYm);
  var values = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 5).getValues();
  for (var vi = 0; vi < values.length; vi++) {
    var date = values[vi][0];
    var name = values[vi][2];
    if (!(date instanceof Date) || !name) continue;
    var monthKey = String(date.getFullYear()).slice(-2) + "-" + String(date.getMonth() + 1);
    if (monthKey !== sheetKey) continue;
    if (normalizeMemberNameKey_(String(name).trim()) !== nameKey) continue;
    return {
      targetDate: date,
      name: String(name).trim(),
      amountIncl: values[vi][3],
      amountExcl: values[vi][4]
    };
  }
  return null;
}

/** J列=支払額(H+I)、K列=手数料(IFS) — 原本どおり数式をセット */
function buildJoyfitFeeFormula_(rowNum) {
  return (
    "=IFS(I" +
    rowNum +
    ">=30000,\"300\",I" +
    rowNum +
    ">=10000,\"250\",I" +
    rowNum +
    ">=3000,\"200\",I" +
    rowNum +
    ">=2000,\"170\",I" +
    rowNum +
    ">=200,\"150\",I" +
    rowNum +
    "<200,\"\")"
  );
}

function applyJoyfitMonthlyRowFormulasJK_(sheet, dataStartRow, rowCount) {
  if (!sheet || rowCount <= 0 || dataStartRow < 1) return;
  var rowNums = [];
  for (var i = 0; i < rowCount; i++) rowNums.push(dataStartRow + i);
  applyJoyfitMonthlyRowFormulasForRows_(sheet, rowNums);
}

/** 会員行ごとに J=H+I、K=手数料(IFS)、L=J+K */
function applyJoyfitMonthlyRowFormulasForRows_(sheet, rowNums) {
  if (!sheet || !rowNums || rowNums.length === 0) return;
  for (var i = 0; i < rowNums.length; i++) {
    var rowNum = rowNums[i];
    sheet.getRange(rowNum, JOYFIT_COL_PAYMENT_).setFormula("=H" + rowNum + "+I" + rowNum);
    sheet.getRange(rowNum, JOYFIT_COL_FEE_).setFormula(buildJoyfitFeeFormula_(rowNum));
    sheet.getRange(rowNum, JOYFIT_COL_TOTAL_).setFormula("=J" + rowNum + "+K" + rowNum);
  }
}

/**
 * 月次シート1枚を原本体裁に整える（データは消さない）
 * - AA〜AI の不要文字を削除
 * - 会員行: B〜I 中央揃え、B〜L 区分色・取り消し線解除、J/K/L 数式
 * - M〜T・U・W・X・V は触らない
 */
function restoreJoyfitMonthlySheetAppearance_(sheet) {
  if (!sheet) return { ok: false, message: "シートがありません" };

  var layout = resolveJoyfitSheetLayoutForScan_(sheet);
  if (!layout) {
    return { ok: false, message: "見出し（会員番号・電話番号など）が標準と違うためスキップしました" };
  }

  var dataStart = layout.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStart) {
    clearLegacyAppPaymentLogColumnsOnSheet_(sheet, layout.headerRow, dataStart);
    return { ok: true, message: "データ行なし（AA〜AIのみクリア）", memberRows: 0 };
  }

  clearLegacyAppPaymentLogColumnsOnSheet_(sheet, layout.headerRow, dataStart);

  var numRows = lastRow - dataStart + 1;
  var lastCol = Math.max(sheet.getLastColumn(), layout.colMemberName, JOYFIT_COL_TOTAL_);
  var rows = sheet.getRange(dataStart, 1, numRows, lastCol).getValues();
  var aDisplay = sheet.getRange(dataStart, 1, numRows, 1).getDisplayValues();

  var groupState = { current: 1 };
  var indexInGroup = { 1: 0, 2: 0, 3: 0 };
  var memberRowNums = [];

  for (var ri = 0; ri < numRows; ri++) {
    updateJoyfitRowGroupStateFromColumnA_(groupState, aDisplay[ri][0]);
    var rowName = extractMemberNameFromSheetRow_(rows[ri], layout);
    if (!rowName) continue;

    var grp = groupState.current || 1;
    indexInGroup[grp]++;
    var rowNum = dataStart + ri;
    memberRowNums.push(rowNum);

    var dataRange = sheet.getRange(rowNum, JOYFIT_UNPAID_DATA_COL_START_, 1, JOYFIT_UNPAID_DATA_COL_COUNT_);
    dataRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

    if (rowIsJoyfitAppReconciled_(sheet, rowNum)) continue;

    var bgColor = joyfitGroupRowBgColor_(grp, indexInGroup[grp]);
    var blRange = sheet.getRange(rowNum, JOYFIT_UNPAID_DATA_COL_START_, 1, JOYFIT_COL_TOTAL_ - 1);
    blRange.setBackground(bgColor).setFontLine("none").setFontColor(null);
  }

  applyJoyfitMonthlyRowFormulasForRows_(sheet, memberRowNums);

  return {
    ok: true,
    message: sheet.getName() + "：" + memberRowNums.length + "名分を整形",
    memberRows: memberRowNums.length
  };
}

/** メニュー用：開いている月次シート、なければ全「○年○月」シートを整形 */
function restoreJoyfitMonthlySheetsFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var active = ss.getActiveSheet();
  var targets = [];

  if (active && parseJoyfitSheetYearMonth_(active.getName())) {
    targets.push(active);
  } else {
    ss.getSheets().forEach(function(sh) {
      if (parseJoyfitSheetYearMonth_(sh.getName())) targets.push(sh);
    });
  }

  if (targets.length === 0) {
    Browser.msgBox("「○年○月」形式の月次シートが見つかりません。");
    return;
  }

  var confirm = SpreadsheetApp.getUi().alert(
    "月次シートの見た目を整える",
    "次を実行します（CSVは不要・B〜Iの文字は消しません）。\n\n" +
      "・AA〜AI列の不要な文字を削除\n" +
      "・会員行の B〜I を中央揃え\n" +
      "・J列=H+I、K列=手数料、L列=J+K の数式を入れ直し\n" +
      "・B〜L の区分色（青/黄/緑）と取り消し線の解除\n\n" +
      "触らない列：M〜T、U、V、W、X\n\n" +
      "対象：" +
      targets.map(function(t) {
        return t.getName();
      }).join("、") +
      "\n\n実行しますか？",
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (confirm !== SpreadsheetApp.getUi().Button.OK) return;

  var lines = [];
  var totalMembers = 0;
  for (var ti = 0; ti < targets.length; ti++) {
    var result = restoreJoyfitMonthlySheetAppearance_(targets[ti]);
    if (result.ok) {
      lines.push(result.message);
      totalMembers += result.memberRows || 0;
    } else {
      lines.push(targets[ti].getName() + "：スキップ（" + result.message + "）");
    }
  }

  lines.push("");
  lines.push("※月次タブ（26年5月など）専用です。App入金タブは「入金メールを反映」で整います。");
  lines.push("入金のグレーアウトは「入金メールを反映」。直らないときは「その他」→「入金消込をやり直す」。");
  Browser.msgBox("✅ 整形完了\n\n" + lines.join("\n"));
}

/** 旧Verが使っていた AA〜AI 列（27〜35）の内容だけクリア（見た目は触らない） */
function clearLegacyAppPaymentLogColumnsOnSheet_(sheet, headerRow, dataStartRow) {
  if (!sheet) return;
  var start = Math.max(1, parseInt(headerRow, 10) || parseInt(dataStartRow, 10) || 1);
  var lastRow = Math.max(sheet.getLastRow(), start);
  if (lastRow < start) return;
  sheet
    .getRange(start, JOYFIT_LEGACY_APP_LOG_COL_START_, lastRow - start + 1, JOYFIT_LEGACY_APP_LOG_COL_COUNT_)
    .clearContent();
}

function clearLegacyAppPaymentLogColumns_(ss, onlySheetNameSet) {
  if (!ss) return;
  ss.getSheets().forEach(function(sh) {
    var sName = sh.getName();
    if (!parseJoyfitSheetYearMonth_(sName)) return;
    if (onlySheetNameSet && !onlySheetNameSet.has(sName)) return;
    var layout = resolveJoyfitSheetLayoutForScan_(sh);
    if (!layout) return;
    clearLegacyAppPaymentLogColumnsOnSheet_(sh, layout.headerRow, layout.dataStartRow);
  });
}

/** 月次シートの氏名→行番号インデックス（1回の読み取りで構築） */
function buildJoyfitMemberRowIndexForSheet_(sheet) {
  var map = new Map();
  if (!sheet) return map;
  var layout = resolveJoyfitSheetLayoutForScan_(sheet);
  if (!layout) return map;
  var dataStart = layout.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStart) return map;
  var lastCol = Math.max(sheet.getLastColumn(), layout.colMemberName);
  var rows = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getValues();
  for (var i = 0; i < rows.length; i++) {
    var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
    if (!rowName) continue;
    var nameKey = normalizeMemberNameKey_(rowName);
    if (nameKey && !map.has(nameKey)) map.set(nameKey, dataStart + i);
  }
  return map;
}

/**
 * App入金_Data に月次シートの行番号・シート名を書き戻す
 * @param {{ monthKeySet?: Set<string> }} [opt] 指定時はその月だけ再計算（他月は既存値を保持）
 */
function backfillAppPaymentSheetLinks_(ss, dataSheet, opt) {
  opt = opt || {};
  var monthKeySet = opt.monthKeySet || null;
  if (!dataSheet || dataSheet.getLastRow() < 2) return;
  var lastRow = dataSheet.getLastRow();
  var values = dataSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var links = [];
  var indexCache = {};
  for (var vi = 0; vi < values.length; vi++) {
    var date = values[vi][0];
    var monthKey = joyfitDateToMonthKey_(date);
    if (monthKeySet && !monthKeySet.has(monthKey)) {
      links.push([values[vi][6] || "", values[vi][7] || "", values[vi][8] || "アプリ"]);
      continue;
    }
    var name = String(values[vi][2] || "").trim();
    var sheetName = getValidMonthlySheetName_(ss, date);
    var nameKey = normalizeMemberNameKey_(name);
    var sheetRow = 0;
    if (sheetName && nameKey) {
      if (!indexCache[sheetName]) {
        var sh = resolveMonthlySheet_(ss, sheetName);
        indexCache[sheetName] = buildJoyfitMemberRowIndexForSheet_(sh);
      }
      sheetRow = indexCache[sheetName].get(nameKey) || 0;
    }
    links.push([sheetRow || "", sheetName, "アプリ"]);
  }
  if (links.length > 0) {
    dataSheet.getRange(2, 7, links.length, 3).setValues(links);
  }
}

/** 消込解除（グレー・取り消し線を戻し、V列を空に） */
function clearJoyfitReconcileOnRow_(targetSheet, rowNum, zebraBgColor) {
  var visual = targetSheet.getRange(
    rowNum,
    JOYFIT_RECONCILE_VISUAL_COL_START_,
    1,
    JOYFIT_RECONCILE_VISUAL_COL_END_ - JOYFIT_RECONCILE_VISUAL_COL_START_ + 1
  );
  visual.setBackground(zebraBgColor || null).setFontLine("none").setFontColor(null);

  var methodCell = targetSheet.getRange(rowNum, JOYFIT_RECONCILE_METHOD_COL_);
  methodCell.clearDataValidations();
  methodCell.clearContent();
}

/** App入金消込：B〜Lグレー+取り消し線、V列=アプリ入金 */
function applyJoyfitReconcileOnRow_(targetSheet, rowNum, ruleMethod) {
  var visual = targetSheet.getRange(
    rowNum,
    JOYFIT_RECONCILE_VISUAL_COL_START_,
    1,
    JOYFIT_RECONCILE_VISUAL_COL_END_ - JOYFIT_RECONCILE_VISUAL_COL_START_ + 1
  );
  visual.setBackground(JOYFIT_PAID_ROW_BG_COLOR_);
  visual.setFontLine("line-through").setFontColor(JOYFIT_PAID_ROW_TEXT_COLOR_);

  var methodCell = targetSheet.getRange(rowNum, JOYFIT_RECONCILE_METHOD_COL_);
  methodCell.setDataValidation(ruleMethod);
  methodCell.setValue("アプリ入金");
  methodCell.setHorizontalAlignment("center");
}

/** 行がApp入金消込済みか */
function rowIsJoyfitAppReconciled_(sheet, rowNum) {
  if (!sheet || rowNum < 1) return false;
  var vVal = String(sheet.getRange(rowNum, JOYFIT_RECONCILE_METHOD_COL_).getValue() || "").trim();
  return isJoyfitAppReconciledFromSnapshot_(vVal, sheet.getRange(rowNum, JOYFIT_RECONCILE_VISUAL_COL_START_).getBackground());
}

/** バッチ照合用：V列とB列背景から消込済みか判定 */
function isJoyfitAppReconciledFromSnapshot_(vVal, firstVisualBg) {
  if (String(vVal || "").trim() === "アプリ入金") return true;
  return isJoyfitPaidReconcileBackground_(firstVisualBg);
}

/** 会員名列のみから氏名取得（G列の未納理由などを氏名と誤認しない） */
function extractMemberNameFromSheetRow_(row, layout) {
  if (!row || !layout) return "";
  var cols = [layout.colMemberName - 1, layout.colMemberId - 1];
  for (var i = 0; i < cols.length; i++) {
    var c = cols[i];
    if (c < 0 || c >= row.length) continue;
    var s = String(row[c] || "").trim();
    if (looksLikeMemberName_(s)) return s;
  }
  return "";
}

/**
 * 月次シートを App入金_Data と照合
 * - 当月の App入金がある行 → B〜Lグレー+取り消し線、V列=アプリ入金
 * - App入金にないのに消込表示がある行 → 区分色に戻して解除
 * @param {{ sheetNameSet?: Set<string> }} [opt] 指定時はその月次タブのみ（当月用）
 */
function reconcileMonthlySheets(ss, dataSheetName, opt) {
  opt = opt || {};
  var onlySheetNameSet = opt.sheetNameSet || null;

  const dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) return { reconciled: 0, cleared: 0 };

  const paidMonthMap = buildAppPaidMemberMonthMap_(dataSheet);
  if (paidMonthMap.size === 0) return { reconciled: 0, cleared: 0 };

  const ruleMethod = SpreadsheetApp.newDataValidation()
    .requireValueInList(["レジ入金", "アプリ入金"], true)
    .setAllowInvalid(false)
    .build();

  let totalReconciled = 0;
  let totalCleared = 0;

  ss.getSheets().forEach(function(targetSheet) {
    var sheetName = targetSheet.getName();
    var sheetYm = parseJoyfitSheetYearMonth_(sheetName);
    if (!sheetYm) return;
    if (onlySheetNameSet && !onlySheetNameSet.has(sheetName)) return;

    const layout = resolveJoyfitSheetLayoutForScan_(targetSheet);
    if (!layout) return;
    const dataStart = layout.dataStartRow;
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < dataStart) return;

    var numRows = targetLastRow - dataStart + 1;
    var visualColCount =
      JOYFIT_RECONCILE_VISUAL_COL_END_ - JOYFIT_RECONCILE_VISUAL_COL_START_ + 1;
    var lastCol = Math.max(targetSheet.getLastColumn(), layout.colMemberName);
    const rows = targetSheet.getRange(dataStart, 1, numRows, lastCol).getValues();
    const aDisplay = targetSheet.getRange(dataStart, 1, numRows, 1).getDisplayValues();
    var backgrounds = targetSheet
      .getRange(dataStart, JOYFIT_RECONCILE_VISUAL_COL_START_, numRows, visualColCount)
      .getBackgrounds();
    var fontLines = targetSheet
      .getRange(dataStart, JOYFIT_RECONCILE_VISUAL_COL_START_, numRows, visualColCount)
      .getFontLines();
    var fontColors = targetSheet
      .getRange(dataStart, JOYFIT_RECONCILE_VISUAL_COL_START_, numRows, visualColCount)
      .getFontColors();
    var vValues = targetSheet.getRange(dataStart, JOYFIT_RECONCILE_METHOD_COL_, numRows, 1).getValues();
    var groupState = { current: 1 };
    var indexInGroup = { 1: 0, 2: 0, 3: 0 };
    var sheetDirty = false;

    for (let i = 0; i < rows.length; i++) {
      updateJoyfitRowGroupStateFromColumnA_(groupState, aDisplay[i][0]);
      var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
      if (!rowName) continue;

      var grp = groupState.current || 1;
      indexInGroup[grp]++;
      var zebraBg = joyfitGroupRowBgColor_(grp, indexInGroup[grp]);
      var nameKey = normalizeMemberNameKey_(rowName);
      var shouldReconcile = isMemberPaidForMonthlyRow_(nameKey, sheetYm, paidMonthMap);
      var hadAppReconcile = isJoyfitAppReconciledFromSnapshot_(vValues[i][0], backgrounds[i][0]);

      if (shouldReconcile) {
        for (var vc = 0; vc < visualColCount; vc++) {
          backgrounds[i][vc] = JOYFIT_PAID_ROW_BG_COLOR_;
          fontLines[i][vc] = "line-through";
          fontColors[i][vc] = JOYFIT_PAID_ROW_TEXT_COLOR_;
        }
        vValues[i][0] = "アプリ入金";
        totalReconciled++;
        sheetDirty = true;
      } else if (hadAppReconcile) {
        for (var vc2 = 0; vc2 < visualColCount; vc2++) {
          backgrounds[i][vc2] = zebraBg;
          fontLines[i][vc2] = "none";
          fontColors[i][vc2] = null;
        }
        vValues[i][0] = "";
        totalCleared++;
        sheetDirty = true;
      }
    }

    if (sheetDirty) {
      var visualRange = targetSheet.getRange(
        dataStart,
        JOYFIT_RECONCILE_VISUAL_COL_START_,
        numRows,
        visualColCount
      );
      visualRange.setBackgrounds(backgrounds).setFontLines(fontLines).setFontColors(fontColors);
      var vRange = targetSheet.getRange(dataStart, JOYFIT_RECONCILE_METHOD_COL_, numRows, 1);
      vRange.setValues(vValues).setDataValidation(ruleMethod).setHorizontalAlignment("center");
    }
  });

  return { reconciled: totalReconciled, cleared: totalCleared };
}

function isJoyfitPaidReconcileBackground_(color) {
  var c = String(color || "")
    .toLowerCase()
    .replace(/\s/g, "");
  return c === "#7f7f7f" || c === "rgb(127,127,127)";
}

/** A列の区分ラベルから、データ行が属するグループ(1=貸倒,2=2ヶ月,3=1ヶ月)を更新 */
function updateJoyfitRowGroupStateFromColumnA_(groupState, aDisplay) {
  var a = String(aDisplay || "").replace(/\r\n|\n|\r/g, " ").trim();
  if (!a) return;
  if (isBadDebtGroupLabel_(a)) groupState.current = 1;
  else if (a.indexOf("2ヶ月") !== -1 || a.indexOf("2カ月") !== -1) groupState.current = 2;
  else if (
    a.indexOf("1ヶ月") !== -1 ||
    a.indexOf("1カ月") !== -1 ||
    a.indexOf("1ヵ月") !== -1
  ) {
    groupState.current = 3;
  }
}

/** 全月次シートの V列消込と AA〜AI 列（旧照合欄）の内容を解除 */
function clearJoyfitReconcileStylingOnMonthlySheets_(ss) {
  var clearedRows = 0;
  var sheetsTouched = 0;

  ss.getSheets().forEach(function(targetSheet) {
    var sheetName = targetSheet.getName();
    if (!parseJoyfitSheetYearMonth_(sheetName)) return;

    var layout = resolveJoyfitSheetLayoutForScan_(targetSheet);
    if (!layout) return;

    var dataStart = layout.dataStartRow;
    var targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < dataStart) return;

    clearLegacyAppPaymentLogColumnsOnSheet_(targetSheet, layout.headerRow, dataStart);

    var lastCol = Math.max(targetSheet.getLastColumn(), layout.colMemberName);
    var rows = targetSheet.getRange(dataStart, 1, targetLastRow - dataStart + 1, lastCol).getValues();
    var aDisplay = targetSheet.getRange(dataStart, 1, targetLastRow - dataStart + 1, 1).getDisplayValues();
    var groupState = { current: 1 };
    var indexInGroup = { 1: 0, 2: 0, 3: 0 };
    var sheetCleared = 0;

    for (var i = 0; i < rows.length; i++) {
      updateJoyfitRowGroupStateFromColumnA_(groupState, aDisplay[i][0]);
      var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
      if (!rowName) continue;
      var grp = groupState.current || 1;
      indexInGroup[grp]++;
      var zebraBg = joyfitGroupRowBgColor_(grp, indexInGroup[grp]);
      var rowNum = dataStart + i;
      if (rowIsJoyfitAppReconciled_(targetSheet, rowNum)) {
        clearJoyfitReconcileOnRow_(targetSheet, rowNum, zebraBg);
        sheetCleared++;
      }
    }

    if (sheetCleared > 0) {
      clearedRows += sheetCleared;
      sheetsTouched++;
    }
  });

  return { clearedRows: clearedRows, sheetsTouched: sheetsTouched };
}

/**
 * 全月次シートの会員一覧と、App入金（当月）との照合結果をログシートに出力
 */
function writeJoyfitReconcileAuditSheet_(ss, paidMonthMap) {
  var auditName = JOYFIT_RECONCILE_AUDIT_SHEET_NAME_;
  var auditSheet = ss.getSheetByName(auditName);
  if (!auditSheet) auditSheet = ss.insertSheet(auditName);
  else auditSheet.clear();

  var headers = [
    [
      "月次シート",
      "行",
      "氏名",
      "会員番号",
      "区分",
      "App入金(当月)",
      "消込表示",
      "備考"
    ]
  ];
  auditSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  auditSheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold");

  var logRows = [];
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");

  ss.getSheets().forEach(function(targetSheet) {
    var sheetName = targetSheet.getName();
    var sheetYm = parseJoyfitSheetYearMonth_(sheetName);
    if (!sheetYm) return;

    var layout = resolveJoyfitSheetLayoutForScan_(targetSheet);
    if (!layout) return;

    var dataStart = layout.dataStartRow;
    var targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < dataStart) return;

    var lastCol = Math.max(targetSheet.getLastColumn(), layout.colMemberName);
    var rows = targetSheet.getRange(dataStart, 1, targetLastRow - dataStart + 1, lastCol).getValues();
    var aDisplay = targetSheet.getRange(dataStart, 1, targetLastRow - dataStart + 1, 1).getDisplayValues();

    var groupState = { current: 1 };
    var groupLabels = { 1: "退会後未納貸倒候補", 2: "2ヶ月未納", 3: "1カ月未納" };
    var sheetKey = joyfitSheetYearMonthKey_(sheetYm);

    for (var i = 0; i < rows.length; i++) {
      updateJoyfitRowGroupStateFromColumnA_(groupState, aDisplay[i][0]);
      var rowName = extractMemberNameFromSheetRow_(rows[i], layout);
      if (!rowName) continue;

      var rowNum = dataStart + i;
      var memberId = "";
      if (layout.colMemberId > 0 && layout.colMemberId <= rows[i].length) {
        memberId = String(rows[i][layout.colMemberId - 1] || "").trim();
      }

      var nameKey = normalizeMemberNameKey_(rowName);
      var hasApp = paidMonthMap.has(nameKey) && paidMonthMap.get(nameKey).has(sheetKey);
      var reconcileShown = rowIsJoyfitAppReconciled_(targetSheet, rowNum);

      var note = "";
      if (hasApp && !reconcileShown) note = "App入金あり・V未設定（要確認）";
      else if (!hasApp && reconcileShown) note = "App入金なし・Vあり（誤消込の可能性）";
      else if (hasApp && reconcileShown) note = "照合OK";
      else note = "未入金";

      logRows.push([
        sheetName,
        rowNum,
        rowName,
        memberId,
        groupLabels[groupState.current || 1],
        hasApp ? "あり" : "なし",
        reconcileShown ? "グレー/V=アプリ入金" : "未消込",
        note
      ]);
    }
  });

  if (logRows.length > 0) {
    auditSheet.getRange(2, 1, logRows.length, headers[0].length).setValues(logRows);
  }

  auditSheet.setFrozenRows(1);
  auditSheet.autoResizeColumns(1, headers[0].length);
  auditSheet.getRange("A1").setNote(
    "最終更新: " +
      now +
      "\n\n「その他」→「入金消込をやり直す」の実行後に作成されます。\n" +
      "5月シートは App入金の5月分（氏名＋対象年月）と一致した行だけが「照合OK」になります。\n" +
      "「App入金なし・W/Xあり」が残る場合は要確認。"
  );

  return logRows.length;
}

/**
 * 【推奨】過去の誤ったグレーアウトを全員分クリアし、App入金_Data から消込を再実行
 */
function rebuildJoyfitReconcileFromAppData() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "入金消込をやり直す",
    "全「○年○月」シートの会員行について、次を実行します。\n\n" +
      "1. V列（アプリ入金）と AA〜AI 列（旧照合欄）をいったん解除\n" +
      "2. App入金_Data で当月分だけ グレー・取り消し線・V列アプリ入金 を再設定\n" +
      "3. 「消込見直しログ」に全員の一覧を出力\n\n" +
      "※先に「入金メールを反映」を実行しておくことを推奨します。\n\n" +
      "実行しますか？",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(JOYFIT_APP_DATA_SHEET_NAME_);
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    Browser.msgBox(
      "App入金_Data がありません。\n\n先に「入金メールを反映」を実行し、App入金を取り込んでから、もう一度お試しください。"
    );
    return;
  }

  clearLegacyAppPaymentLogColumns_(ss);
  var clearResult = clearJoyfitReconcileStylingOnMonthlySheets_(ss);
  var paidMonthMap = buildAppPaidMemberMonthMap_(dataSheet);
  var reconcileResult = reconcileMonthlySheets(ss, JOYFIT_APP_DATA_SHEET_NAME_);
  backfillAppPaymentSheetLinks_(ss, dataSheet);
  setupUISheetV5(ss, "App入金", JOYFIT_APP_DATA_SHEET_NAME_);
  var auditCount = writeJoyfitReconcileAuditSheet_(ss, paidMonthMap);
  var summaryText = buildJoyfitReconcileSummaryText_(ss, JOYFIT_APP_DATA_SHEET_NAME_);

  var msg =
    "✅ 消込の見直しが完了しました\n\n" +
    "・V列を解除した行: " +
    clearResult.clearedRows +
    " 行（" +
    clearResult.sheetsTouched +
    " シート）\n" +
    "・App入金で再消込: " +
    reconcileResult.reconciled +
    " 行\n" +
    "・誤消込解除: " +
    reconcileResult.cleared +
    " 行\n" +
    "・ログ出力: " +
    auditCount +
    " 名";
  if (summaryText) msg += "\n\n【人数突合】\n" + summaryText;
  msg +=
    "\n\n「" +
    JOYFIT_RECONCILE_AUDIT_SHEET_NAME_ +
    "」で確認してください。\n" +
    "26年5月なら「App入金19名＝シート消込19名」が理想です。";

  Browser.msgBox(msg);
}

/**
 * App入金シートの一覧（A3〜）を App入金_Data から書き込む（QUERYは使わない）
 */
function refreshAppPaymentUiDisplay_(ss, uiSheet, dataSheet) {
  if (!ss || !uiSheet || !dataSheet) return;
  applyAppPaymentUiChrome_(uiSheet);
  var filterVal = String(uiSheet.getRange("A1").getDisplayValue() || "").trim();
  var out = [];
  var lastRow = dataSheet.getLastRow();
  if (lastRow >= 2) {
    var values = dataSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var tz = ss.getSpreadsheetTimeZone();
    for (var vi = 0; vi < values.length; vi++) {
      var d = values[vi][0];
      if (!(d instanceof Date)) continue;
      if (filterVal) {
        var fm = Utilities.formatDate(d, tz, "yyyy/MM");
        if (fm !== filterVal) continue;
      }
      var monthlyName = String(values[vi][7] || "").trim();
      if (monthlyName && /^\d{5,}$/.test(monthlyName)) {
        var resolved = resolveMonthlySheet_(ss, monthlyName);
        monthlyName = resolved ? resolved.getName() : "";
      }
      out.push([
        values[vi][6] === "" || values[vi][6] == null ? "" : values[vi][6],
        d,
        values[vi][2],
        values[vi][3],
        values[vi][4],
        monthlyName,
        values[vi][8] || "アプリ"
      ]);
    }
  }
  out.sort(function(a, b) {
    var na = parseInt(a[0], 10);
    var nb = parseInt(b[0], 10);
    if (isNaN(na)) na = 999999;
    if (isNaN(nb)) nb = 999999;
    return na - nb;
  });

  var clearRows = Math.max(uiSheet.getLastRow() - 2, out.length, 0);
  if (clearRows > 0) {
    uiSheet.getRange(3, 1, clearRows, 7).clearContent();
  }
  if (out.length > 0) {
    uiSheet.getRange(3, 1, out.length, 7).setValues(out);
    applyAppPaymentUiDataRowStyle_(uiSheet, out.length);
  }
}

/** App入金シートの1〜2行目・列幅（データ行は applyAppPaymentUiDataRowStyle_） */
function applyAppPaymentUiChrome_(uiSheet) {
  if (!uiSheet) return;
  var mono = {
    headerBg: "#3d3d3d",
    headerText: "#ffffff",
    totalBg: "#f0f0f0",
    border: "#9e9e9e",
    amountIncl: "#212121",
    amountExcl: "#616161"
  };
  uiSheet.getRange("1:1").setBackground(mono.totalBg).setFontColor("#333333");
  uiSheet
    .getRange("A1")
    .setBackground("#ffffff")
    .setFontWeight("bold")
    .setBorder(true, true, true, true, true, true, mono.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  uiSheet.getRange("C1").setValue("【 表示分合計 】").setFontWeight("bold").setHorizontalAlignment("right");
  uiSheet
    .getRange("D1")
    .setNumberFormat('"税込" #,##0')
    .setFontWeight("bold")
    .setFontColor(mono.amountIncl)
    .setFontSize(11);
  uiSheet
    .getRange("E1")
    .setNumberFormat('"税抜" #,##0')
    .setFontWeight("bold")
    .setFontColor(mono.amountExcl)
    .setFontSize(11);
  var headers = [["行", "対象年月", "氏名", "金額(税込)", "金額(税抜)", "月次シート", "経路"]];
  uiSheet
    .getRange(2, 1, 1, 7)
    .setValues(headers)
    .setFontWeight("bold")
    .setBackground(mono.headerBg)
    .setFontColor(mono.headerText)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, mono.border, SpreadsheetApp.BorderStyle.SOLID);
  uiSheet.setFrozenRows(2);
  uiSheet.setColumnWidth(1, 48);
  uiSheet.setColumnWidth(2, 92);
  uiSheet.setColumnWidth(3, 160);
  uiSheet.setColumnWidth(4, 110);
  uiSheet.setColumnWidth(5, 110);
  uiSheet.setColumnWidth(6, 88);
  uiSheet.setColumnWidth(7, 72);
}

/** App入金シートのデータ行の見た目（ゼブラ行・中央揃え） */
function applyAppPaymentUiDataRowStyle_(uiSheet, dataRowCount) {
  if (!uiSheet || dataRowCount <= 0) return;
  var zebra = ["#ffffff", "#f5f7fa"];
  var bgs = [];
  for (var i = 0; i < dataRowCount; i++) {
    bgs.push(new Array(7).fill(zebra[i % 2]));
  }
  uiSheet
    .getRange(3, 1, dataRowCount, 7)
    .setBackgrounds(bgs)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

/** A1の年月を変えたあと一覧だけ更新したいとき用 */
function refreshAppPaymentUiFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiSheet = ss.getSheetByName("App入金");
  var dataSheet = ss.getSheetByName(JOYFIT_APP_DATA_SHEET_NAME_);
  if (!uiSheet || !dataSheet) {
    Browser.msgBox("App入金 または App入金_Data シートがありません。");
    return;
  }
  refreshAppPaymentUiDisplay_(ss, uiSheet, dataSheet);
  Browser.msgBox("App入金の一覧を更新しました（見出し・色分けも反映）。");
}

function setupUISheetV5(ss, uiSheetName, dataSheetName) {
  let uiSheet = ss.getSheetByName(uiSheetName);
  if (!uiSheet) uiSheet = ss.insertSheet(uiSheetName);
  uiSheet.clear();

  applyAppPaymentUiChrome_(uiSheet);
  uiSheet.getRange("D1").setFormula("=SUBTOTAL(9, D3:D)");
  uiSheet.getRange("E1").setFormula("=SUBTOTAL(9, E3:E)");

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
      const filterCell = uiSheet.getRange("A1");
      filterCell.setDataValidation(rule);
      filterCell.setNumberFormat("@");
      if (filterCell.getValue() === "") filterCell.setValue(dateList[0]);
      filterCell.setNote(
        "【対象年月】で絞り込みます（A1）。\n" +
          "変更後は「入金メールを反映」または「その他」→「App入金一覧だけ更新」を実行してください。\n" +
          "A列=月次シートの行番号（17行目なら17）。"
      );
    }
  }

  refreshAppPaymentUiDisplay_(ss, uiSheet, dataSheet);
  uiSheet.getRange("A3:A").setNumberFormat("0");
  uiSheet.getRange("B3:B").setNumberFormat("yyyy/mm");
  uiSheet.getRange("D3:D").setNumberFormat("#,##0");
  uiSheet.getRange("E3:E").setNumberFormat("#,##0");
}

/** メール本文の「合計 税込○○円」を取得 */
function parseJoyfitEmailTotalIncl_(body) {
  const m = String(body || "").match(/合計[\s\S]*?税込\s*([\d,]+)\s*円/);
  if (!m) return 0;
  return parseInt(m[1].replace(/,/g, ""), 10) || 0;
}

/**
 * JOYFITアプリ入金メールを解析
 * 1通に複数月（3月+4月+5月など）がある場合 → 対象年月ごとに1行（各月の合計金額）
 * 例: 城下さん 合計31,290円 = 3行×10,430円（未納の各月分）
 */
function parseJoyfitEmailV5(body, receiveDate, messageId, taxRate) {
  const results = [];
  body = String(body || "");
  const nameMatch = body.match(/^(.*?)[\s　]*様/m);
  const name = nameMatch ? nameMatch[1].trim() : "不明";
  const emailTotal = parseJoyfitEmailTotalIncl_(body);
  const groupedData = new Map();

  const lines = body.split(/\r?\n/);
  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    if (line.indexOf("税込") === -1 || line.indexOf("円") === -1) continue;

    const monthMatch = line.match(/[\(（](\d{4})年(\d{1,2})月[\)）]/);
    const amountMatch = line.match(/税込\s*([\d,]+)\s*円/);
    if (!monthMatch || !amountMatch) continue;

    const yy = parseInt(monthMatch[1], 10);
    const mm = parseInt(monthMatch[2], 10);
    const amount = parseInt(amountMatch[1].replace(/,/g, ""), 10) || 0;
    if (!amount) continue;

    const targetDateObj = new Date(yy, mm - 1, 1);
    const key = targetDateObj.getTime();
    if (!groupedData.has(key)) {
      groupedData.set(key, { dateObj: targetDateObj, totalAmount: 0 });
    }
    groupedData.get(key).totalAmount += amount;
  }

  if (groupedData.size > 0) {
    let sumParsed = 0;
    groupedData.forEach(function(data) {
      sumParsed += data.totalAmount;
    });
    if (emailTotal > 0 && sumParsed !== emailTotal) {
      Logger.log(
        "App入金メール: 明細合計(" +
          sumParsed +
          ")とメール合計(" +
          emailTotal +
          ")が不一致: " +
          name +
          " / " +
          messageId
      );
    }

    groupedData.forEach(function(data) {
      results.push([
        data.dateObj,
        receiveDate,
        name,
        data.totalAmount,
        Math.round(data.totalAmount / taxRate),
        messageId
      ]);
    });
    return results;
  }

  if (emailTotal > 0) {
    const rd = new Date(receiveDate);
    const targetDateObj = new Date(rd.getFullYear(), rd.getMonth(), 1);
    results.push([
      targetDateObj,
      receiveDate,
      name,
      emailTotal,
      Math.round(emailTotal / taxRate),
      messageId
    ]);
  }

  return results;
}
