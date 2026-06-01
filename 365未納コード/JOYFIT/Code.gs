/**
 * JOYFIT未納金回収メール自動集計スクリプト (Ver 5.5) ＋ Excel(CSV)取り込み機能統合版
 *
 * ローカル配置: このフォルダ（JOYFIT/）を JOYFIT 用スプレッドシートの Apps Script にコピーして使う。
 * FIT365 用の別プロジェクトは同階層の FIT365/ フォルダ。
 */

/**
 * スプレッドシートを開いた時にメニューを追加する関数（統合版）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('請求データ管理')
    .addItem('入金データを更新', 'updateJoyfitPayments')
    .addItem('SMSリストをCSV保存', 'downloadSmsCsv')
    .addSeparator()
    .addItem('未納者CSVデータを取り込み', 'showImportDialog')
    .addToUi();
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
    if (k.indexOf("クレカ") !== -1 && (low.indexOf("jaccs") !== -1 || k.indexOf("ジャックス") !== -1)) return c + 1;
  }
  return 2;
}

/**
 * JOYFIT月次シートの列レイアウトを見出しから解決
 * 必須: 会員番号, 電話番号
 * 任意: 会員名 / 氏名、JACCS・クレカ列
 * 返却: { headerRow, dataStartRow, colGroup, colPayType, colMemberId, colMemberName, colPhone }
 * 見つからない場合は旧固定にフォールバック
 */
function resolveJoyfitSheetLayout_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var scanLastRow = Math.min(sheet.getLastRow(), 10);
  for (var r = 1; r <= scanLastRow; r++) {
    var rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var map = buildJoyfitHeaderColumnMap_(rowVals);
    var idCol = map[normalizeJoyfitHeaderLabel_("会員番号")] || 0;
    var phoneCol = map[normalizeJoyfitHeaderLabel_("電話番号")] || 0;
    var nameCol =
      map[normalizeJoyfitHeaderLabel_("会員名")] ||
      map[normalizeJoyfitHeaderLabel_("氏名")] ||
      0;
    var payCol = findJoyfitPayTypeColumnFromRowVals_(rowVals);
    if (idCol && phoneCol) {
      return {
        headerRow: r,
        dataStartRow: r + 1,
        colGroup: 1,
        colPayType: payCol || 2,
        colMemberId: idCol,
        colMemberName: nameCol || 4,
        colPhone: phoneCol
      };
    }
  }
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
    var data = s.getDataRange().getValues();
    var layout = resolveJoyfitSheetLayout_(s);
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
 * =========================================================================
 * 【追加機能】 ダイアログからデータを受け取り、既存シートへ書き込む処理
 * =========================================================================
 */
function processUnpaidData(csvData, targetSheetName, targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  targetSheetName = String(targetSheetName || "").trim();
  if (!targetSheetName) {
    targetSheetName = ss.getActiveSheet().getName();
  }
  var sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("指定されたシート「" + targetSheetName + "」が見つかりません。");
  }
  if (!targetMonth) {
    var now = new Date();
    targetMonth = String(now.getFullYear()).slice(-2) + "年" + (now.getMonth() + 1) + "月";
  }

  // 過去シートから会員番号 -> 電話番号を収集（値のみ利用）
  var phoneMap = buildJoyfitHistoricalPhoneMap_(ss, targetSheetName);

  // 名寄せ（値計算のみ）
  var members = new Map();
  for (var i = 1; i < csvData.length; i++) {
    var row = csvData[i];
    if (row.length < 15) continue;

    var memberId = String(row[3] || "").replace(/^'/, "").trim();
    var memberName = String(row[4] || "").trim();
    if (!memberId || !memberName) continue;

    var rawStatus = String(row[15] || "").trim();
    var isWithdrawn = rawStatus.includes("退会");
    var monthlyFee = parseInt(row[10], 10) || 0;
    var currentPayment = parseInt(row[14], 10) || 0;
    var rawReason = String(row[18] || row[row.length - 1] || "").trim();
    var reason = "その他";

    if (rawReason === "") reason = "";
    else if (rawReason.includes("使用不可")) reason = "使用不可";
    else if (rawReason.includes("取扱不可")) reason = "取扱不可";
    else if (rawReason.includes("有効期限")) reason = "有効期限切れ";
    else if (rawReason.includes("限度額") || rawReason.includes("限度超過")) reason = "利用限度額超過";

    if (!members.has(memberId)) {
      members.set(memberId, {
        name: memberName,
        reason: reason,
        monthlyFee: monthlyFee,
        currentPayment: currentPayment,
        isWithdrawn: isWithdrawn
      });
    } else {
      var m = members.get(memberId);
      m.monthlyFee += monthlyFee;
      m.currentPayment += currentPayment;
      if (isWithdrawn) m.isWithdrawn = true;
    }
  }

  var membersArray = [];
  members.forEach(function(data, id) {
    var group = 3;
    if (data.isWithdrawn) group = 1;
    else if (data.currentPayment > data.monthlyFee) group = 2;

    var carry = 0;
    var monthly = 0;
    if (group === 1) {
      carry = data.currentPayment;
      monthly = 0;
    } else if (group === 2) {
      monthly = Math.floor(data.currentPayment / 2);
      carry = data.currentPayment - monthly;
    } else {
      carry = 0;
      monthly = data.currentPayment;
    }

    membersArray.push({
      id: id,
      name: data.name,
      reason: data.reason,
      carry: carry,
      monthly: monthly,
      totalPayment: data.currentPayment,
      group: group
    });
  });

  // 並び順のみ維持
  membersArray.sort(function(a, b) {
    if (a.group !== b.group) return a.group - b.group;
    return b.totalPayment - a.totalPayment;
  });

  // ここから先は「値だけ」更新する。背景色・枠線・入力規則は変更しない。
  var startRow = 3; // 実データ開始行（1行ズレ解消）
  var outputData = []; // B〜L の11列を書き込む（B=区分、Cから実データ）
  var groupCounts = { 1: 0, 2: 0, 3: 0 };
  var colorRows = []; // A〜Z の色を自動反映

  membersArray.forEach(function(item) {
    groupCounts[item.group]++;
    var matchedPhone = phoneMap.get(item.id) || "";
    var bgColor = "#d9ead3";
    if (item.group === 1) bgColor = (groupCounts[1] % 2 !== 0) ? "#c9daf8" : "#a4c2f4";
    else if (item.group === 2) bgColor = (groupCounts[2] % 2 !== 0) ? "#fff2cc" : "#ffe599";
    else bgColor = (groupCounts[3] % 2 !== 0) ? "#d9ead3" : "#b6d7a8";
    var payment = (item.carry || 0) + (item.monthly || 0);
    var fee = "";
    if (payment >= 30000) fee = 300;
    else if (payment >= 10000) fee = 250;
    else if (payment >= 3000) fee = 200;
    else if (payment >= 2000) fee = 170;
    else if (payment >= 200) fee = 150;
    var total = payment ? payment + (fee || 0) : "";

    outputData.push([
      "クレカ",           // B: 区分
      item.id,            // C: 会員番号
      item.name,          // D: 会員名
      matchedPhone,       // E: 電話番号
      targetMonth,        // F: 未納対象月
      item.reason,        // G: 未納理由
      item.carry || "",   // H
      item.monthly || "", // I
      payment || "",      // J
      fee,                // K
      total               // L
    ]);
    colorRows.push(new Array(26).fill(bgColor)); // A〜Z
  });

  // 既存レイアウトを保持しつつ、B〜Lの値だけクリアして再入力
  var clearLastRow = Math.max(sheet.getLastRow(), startRow);
  if (clearLastRow >= startRow) {
    sheet.getRange(startRow, 2, clearLastRow - startRow + 1, 11).clearContent();
  }

  if (outputData.length > 0) {
    if (sheet.getMaxRows() < (startRow + outputData.length)) {
      sheet.insertRowsAfter(sheet.getMaxRows(), startRow + outputData.length - sheet.getMaxRows());
    }
    var inputRange = sheet.getRange(startRow, 2, outputData.length, 11);
    inputRange.setValues(outputData);
    inputRange.setHorizontalAlignment("left");
    sheet.getRange(startRow, 1, outputData.length, 26).setBackgrounds(colorRows); // A〜Zまで色反映
  }

  // 下部の空き行をカット（データ終端で切る）
  var END_ROW_MIN = 65; // 65行目を最低終端として扱う
  var lastDataRow = outputData.length > 0 ? (startRow + outputData.length - 1) : startRow;
  var finalRow = Math.max(END_ROW_MIN, lastDataRow);
  if (sheet.getMaxRows() > finalRow) {
    sheet.deleteRows(finalRow + 1, sheet.getMaxRows() - finalRow);
  }

  // A列の見出しを結合セルで作成（色・枠線は変更しない）
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
    g1.setBackground("#a4c2f4");
    g1.setFontSize(12).setFontWeight("bold");
    cursor += groupCounts[1];
  }
  if (groupCounts[2] > 0) {
    var g2 = sheet.getRange(cursor, 1, groupCounts[2], 1);
    g2.setValue("2ヶ月未納");
    if (groupCounts[2] > 1) g2.merge();
    g2.setBackground("#ffe599");
    g2.setFontSize(12).setFontWeight("bold");
    cursor += groupCounts[2];
  }
  if (groupCounts[3] > 0) {
    var g3 = sheet.getRange(cursor, 1, groupCounts[3], 1);
    g3.setValue("1カ月未納");
    if (groupCounts[3] > 1) g3.merge();
    g3.setBackground("#b6d7a8");
    g3.setFontSize(12).setFontWeight("bold");
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
  
  const COLOR_BG_HEADER = '#c8102e';
  const COLOR_TEXT_HEADER = '#ffffff';
  const COLOR_BG_TOTAL = '#f5f5f5';

  const QUERY = 'from:info@joyfit-service.jp subject:"未納金のお支払いについて" newer_than:6m';

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
  if (!label) label = GmailApp.createLabel(LABEL_NAME);

  const threads = GmailApp.search(QUERY, 0, 200);
  const newRecords = [];

  if (threads.length > 0) {
    for (const thread of threads) {
      const messages = thread.getMessages();
      let hasJoyfitMail = false;
      
      for (const message of messages) {
        if (message.getFrom().indexOf('joyfit-service.jp') === -1) continue;
        hasJoyfitMail = true;
        const records = parseJoyfitEmailV5(message.getPlainBody(), message.getDate(), message.getId(), TAX_RATE);
        if (records.length > 0) records.forEach(r => newRecords.push(r));
      }
      if (hasJoyfitMail) thread.addLabel(label);
    }
  }

  if (newRecords.length > 0) {
    newRecords.sort((a, b) => a[1] - b[1]);
    dataSheet.getRange(2, 1, newRecords.length, 7).setValues(newRecords);
    dataSheet.getRange("A:A").setNumberFormat("yyyy/mm/dd");
    dataSheet.getRange("B:B").setNumberFormat("yyyy/mm/dd HH:mm");
  }

  setupUISheetV5(ss, UI_SHEET_NAME, DATA_SHEET_NAME, COLOR_BG_HEADER, COLOR_TEXT_HEADER, COLOR_BG_TOTAL);
  const updatedCount = reconcileMonthlySheets(ss, DATA_SHEET_NAME);

  let msg = `✅ 更新完了\n\nデータを最新の状態に再構築しました。\n\n・取得件数: ${newRecords.length}件\n・消し込み更新: ${updatedCount}件`;
  msg += `\n\n「${UI_SHEET_NAME}」シートをご確認ください。`;

  if (newRecords.length > 0 || updatedCount > 0) Browser.msgBox(msg);
  else Browser.msgBox("対象期間内のメールが見つかりませんでした。");
}

function reconcileMonthlySheets(ss, dataSheetName) {
  const dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) return 0;
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = dataSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const monthNameMap = new Map();

  values.forEach(row => {
    const date = row[0];
    const name = row[2];
    if (date instanceof Date && name) {
      const yearStr = String(date.getFullYear()).slice(-2);
      const monthStr = String(date.getMonth() + 1);
      const monthKey = `${yearStr}-${monthStr}`;
      if (!monthNameMap.has(monthKey)) monthNameMap.set(monthKey, new Set());
      monthNameMap.get(monthKey).add(String(name).trim());
    }
  });

  let totalReconciled = 0;
  monthNameMap.forEach((namesSet, monthKey) => {
    const p = monthKey.split("-");
    const yy = p[0];
    const mm = p[1];
    const candidateNames = [`${yy}年${mm}月`, `${yy}年${mm}年`];
    let targetSheet = null;
    for (let ci = 0; ci < candidateNames.length; ci++) {
      targetSheet = ss.getSheetByName(candidateNames[ci]);
      if (targetSheet) break;
    }
    if (!targetSheet) return;
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < 2) return;

    const targetNames = targetSheet.getRange(1, 4, targetLastRow, 1).getValues().flat();
    const styleRanges = [];
    const checkRanges = [];
    const methodRanges = [];

    for (let i = 0; i < targetNames.length; i++) {
      const rowName = String(targetNames[i]).trim();
      if (rowName !== "" && namesSet.has(rowName)) {
        const rowNum = i + 1;
        styleRanges.push(`B${rowNum}:U${rowNum}`);
        checkRanges.push(`W${rowNum}`);
        methodRanges.push(`X${rowNum}`);
      }
    }

    if (styleRanges.length > 0) {
      const rangeList = targetSheet.getRangeList(styleRanges);
      rangeList.setBackground('#7f7f7f');
      rangeList.setFontColor('#ffffff');
      rangeList.setFontLine('line-through');

      const ruleW = SpreadsheetApp.newDataValidation().requireValueInList(["レジ入金", "アプリ入金"], true).setAllowInvalid(false).build();
      for (let k = 0; k < checkRanges.length; k++) {
        const vCell = targetSheet.getRange(checkRanges[k]);
        vCell.clearDataValidations();
        vCell.insertCheckboxes();
        vCell.setValue(true);

        const wCell = targetSheet.getRange(methodRanges[k]);
        wCell.setDataValidation(ruleW);
        wCell.setValue("アプリ入金");
      }
      totalReconciled += styleRanges.length;
    }
  });

  return totalReconciled;
}

function setupUISheetV5(ss, uiSheetName, dataSheetName, headerBg, headerText, totalBg) {
  let uiSheet = ss.getSheetByName(uiSheetName);
  if (!uiSheet) uiSheet = ss.insertSheet(uiSheetName);
  uiSheet.clear();

  uiSheet.getRange("1:1").setBackground(totalBg).setFontColor("#000000");
  uiSheet.getRange("A1").setBackground("#ffffff").setFontWeight("bold").setBorder(true, true, true, true, true, true, "#c8102e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  uiSheet.getRange("B1").setValue("【 表示分合計 】").setFontWeight("bold").setHorizontalAlignment("right");
  uiSheet.getRange("E1").setFormula("=SUBTOTAL(9, E3:E)");
  uiSheet.getRange("F1").setFormula("=SUBTOTAL(9, F3:F)");
  uiSheet.getRange("E1").setNumberFormat('"税込" #,##0').setFontWeight("bold").setFontColor("#c8102e").setFontSize(11);
  uiSheet.getRange("F1").setNumberFormat('"税抜" #,##0').setFontWeight("bold").setFontColor("#555555").setFontSize(11);

  const headers = [['対象年月', '受信日時', '氏名', '内訳項目', '金額(税込)', '金額(税抜)']];
  uiSheet.getRange(2, 1, 1, 6).setValues(headers)
    .setFontWeight('bold')
    .setBackground(headerBg)
    .setFontColor(headerText)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
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
      uiSheet.getRange("A1").setDataValidation(rule);
      uiSheet.getRange("A1").setNumberFormat("@");
      if (uiSheet.getRange("A1").getValue() === "") uiSheet.getRange("A1").setValue(dateList[0]);
    }
  }

  const queryFormula =
    `=QUERY('${dataSheetName}'!A2:F, ` +
    `"SELECT A, B, C, D, E, F WHERE A IS NOT NULL " & ` +
    `IF(A1="", "", " AND YEAR(A) = " & VALUE(LEFT(A1, 4)) & " AND MONTH(A) = " & (VALUE(RIGHT(A1, 2)) - 1)) & ` +
    `" ORDER BY A DESC, B DESC", 0)`;
  uiSheet.getRange("A3").setFormula(queryFormula);
  uiSheet.getRange("A3:A").setNumberFormat("yyyy/mm");
  uiSheet.getRange("B3:B").setNumberFormat("yyyy/mm/dd HH:mm");
  uiSheet.getRange("E3:F").setNumberFormat("#,##0");
  uiSheet.setColumnWidth(3, 150);
  uiSheet.setColumnWidth(4, 350);
}

function parseJoyfitEmailV5(body, receiveDate, messageId, taxRate) {
  const results = [];
  const nameRegex = /^(.*?)[\s　]*様/m;
  const nameMatch = body.match(nameRegex);
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
    if (dateMatch) targetDateObj = new Date(dateMatch[1].replace("年", "/").replace("月", "/1"));
    else { targetDateObj = new Date(receiveDate); targetDateObj.setDate(1); }
    targetDateObj.setHours(0, 0, 0, 0);

    const key = targetDateObj.getTime();
    if (!groupedData.has(key)) groupedData.set(key, { dateObj: targetDateObj, items: [], totalAmount: 0 });
    const data = groupedData.get(key);
    data.items.push(fullItemName);
    data.totalAmount += amount;
  }

  if (found) {
    groupedData.forEach((data) => {
      results.push([
        data.dateObj,
        receiveDate,
        name,
        data.items.join('、'),
        data.totalAmount,
        Math.round(data.totalAmount / taxRate),
        messageId
      ]);
    });
  }

  if (!found) {
    const totalMatch = body.match(/合計.*?税込([\d,]+)円/);
    if (totalMatch) {
      const amount = parseInt(totalMatch[1].replace(/,/g, ''), 10);
      const amountEx = Math.round(amount / taxRate);
      let targetDateObj;
      const dateMatchG = body.match(/\((\d{4}年\d{1,2}月)\)/);
      if (dateMatchG) targetDateObj = new Date(dateMatchG[1].replace("年", "/").replace("月", "/1"));
      else { targetDateObj = new Date(receiveDate); targetDateObj.setDate(1); }
      targetDateObj.setHours(0, 0, 0, 0);

      results.push([
        targetDateObj,
        receiveDate,
        name,
        "明細不明（合計のみ抽出）",
        amount,
        amountEx,
        messageId
      ]);
    }
  }

  return results;
}
