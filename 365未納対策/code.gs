/**
 * FIT365未納金回収（入金メール集計・SMS CSV 出力）
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('FIT365 請求データ管理')
    .addItem('📩 入金データを更新する', 'updateFit365Payments')
    .addItem('📥 SMSリストをCSV保存', 'downloadSmsCsv')
    .addToUi();
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
  if (lastRow < 13) return "";

  // 13行目以降の F列(5)〜I列(8) を取得
  const data = targetSheet.getRange(13, 6, lastRow - 12, 4).getValues();
  let csvContent = "";

  for (let i = 0; i < data.length; i++) { 
    const currentGroup = String(data[i][0]).trim(); // F列: ステータス
    const name = String(data[i][2]).trim();         // H列: 氏名
    let phone = String(data[i][3]).trim();          // I列: 電話番号
    
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
    dataSheet.getRange(nextRow, 1, newRecords.length, 7).setValues(newRecords);
    dataSheet.getRange("A:A").setNumberFormat("yyyy/mm/dd"); 
    dataSheet.getRange("B:B").setNumberFormat("yyyy/mm/dd HH:mm");
  }

  setupUISheetV5(ss, UI_SHEET_NAME, DATA_SHEET_NAME, COLOR_BG_HEADER, COLOR_TEXT_HEADER, COLOR_BG_TOTAL);
  const updatedCount = reconcileMonthlySheets(ss, DATA_SHEET_NAME);

  let msg = `✅ 更新完了\n\nデータを最新の状態に再構築しました。\n\n・取得件数: ${newRecords.length}件\n・消し込み更新: ${updatedCount}件`;
  Browser.msgBox(msg);
}

function reconcileMonthlySheets(ss, dataSheetName) {
  return 0; 
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
  uiSheet.getRange(2, 1, 1, 6).setValues(headers)
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