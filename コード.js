// Copyright Mitsunori KIMURA

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('名簿システム')
    .addItem('1. CSVから名簿・PDF作成', 'openCsvDialog')
    .addItem('2. メールの確認・一括送信', 'openEmailDialog')
    .addItem('3. ルーム・オリエン割り振り表', 'openAllocationDialog')
    .addItem('4. ビジター情報サマリー', 'openVisitorSummaryDialog')
    .addItem('5. 作成済みPDFの確認', 'openPdfLinksDialog')
    .addSeparator()
    .addItem('⚙️ メンバーブック(PDF)の更新', 'openMemberBookDialog')
    .addItem('⚙️ メンバーリスト(OCR)の更新', 'openPdfDialog')
    .addItem('⚙️ 休会日の管理', 'openHolidayDialog')
    .addItem('⚙️ メールテンプレート設定', 'openTemplateDialog')
    .addItem('⚙️ 割り振り表の特記事項設定', 'openAllocationNoteDialog')
    .addItem('⚙️ ビジターホストの設定', 'openVisitorHostDialog')
    .addItem('⚙️ Gemini API・モデル設定', 'openApiSettingsDialog')
    .addItem('⚙️ Webアプリ(送信元)設定', 'openWebAppSettingsDialog')
    .addSeparator()
    .addItem('📖 ご利用マニュアル', 'openManualDialog')
    .addToUi();
}

function openCsvDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('dialog').evaluate().setWidth(1000).setHeight(700), 'データの確認・PDF作成'); }
function openHolidayDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('holiday').setWidth(450).setHeight(400), '休会日の管理'); }
function openPdfDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('pdf').setWidth(450).setHeight(250), 'メンバーリスト(OCR)登録'); }
function openMemberBookDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('memberbook').setWidth(450).setHeight(250), 'メンバーブックの登録・更新'); }
function openTemplateDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('template').setWidth(600).setHeight(750), 'メールテンプレート設定'); }
function openAllocationNoteDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('allocation_note').setWidth(500).setHeight(400), '割り振り表の特記事項設定'); }
function openEmailDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('email').setWidth(800).setHeight(650), 'メールの確認・送信'); }
function openAllocationDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('allocation').setWidth(1000).setHeight(750), '割り振り表の作成'); }
function openVisitorHostDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('visitor_host').setWidth(400).setHeight(500), 'ビジターホストの設定'); }
function openApiSettingsDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('api_settings').setWidth(450).setHeight(350), 'Gemini API・モデル設定'); }
function openPdfLinksDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('pdf_links').setWidth(450).setHeight(400), '作成済みPDFの確認'); }
function openWebAppSettingsDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('webapp_settings').setWidth(500).setHeight(450), 'Webアプリ(送信元)設定'); }
function openManualDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('manual_viewer').setWidth(800).setHeight(700), 'ご利用マニュアル'); }

function getPdfLinks() {
  var props = PropertiesService.getScriptProperties();
  return {
    visitorList: props.getProperty('LATEST_VISITOR_LIST_URL') || "", allocation: props.getProperty('LATEST_ALLOCATION_URL') || "",
    memberBook: props.getProperty('MEMBER_BOOK_URL') || "", date: props.getProperty('LATEST_MEETING_DATE') || "未設定"
  };
}

function openVisitorSummaryDialog() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('visitor_summary').setWidth(650).setHeight(600), 'ビジター情報サマリー'); }

function getMeetingDateList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(), list = [];
  var re = /^(\d{4})参加者$/;
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(re);
    if (m) list.push(m[1]); // "0325" etc.
  }
  // 降順ソート（新しい日付が先）
  list.sort(function(a, b) { return b.localeCompare(a); });
  var latest = PropertiesService.getScriptProperties().getProperty('LATEST_MEETING_DATE') || "";
  var latestMmdd = "";
  if (latest) {
    var d = new Date(latest);
    latestMmdd = Utilities.formatDate(d, "Asia/Tokyo", "MMdd");
  }
  return { dates: list, latest: latestMmdd };
}

function getVisitorSummaryData(mmddParam) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mmdd, dateObj;
  if (mmddParam) {
    // mmddParam is like "0325" — resolve to a full date by finding the sheet
    mmdd = mmddParam;
    var month = parseInt(mmdd.substring(0, 2), 10) - 1;
    var day = parseInt(mmdd.substring(2, 4), 10);
    // Guess the year based on proximity to current date
    var now = new Date();
    dateObj = new Date(now.getFullYear(), month, day);
    if (dateObj.getTime() - now.getTime() > 6 * 30 * 24 * 3600000) dateObj.setFullYear(now.getFullYear() - 1);
    else if (now.getTime() - dateObj.getTime() > 6 * 30 * 24 * 3600000) dateObj.setFullYear(now.getFullYear() + 1);
  } else {
    var props = PropertiesService.getScriptProperties();
    var meetingDateVal = props.getProperty('LATEST_MEETING_DATE');
    if (!meetingDateVal) return { error: "定例会データがまだ作成されていません。先にCSVから名簿を作成してください。" };
    dateObj = new Date(meetingDateVal);
    mmdd = Utilities.formatDate(dateObj, "Asia/Tokyo", "MMdd");
  }
  var sheetName = mmdd + "参加者", dataSheet = ss.getSheetByName(sheetName);
  if (!dataSheet) return { error: "シート「" + sheetName + "」が見つかりません。" };

  // 定例会回数を計算
  var targetDateStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy/MM/dd");
  var baseDate = new Date("2026/03/18 00:00:00"), baseCount = 509, holidays = getHolidays();
  var meetingCount = 0;
  // 対象日が基準日以降なら前方探索
  if (dateObj.getTime() >= baseDate.getTime()) {
    var currDate = new Date(baseDate.getTime()), currCount = baseCount;
    for (var limit = 0; limit < 500; limit++) {
      var dStr = Utilities.formatDate(currDate, "Asia/Tokyo", "yyyy/MM/dd");
      if (dStr === targetDateStr) { meetingCount = currCount; break; }
      if (holidays.indexOf(dStr) === -1) currCount++;
      currDate.setDate(currDate.getDate() + 7);
    }
  } else {
    // 対象日が基準日より前なら後方探索
    var currDate = new Date(baseDate.getTime()), currCount = baseCount;
    currDate.setDate(currDate.getDate() - 7);
    currCount--;
    for (var limit = 0; limit < 500; limit++) {
      var dStr = Utilities.formatDate(currDate, "Asia/Tokyo", "yyyy/MM/dd");
      if (holidays.indexOf(dStr) !== -1) {
        // 休会日はカウントしない
        currDate.setDate(currDate.getDate() - 7);
        continue;
      }
      if (dStr === targetDateStr) { meetingCount = currCount; break; }
      currCount--;
      currDate.setDate(currDate.getDate() - 7);
    }
  }

  var data = dataSheet.getDataRange().getValues(), headers = data[0];
  var noIdx = headers.indexOf("No."), nameIdx = headers.indexOf("参加者氏名"), kanaIdx = headers.indexOf("ふりがな");
  var catIdx = headers.indexOf("カテゴリー"), invIdx = headers.indexOf("招待者"), typeIdx = headers.indexOf("種別");
  var billedIdx = headers.indexOf("請求額"), paidAmtIdx = headers.indexOf("支払額");

  var visitors = [], guests = [], substitutes = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[noIdx]) continue;
    var billed = billedIdx !== -1 ? Number(row[billedIdx]) || 0 : 0;
    var paidAmt = paidAmtIdx !== -1 ? Number(row[paidAmtIdx]) || 0 : 0;
    var isPaid = billed > 0 ? paidAmt >= billed : true;  // 請求額なし→判定不能→入金済み扱い
    var entry = { no: String(row[noIdx]), name: row[nameIdx] || "", kana: row[kanaIdx] || "", cat: row[catIdx] || "", inviter: row[invIdx] || "", type: typeIdx !== -1 ? (row[typeIdx] || "") : "", paid: isPaid };
    if (entry.type === "Guest") guests.push(entry);
    else if (entry.type === "Substitute") substitutes.push(entry);
    else visitors.push(entry);
  }

  // 代理をメンバーリストの番号順にソート
  substitutes.sort(function(a, b) {
    var numA = parseInt(a.no.replace(/[^0-9]/g, ""), 10) || 9999;
    var numB = parseInt(b.no.replace(/[^0-9]/g, ""), 10) || 9999;
    return numA - numB;
  });

  // シートの最終更新日時を取得（DriveApp経由）
  var fileId = ss.getId(), file = DriveApp.getFileById(fileId);
  var lastUpdated = Utilities.formatDate(file.getLastUpdated(), "Asia/Tokyo", "M/d HH:mm");

  var md = Utilities.formatDate(dateObj, "Asia/Tokyo", "M/d");
  return {
    meetingCount: meetingCount, dateMd: md, lastUpdated: lastUpdated,
    visitors: visitors, guests: guests, substitutes: substitutes
  };
}

function getApiSettings() {
  var props = PropertiesService.getScriptProperties();
  return { apiKey: props.getProperty('GEMINI_API_KEY') || "", modelName: props.getProperty('GEMINI_MODEL_NAME') || "gemini-2.5-flash" };
}
function saveApiSettings(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('GEMINI_API_KEY', data.apiKey.trim()); props.setProperty('GEMINI_MODEL_NAME', data.modelName.trim());
  return "APIキーとモデルを保存しました。";
}

function getWebAppSettings() {
  var props = PropertiesService.getScriptProperties();
  return { webAppUrl: props.getProperty('WEB_APP_URL') || "", secretToken: props.getProperty('SECRET_TOKEN') || SECRET_TOKEN };
}
function saveWebAppSettings(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WEB_APP_URL', data.webAppUrl.trim());
  if (data.secretToken && data.secretToken.trim() !== "") props.setProperty('SECRET_TOKEN', data.secretToken.trim());
  return "Webアプリ設定を保存しました。";
}

function getAllocationNote() {
  var props = PropertiesService.getScriptProperties();
  var note = props.getProperty('ALLOCATION_NOTE');
  if (note === null) {
    note = "※ビジターホストの皆様へ（当日の動き）\n① 6:45〜 メインルームにて招待者へ本日のルーム分けをお伝えします。\n② 6:55〜 順次ブレイクアウトルームへご案内します。\n③ ミーティング終了後、オリエンテーションルームを作成します。";
  }
  return note;
}
function saveAllocationNote(text) {
  PropertiesService.getScriptProperties().setProperty('ALLOCATION_NOTE', text);
  return "保存しました。";
}

function getManualHtml() {
  return HtmlService.createHtmlOutputFromFile('manual_content').getContent();
}

function getHolidays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("休会日");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), holidays = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) { var d = new Date(data[i][0]); if (!isNaN(d.getTime())) holidays.push(Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd")); }
  }
  return holidays;
}
function saveHolidays(holidaysArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName("休会日");
  if (!sheet) { sheet = ss.insertSheet("休会日"); sheet.hideSheet(); } else { sheet.clear(); }
  var data = holidaysArray.map(function(h){ return [h]; });
  if(data.length > 0) sheet.getRange(1, 1, data.length, 1).setValues(data);
  return "休会日を保存しました";
}
function getMeetingCandidates() {
  var baseDate = new Date("2026/03/18 00:00:00"), baseCount = 509, holidays = getHolidays(), candidates = [], today = new Date(); today.setHours(0,0,0,0);
  var currDate = new Date(baseDate.getTime()), currCount = baseCount, found = 0, limit = 0;
  while(found < 4 && limit < 100) {
    var dateStr = Utilities.formatDate(currDate, "Asia/Tokyo", "yyyy/MM/dd");
    if (currDate >= today && holidays.indexOf(dateStr) === -1) { candidates.push({ dateValue: dateStr, display: Utilities.formatDate(currDate, "Asia/Tokyo", "yyyy/M/d") + "(水) 第" + currCount + "回" }); found++; }
    if (holidays.indexOf(dateStr) === -1) currCount++;
    currDate.setDate(currDate.getDate() + 7); limit++;
  }
  return candidates;
}

function zenkakuToHankaku(str) { return str ? str.toString().replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); }).replace(/　/g, ' ') : ""; }
function needsSpaceReview(str) {
  if (!str) return false;
  var trimmed = str.toString().trim();
  if (/[\s]/.test(trimmed) || /[a-zA-Z]/.test(trimmed) || trimmed.length < 2) return false;
  return true; 
}
function formatNameSuggest(str) {
  if (!str) return "";
  var trimmed = str.toString().trim();
  if (/[\s]/.test(trimmed)) return trimmed.replace(/[\s]+/g, ' ');
  if (/[a-zA-Z]/.test(trimmed)) return trimmed;
  var len = trimmed.length;
  if (len === 3 || len === 4 || len === 5) return trimmed.slice(0, 2) + " " + trimmed.slice(2);
  return trimmed;
}
function normalizeSpace(str) { return str ? str.toString().replace(/[\s]+/g, ' ').trim() : ""; }

// 異体字対応のファジー名前マッチ（邊/邉、齋/斎/齊 等の1文字違いを許容）
function fuzzyNameMatch(searchStr, memberStr) {
  if (searchStr === memberStr) return true;
  if (searchStr.indexOf(memberStr) !== -1 || memberStr.indexOf(searchStr) !== -1) return true;
  // 同じ長さで1文字だけ異なる場合を許容（異体字対応）
  if (searchStr.length === memberStr.length && searchStr.length >= 2) {
    var diff = 0;
    for (var i = 0; i < searchStr.length; i++) { if (searchStr[i] !== memberStr[i]) diff++; }
    if (diff <= 1) return true;
  }
  return false;
}

// メンバーリストから招待者名を照合し、正規化されたメンバー名を返す（見つからなければ元の名前を返す）
function matchInviterToMember(inviterName, membersList) {
  if (!inviterName) return "";
  var searchInv = String(inviterName).replace(/[\s\u3000さん]/g, "");
  if (searchInv === "") return "";
  for (var k = 0; k < membersList.length; k++) {
    var mNameSearch = membersList[k].name.replace(/[\s]/g, "");
    if (fuzzyNameMatch(searchInv, mNameSearch)) return membersList[k].name;
  }
  return String(inviterName);
}

function getMembersList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("メンバーリスト");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), list = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) list.push({ no: data[i][0].toString(), name: normalizeSpace(data[i][1].toString()) }); }
  return list;
}

function analyzeCsvData(csvText) {
  var data = Utilities.parseCsv(csvText);
  if (data.length < 2) throw new Error("データがありません");
  var header = data[0];
  for(var h = 0; h < header.length; h++) header[h] = zenkakuToHankaku(header[h]).trim();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('').trim() === '') continue;
    var obj = {};
    for (var j = 0; j < header.length; j++) obj[header[j]] = zenkakuToHankaku(data[i][j]);
    rows.push(obj);
  }
  var typeOrder = { "Visitor": 1, "Guest": 2, "Substitute": 3 };
  rows.sort(function(a, b) {
    var tA = typeOrder[a["種別"]] || 99, tB = typeOrder[b["種別"]] || 99;
    if (tA !== tB) return tA - tB;
    return (a["ふりがな"] || "").localeCompare((b["ふりがな"] || ""), 'ja');
  });
  var membersList = getMembersList(), vCount = 1, gCount = 1, results = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], origName = r["参加者氏名"] || "", origKana = r["ふりがな"] || "";
    r._needsNameReview = needsSpaceReview(origName) || needsSpaceReview(origKana);
    r["参加者氏名"] = formatNameSuggest(origName); r["ふりがな"] = formatNameSuggest(origKana);
    r._needsInviterReview = false;
    var originalInviter = r["招待者"] || "", searchInviter = originalInviter.replace(/[\sさん]/g, ""), matchedMember = null;
    if (searchInviter !== "") {
      for (var k = 0; k < membersList.length; k++) {
        var mNameSearch = membersList[k].name.replace(/[\s]/g, "");
        if (fuzzyNameMatch(searchInviter, mNameSearch)) { matchedMember = membersList[k]; break; }
      }
    }
    if (matchedMember) { r["招待者"] = matchedMember.name; } else if (originalInviter !== "") { r._needsInviterReview = true; r["招待者"] = formatNameSuggest(originalInviter); }
    if (r["種別"] === "Visitor") { r._No = "V" + ("0" + vCount).slice(-2); vCount++; } 
    else if (r["種別"] === "Guest") { r._No = "G" + ("0" + gCount).slice(-2); gCount++; } 
    else if (r["種別"] === "Substitute") { r._No = matchedMember ? "代理" + matchedMember.no : "代理??"; } 
    else { r._No = ""; }
    results.push(r);
  }
  return { rows: results, header: header, membersList: membersList };
}

function createFinalSheet(meetingDateVal, meetingDisplay, finalRows, originalHeader) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), dateObj = new Date(meetingDateVal), baseSheetName = Utilities.formatDate(dateObj, "Asia/Tokyo", "MMdd") + "参加者";
  var dataSheetName = baseSheetName, dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) dataSheet = ss.insertSheet(dataSheetName); else dataSheet.clear();
  var printSheetName = baseSheetName + "_印刷用", printSheet = ss.getSheetByName(printSheetName);
  if (!printSheet) printSheet = ss.insertSheet(printSheetName); else printSheet.clear();
  var fixedHeaders = ["No.", "参加者氏名", "ふりがな", "カテゴリー", "会社名", "招待者", "備考"], mapKeys = ["_No", "参加者氏名", "ふりがな", "カテゴリー", "会社名", "招待者", "メモ（ビジターリストに表示）"];
  var dataHeaders = fixedHeaders.slice(), otherKeys = [];
  for (var i = 0; i < originalHeader.length; i++) { if (mapKeys.indexOf(originalHeader[i]) === -1) { dataHeaders.push(originalHeader[i]); otherKeys.push(originalHeader[i]); } }
  if (dataHeaders.indexOf("種別") === -1) { dataHeaders.push("種別"); otherKeys.push("種別"); }
  if (dataHeaders.indexOf("メール") === -1) { dataHeaders.push("メール"); otherKeys.push("メール"); }
  var dataOutput = [dataHeaders], printOutput = [
    ["Activeチャプターの定例会へようこそ", "", "", "", "", "", ""], ["", "", "", "", "", "", ""], [meetingDisplay, "", "", "", "", "", ""], ["", "", "", "", "", "", ""], fixedHeaders
  ];
  for (var i = 0; i < finalRows.length; i++) {
    var r = finalRows[i], baseRow = [ r["_No"]||"", r["参加者氏名"]||"", r["ふりがな"]||"", r["カテゴリー"]||"", r["会社名"]||"", r["招待者"]||"", r["メモ（ビジターリストに表示）"]||"" ];
    printOutput.push(baseRow);
    var fullRow = baseRow.slice();
    for (var k = 0; k < otherKeys.length; k++) fullRow.push(r[otherKeys[k]] || "");
    dataOutput.push(fullRow);
  }
  dataSheet.getRange(1, 1, dataOutput.length, dataOutput[0].length).setValues(dataOutput);
  printSheet.getRange(1, 1, printOutput.length, printOutput[0].length).setValues(printOutput);
  printSheet.getRange("A1").setFontSize(16).setFontWeight("bold"); printSheet.getRange("A3").setFontSize(12).setFontWeight("bold");
  var lastPrintRow = printOutput.length;
  var printDataRange = printSheet.getRange(5, 1, lastPrintRow - 4, 7);
  printDataRange.setWrap(true).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
  printSheet.getRange(5, 1, 1, 7).setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center");
  if (lastPrintRow > 5) {
    printSheet.getRange(6, 1, lastPrintRow - 5, 7).setHorizontalAlignment("left"); 
    printSheet.getRange(6, 1, lastPrintRow - 5, 3).setHorizontalAlignment("center");
    printSheet.getRange(6, 6, lastPrintRow - 5, 1).setHorizontalAlignment("center");
  }
  var widths = [40, 100, 110, 160, 160, 100, 180];
  for(var w=0; w<widths.length; w++) printSheet.setColumnWidth(w+1, widths[w]);
  SpreadsheetApp.flush();
  var pdfUrl = exportSheetToPdf(printSheet, meetingDisplay + " ビジター様リスト.pdf", meetingDateVal);
  PropertiesService.getScriptProperties().setProperty('LATEST_VISITOR_LIST_URL', pdfUrl);
  PropertiesService.getScriptProperties().setProperty('LATEST_MEETING_DATE', meetingDateVal); 
  ss.setActiveSheet(dataSheet);
  return "<h3>処理が完了しました🎉</h3><p>PDFを作成し、全員が閲覧できるよう権限を付与しました。</p><br><a href='" + pdfUrl + "' target='_blank' style='background:#0055ff; color:#fff; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold;'>📄 作成されたPDFを開く</a>";
}

function exportSheetToPdf(sheet, fileName, meetingDateVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), spreadsheetId = ss.getId(), sheetId = sheet.getSheetId(), lastRow = sheet.getLastRow();
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=" + sheetId + "&r1=0&c1=0&r2=" + lastRow + "&c2=7";
  var token = ScriptApp.getOAuthToken(), response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
  var blob = response.getBlob().setName(fileName);
  var props = PropertiesService.getScriptProperties();
  // 開催日ごとにファイルIDを管理（同日なら上書き更新、別日なら新規作成）
  var mmdd = meetingDateVal ? Utilities.formatDate(new Date(meetingDateVal), "Asia/Tokyo", "MMdd") : "";
  var propKey = mmdd ? 'VISITOR_LIST_PDF_ID_' + mmdd : '';
  var fileId = propKey ? props.getProperty(propKey) : null;
  if (fileId) {
    try { Drive.Files.update({}, fileId, blob); return DriveApp.getFileById(fileId).getUrl(); } catch(e) { fileId = null; }
  }
  var file = DriveApp.getFileById(spreadsheetId), folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  var pdfFile = folder.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (propKey) props.setProperty(propKey, pdfFile.getId());
  return pdfFile.getUrl();
}

function uploadMemberBook(formObject) {
  var blob = formObject.pdfFile, props = PropertiesService.getScriptProperties(), fileId = props.getProperty('MEMBER_BOOK_ID');
  if (fileId) { try { Drive.Files.update({}, fileId, blob); return { msg: "更新しました。", url: props.getProperty('MEMBER_BOOK_URL') }; } catch(e) { fileId = null; } }
  if (!fileId) {
    var file = Drive.Files.create({name: 'MemberBook.pdf', mimeType: 'application/pdf'}, blob);
    DriveApp.getFileById(file.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('MEMBER_BOOK_ID', file.id);
    var url = DriveApp.getFileById(file.id).getUrl();
    props.setProperty('MEMBER_BOOK_URL', url);
    return { msg: "新規登録しました。", url: url };
  }
}

function processPdfForm(formObject) {
  var blob = formObject.pdfFile, file = Drive.Files.create({ name: blob.getName(), mimeType: 'application/vnd.google-apps.document' }, blob);
  var text = DocumentApp.openById(file.id).getBody().getText();
  Drive.Files.remove(file.id);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("メンバーリスト");
  if(!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("メンバーリスト");
  sheet.clear(); sheet.appendRow(["No", "氏名"]);
  var lines = text.split('\n'), members = [], no = null;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (/^\d{1,3}$/.test(line)) { no = line; }
    else if (no && line.length > 0) { var name = line.split('(')[0].split('（')[0].trim(); if(name) members.push([no, normalizeSpace(name)]); no = null; }
  }
  if (members.length > 0) sheet.getRange(2, 1, members.length, 2).setValues(members);
  return "抽出件数: " + members.length + "件。\nシートを確認してください。";
}

// === メール関連処理 ===
var SECRET_TOKEN = "ActiveChapterSecret2026";

function getWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || "";
}

function getTemplates() {
  var props = PropertiesService.getScriptProperties();
  return {
    cc: props.getProperty('MAIL_TPL_CC') || "", bcc: props.getProperty('MAIL_TPL_BCC') || "",
    visitorSubj: props.getProperty('MAIL_TPL_VISITOR_SUBJ') || "【Activeチャプター】{{date}} 定例会のご案内",
    visitorBody: props.getProperty('MAIL_TPL_VISITOR_BODY') || "{{name}} 様\n\nご参加ありがとうございます。\n\n定例会開催日: {{date}}\n\nビジターリスト:\n{{visitorlist}}\n\nメンバーブック:\n{{memberbook}}",
    guestSubj: props.getProperty('MAIL_TPL_GUEST_SUBJ') || "【Activeチャプター】{{date}} ゲスト様へのご案内",
    guestBody: props.getProperty('MAIL_TPL_GUEST_BODY') || "{{name}} 様\n\nご参加ありがとうございます。\n\n...",
    substituteSubj: props.getProperty('MAIL_TPL_SUBSTITUTE_SUBJ') || "【Activeチャプター】{{date}} 代理参加のご案内",
    substituteBody: props.getProperty('MAIL_TPL_SUBSTITUTE_BODY') || "{{name}} 様\n\n代理でのご参加ありがとうございます。\n\n..."
  };
}

function saveTemplates(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('MAIL_TPL_CC', data.cc); props.setProperty('MAIL_TPL_BCC', data.bcc);
  props.setProperty('MAIL_TPL_VISITOR_SUBJ', data.visitorSubj); props.setProperty('MAIL_TPL_VISITOR_BODY', data.visitorBody);
  props.setProperty('MAIL_TPL_GUEST_SUBJ', data.guestSubj); props.setProperty('MAIL_TPL_GUEST_BODY', data.guestBody);
  props.setProperty('MAIL_TPL_SUBSTITUTE_SUBJ', data.substituteSubj); props.setProperty('MAIL_TPL_SUBSTITUTE_BODY', data.substituteBody);
  return "テンプレートを保存しました。";
}

function generateEmailDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getActiveSheet(), data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("データがありません。作成された「〇〇参加者」シートを開いた状態で実行してください。");
  var headers = data[0];
  var nameIdx = headers.indexOf("参加者氏名"), emailIdx = headers.indexOf("メール"), typeIdx = headers.indexOf("種別"), inviterIdx = headers.indexOf("招待者");
  if (nameIdx === -1 || emailIdx === -1 || typeIdx === -1) throw new Error("現在のシートに必須列が見つかりません。");
  
  var tpls = getTemplates(), props = PropertiesService.getScriptProperties();
  var visitorListUrl = props.getProperty('LATEST_VISITOR_LIST_URL') || "【未生成】", memberBookUrl = props.getProperty('MEMBER_BOOK_URL') || "【未生成】";
  var rawDate = props.getProperty('LATEST_MEETING_DATE') || "", dateFormatted = "";
  if(rawDate) { var d = new Date(rawDate); dateFormatted = d.getFullYear() + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日"; }
  
  var drafts = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][nameIdx], email = data[i][emailIdx], type = data[i][typeIdx];
    var inviter = inviterIdx !== -1 ? data[i][inviterIdx] : "";
    
    if (!email || email.trim() === "") continue;
    var tplSubj = "", tplBody = "";
    if (type === "Visitor") { tplSubj = tpls.visitorSubj; tplBody = tpls.visitorBody; }
    else if (type === "Guest") { tplSubj = tpls.guestSubj; tplBody = tpls.guestBody; }
    else if (type === "Substitute") { tplSubj = tpls.substituteSubj; tplBody = tpls.substituteBody; }
    else continue;
    
    // {{inviter}} と、念のための {{invitee}} 両方で置換対応
    var subject = tplSubj.replace(/{{name}}/g, name).replace(/{{inviter}}/g, inviter).replace(/{{invitee}}/g, inviter).replace(/{{date}}/g, dateFormatted).replace(/{{visitorlist}}/g, visitorListUrl).replace(/{{memberbook}}/g, memberBookUrl);
    var body = tplBody.replace(/{{name}}/g, name).replace(/{{inviter}}/g, inviter).replace(/{{invitee}}/g, inviter).replace(/{{date}}/g, dateFormatted).replace(/{{visitorlist}}/g, visitorListUrl).replace(/{{memberbook}}/g, memberBookUrl);
    drafts.push({ name: name, email: email, type: type, subject: subject, body: body, send: true });
  }
  return { drafts: drafts, cc: tpls.cc, bcc: tpls.bcc };
}

function sendSingleEmail(e, cc, bcc) {
  try {
    var options = { name: "Activeチャプター" };
    if (cc && cc.trim() !== "") options.cc = cc.trim();
    if (bcc && bcc.trim() !== "") options.bcc = bcc.trim();
    var toEmail = e.email ? e.email.toString().trim() : "";
    if (!toEmail || toEmail.indexOf('@') === -1) return { success: false, error: "無効なメールアドレス形式 (" + toEmail + ")" };
    
    var webAppUrl = getWebAppUrl();
    if (webAppUrl === "") {
      GmailApp.sendEmail(toEmail, e.subject, e.body, options);
      return { success: true };
    } else {
      var token = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || SECRET_TOKEN;
      var payload = { token: token, to: toEmail, subject: e.subject, body: e.body, options: options };
      var res = UrlFetchApp.fetch(webAppUrl, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
      var resData = JSON.parse(res.getContentText());
      if (!resData.success) throw new Error(resData.error);
      return { success: true };
    }
  } catch (error) { return { success: false, error: error.toString() }; }
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || SECRET_TOKEN;
    if (params.token !== token) throw new Error("アクセス権限がありません");
    GmailApp.sendEmail(params.to, params.subject, params.body, params.options);
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function getVisitorHosts() {
  var props = PropertiesService.getScriptProperties(), hosts = props.getProperty('VISITOR_HOSTS');
  return hosts ? JSON.parse(hosts) : [];
}
function saveVisitorHosts(hostIds) { PropertiesService.getScriptProperties().setProperty('VISITOR_HOSTS', JSON.stringify(hostIds)); return "保存しました。"; }

function getMemberPriorities() {
  var props = PropertiesService.getScriptProperties(), data = props.getProperty('MEMBER_PRIORITIES');
  return data ? JSON.parse(data) : {};
}
function saveMemberPriorities(priorities) {
  PropertiesService.getScriptProperties().setProperty('MEMBER_PRIORITIES', JSON.stringify(priorities));
  return "保存しました。";
}

function getAllocationData(meetingDateVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), dateObj = new Date(meetingDateVal), mmdd = Utilities.formatDate(dateObj, "Asia/Tokyo", "MMdd");
  var sheetName = mmdd + "参加者", allocSheetName = mmdd + "割り振り表", dataSheet = ss.getSheetByName(sheetName);
  if(!dataSheet) throw new Error("対象のデータシートがありません。");
  var data = dataSheet.getDataRange().getValues(), headerRowIdx = -1;
  for(var i=0; i<Math.min(10, data.length); i++){ if(data[i].indexOf("No.") !== -1) { headerRowIdx = i; break; } }
  if(headerRowIdx === -1) throw new Error("シート形式が不正です。");
  var headers = data[headerRowIdx], noIdx = headers.indexOf("No."), nameIdx = headers.indexOf("参加者氏名"), catIdx = headers.indexOf("カテゴリー"), invIdx = headers.indexOf("招待者");

  var membersList = getMembersList();

  var visitors = [], inviters = {};
  for(var i = headerRowIdx + 1; i < data.length; i++) {
    var row = data[i];
    if(!row[noIdx]) continue;
    var detailsObj = {};
    for(var j=0; j<headers.length; j++) detailsObj[headers[j]] = row[j];
    // 招待者名をメンバーリストと照合して正規化（異体字対応: 邊/邉 等）
    var rawInviter = matchInviterToMember(row[invIdx], membersList);
    visitors.push({ no: String(row[noIdx]), name: row[nameIdx], cat: row[catIdx], inviter: rawInviter, details: detailsObj });
    if(rawInviter) inviters[normalizeSpace(rawInviter)] = true;
  }

  var pool = [];
  for(var i = 0; i < membersList.length; i++) {
     if(!inviters[normalizeSpace(membersList[i].name)]) pool.push(membersList[i]);
  }

  var facilAlloc = {}, orienAlloc = {}, roomAlloc = {}, connectReq = {}, mergedWith = {};
  var allocSheet = ss.getSheetByName(allocSheetName);
  
  if(allocSheet) {
     var aData = allocSheet.getDataRange().getValues(), aHeaderIdx = -1;
     for(var i=0; i<Math.min(10, aData.length); i++){ if(aData[i].indexOf("No.") !== -1) { aHeaderIdx = i; break; } }
     if(aHeaderIdx !== -1) {
         var aHeaders = aData[aHeaderIdx];
         var nameColIdx = aHeaders.indexOf("お名前") !== -1 ? aHeaders.indexOf("お名前") : 1;
         var facilIdx = aHeaders.indexOf("ファシリテーター"), roomIdx = aHeaders.indexOf("ルームメンバー"), orienIdx = aHeaders.indexOf("オリエンテーション"), connIdx = aHeaders.indexOf("つなげたいメンバー");
         
         for(var i = aHeaderIdx + 1; i < aData.length; i++) {
             var row = aData[i];
             if(!row[0] || String(row[0]).indexOf("※")===0) break; 
             
             var savedName = normalizeSpace(row[nameColIdx]);
             var matchedVisitor = visitors.filter(function(v) { return normalizeSpace(v.name) === savedName; })[0];
             if (!matchedVisitor) continue; 
             
             var vNo = String(matchedVisitor.no);
             connectReq[vNo] = connIdx !== -1 ? row[connIdx] : "";
             var fName = facilIdx !== -1 ? row[facilIdx] : ""; 
             
             if (fName && String(fName).indexOf("【合同】") === 0) {
                 var targetName = String(fName).replace("【合同】", "").replace("と同室", "").trim();
                 var targetV = visitors.filter(function(v) { return normalizeSpace(v.name) === targetName; })[0];
                 if(targetV) mergedWith[vNo] = String(targetV.no);
             } else if(fName) { 
                 var fm = pool.filter(function(x){ return normalizeSpace(x.name) === normalizeSpace(fName); })[0]; 
                 if(fm) facilAlloc[vNo] = String(fm.no); 
             }
             
             roomAlloc[vNo] = [];
             if (!mergedWith[vNo]) {
                 var rNamesStr = roomIdx !== -1 ? (row[roomIdx] ? row[roomIdx].toString() : "") : "";
                 if(rNamesStr) {
                     var rNames = rNamesStr.split("\n");
                     for(var j=0; j<rNames.length; j++) { 
                       var rmName = normalizeSpace(rNames[j]);
                       if(!rmName || rmName === "（同上）") continue;
                       var rm = pool.filter(function(x){ return normalizeSpace(x.name) === rmName; })[0]; 
                       if(rm) roomAlloc[vNo].push(String(rm.no)); 
                     }
                 }
             }
             
             // ファシリがルームメンバーに重複していたら除去
             if (facilAlloc[vNo] && roomAlloc[vNo]) {
                 roomAlloc[vNo] = roomAlloc[vNo].filter(function(id) { return String(id) !== String(facilAlloc[vNo]); });
             }

             orienAlloc[vNo] = [];
             var oNamesStr = orienIdx !== -1 ? (row[orienIdx] ? row[orienIdx].toString() : "") : "";
             if(oNamesStr) {
                 var oNames = oNamesStr.split("\n");
                 for(var j=0; j<oNames.length; j++) {
                   var omName = normalizeSpace(oNames[j]);
                   if(!omName) continue;
                   var om = pool.filter(function(x){ return normalizeSpace(x.name) === omName; })[0];
                   if(om) orienAlloc[vNo].push(String(om.no));
                 }
             }
         }
     }
  }
  return { visitors: visitors, pool: pool, facilAlloc: facilAlloc, orienAlloc: orienAlloc, roomAlloc: roomAlloc, connectReq: connectReq, hosts: getVisitorHosts(), mergedWith: mergedWith, priorities: getMemberPriorities() };
}

function saveAllocationSheet(meetingDateVal, displayVal, visitors, pool, facilAlloc, orienAlloc, roomAlloc, connectReq, mergedWith) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), dateObj = new Date(meetingDateVal), mmdd = Utilities.formatDate(dateObj, "Asia/Tokyo", "MMdd");
  var sheetName = mmdd + "割り振り表", sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();
  
  var visCount = 0, guestCount = 0, subCount = 0;
  visitors.forEach(function(v) {
    if (v.no.indexOf("V") === 0) visCount++;
    else if (v.no.indexOf("G") === 0) guestCount++;
    else subCount++; 
  });
  
  var activeRooms = new Set();
  visitors.forEach(function(v){
    if(!mergedWith[v.no]) {
      if(facilAlloc[v.no] || (roomAlloc[v.no] && roomAlloc[v.no].length > 0)) activeRooms.add(v.no);
    }
  });
  var roomCount = activeRooms.size;
  
  // 招待者名を正規化（全角スペース・半角スペース等を全て除去して比較用キーにする）
  function normalizeInvName(name) {
    return name ? String(name).replace(/[\s\u3000\u00A0]+/g, "").trim() : "";
  }

  // 招待者ごとの招待数を集計（同じ招待者が複数人招待→どれか1ルームに行くので 1/N で按分）
  var inviterCounts = {};
  visitors.forEach(function(v){
    var inv = normalizeInvName(v.inviter);
    if (inv !== "") inviterCounts[inv] = (inviterCounts[inv] || 0) + 1;
  });

  var totalPeopleInRooms = 0;
  activeRooms.forEach(function(vNo){
    var v = visitors.filter(function(x){ return x.no === vNo; })[0];
    // ファシリ＋ルームメンバーをカウント（オリエンは含めない。重複はSetで排除）
    var uniqueMembers = new Set();
    if (facilAlloc[vNo]) uniqueMembers.add(String(facilAlloc[vNo]));
    if (roomAlloc[vNo]) roomAlloc[vNo].forEach(function(id){ uniqueMembers.add(String(id)); });
    totalPeopleInRooms += uniqueMembers.size;
    totalPeopleInRooms += 1; // ビジター本人
    // 招待者は招待数で按分（複数招待→どれか1ルームへ行くため）
    var invName = normalizeInvName(v ? v.inviter : "");
    if (invName !== "") totalPeopleInRooms += 1 / (inviterCounts[invName] || 1);
    // 合同で入ってきた子ビジター分も加算
    visitors.forEach(function(cv){
      if (mergedWith[cv.no] === vNo) {
        totalPeopleInRooms += 1; // 子ビジター本人
        var childInv = normalizeInvName(cv.inviter);
        if (childInv !== "") totalPeopleInRooms += 1 / (inviterCounts[childInv] || 1);
      }
    });
  });
  
  var avgMembers = roomCount > 0 ? (totalPeopleInRooms / roomCount).toFixed(1) : "0.0";

  var outputData = [
    ["Activeチャプター " + displayVal + " 定例会 ビジター・見学者・代理様 割り振り表", "", "", "", "", "", "", ""],
    ["【ダッシュボード】 ビジター: " + visCount + "名 / ゲスト: " + guestCount + "名 / 代理: " + subCount + "名 / ルーム数: " + roomCount + " / 1ルーム平均総人数: " + avgMembers + "名", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["No.", "お名前", "カテゴリー", "招待者", "つなげたいメンバー", "ファシリテーター", "ルームメンバー", "オリエンテーション"]
  ];
  
  for (var i = 0; i < visitors.length; i++) {
    var v = visitors[i], fName = "", rNames = [], oNames = [];
    
    if (mergedWith[v.no]) {
       var targetV = visitors.filter(function(x){ return x.no === mergedWith[v.no]; })[0];
       var targetName = targetV ? targetV.name : mergedWith[v.no];
       fName = "【合同】" + targetName + " と同室";
       rNames = ["（同上）"];
    } else {
       if(facilAlloc[v.no]) { var fm = pool.filter(function(m){ return String(m.no) === String(facilAlloc[v.no]); })[0]; if(fm) fName = fm.name; }
       if(roomAlloc[v.no]) { for(var j=0; j<roomAlloc[v.no].length; j++) { var rm = pool.filter(function(m){ return String(m.no) === String(roomAlloc[v.no][j]); })[0]; if(rm) rNames.push(rm.name); } }
    }
    
    if(orienAlloc[v.no]) { for(var j=0; j<orienAlloc[v.no].length; j++) { var om = pool.filter(function(m){ return String(m.no) === String(orienAlloc[v.no][j]); })[0]; if(om) oNames.push(om.name); } }
    
    outputData.push([ v.no, v.name, v.cat, v.inviter, connectReq[v.no] || "", fName, rNames.join("\n"), oNames.join("\n") ]);
  }
  var lastDataRow = outputData.length;
  
  outputData.push(["", "", "", "", "", "", "", ""]);
  
  var noteText = getAllocationNote();
  var noteLines = noteText.split('\n');
  for (var n = 0; n < noteLines.length; n++) {
     outputData.push([noteLines[n], "", "", "", "", "", "", ""]);
  }

  outputData.push(["", "", "", "", "", "", "", ""]);
  
  var fullMembersList = getMembersList();
  var allMemberRoles = [];
  fullMembersList.forEach(function(m) {
    var roles = [];
    var mNoStr = String(m.no);
    visitors.forEach(function(v) {
      if (v.inviter === m.name) roles.push(v.name + " 様 (招待者)");
      
      var primaryNo = mergedWith[v.no] ? mergedWith[v.no] : v.no;
      if (String(facilAlloc[primaryNo]) === mNoStr) roles.push(v.name + " 様 (ファシリ)");
      if (roomAlloc[primaryNo] && roomAlloc[primaryNo].map(String).indexOf(mNoStr) !== -1) roles.push(v.name + " 様 (ルーム)");
      if (orienAlloc[v.no] && orienAlloc[v.no].map(String).indexOf(mNoStr) !== -1) roles.push(v.name + " 様 (オリエン)");
    });
    allMemberRoles.push([m.no, m.name, roles.length > 0 ? roles.join("\n") : "", "", "", "", "", ""]);
  });

  outputData.push(["【メンバー別 ルーム・オリエン担当表】", "", "", "", "", "", "", ""]);
  outputData.push(["No.", "メンバー名", "担当ビジター・役割", "", "", "", "", ""]);
  var reverseTableStart = outputData.length;
  allMemberRoles.forEach(function(row) { outputData.push(row); });

  sheet.getRange(1, 1, outputData.length, 8).setValues(outputData);
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A2:H2").mergeAcross().setFontSize(11).setFontColor("#333").setFontWeight("bold");
  
  sheet.getRange(4, 1, lastDataRow - 3, 8).setWrap(true).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
  sheet.getRange("A4:E4").setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("F4").setBackground("#fff2cc").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4").setBackground("#e6f2ff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("H4").setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("center");
  
  if(lastDataRow > 4) {
    sheet.getRange(5, 1, lastDataRow - 4, 1).setHorizontalAlignment("center"); 
    sheet.getRange(5, 2, lastDataRow - 4, 4).setHorizontalAlignment("left"); 
    sheet.getRange(5, 6, lastDataRow - 4, 3).setHorizontalAlignment("center"); 
  }
  
  var noteStartRow = lastDataRow + 2;
  if (noteLines.length > 0) {
    for (var n = 0; n < noteLines.length; n++) {
      var cell = sheet.getRange(noteStartRow + n, 1);
      if (noteLines[n].trim().indexOf("※") === 0) {
        cell.setFontWeight("bold").setFontColor("#d35400").setFontSize(11);
      } else {
        cell.setFontWeight("normal").setFontColor("#333").setFontSize(11);
      }
    }
  }
  
  if (allMemberRoles.length > 0) {
    sheet.getRange(reverseTableStart, 1, 1, 8).setBackground("#f3f3f3").setFontWeight("bold");
    sheet.getRange(reverseTableStart + 1, 1, allMemberRoles.length, 8).setBorder(true, true, true, true, true, true).setWrap(true).setVerticalAlignment("middle");
    for (var i = 0; i <= allMemberRoles.length; i++) {
       sheet.getRange(reverseTableStart + i, 3, 1, 6).mergeAcross();
    }
  }

  var widths = [40, 100, 120, 90, 120, 110, 120, 130];
  for(var w=0; w<widths.length; w++) sheet.setColumnWidth(w+1, widths[w]);
  SpreadsheetApp.flush();
  
  var pdfUrl = exportAllocationSheetToPdf(sheet, displayVal + " 割り振り表.pdf", meetingDateVal);
  PropertiesService.getScriptProperties().setProperty('LATEST_ALLOCATION_URL', pdfUrl); 
  return "<h3>作成完了しました🎉</h3><p>割り振り表を作成しました。</p><br><a href='" + pdfUrl + "' target='_blank' style='background:#0055ff; color:#fff; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold;'>📄 作成されたPDFを開く</a>";
}

function exportAllocationSheetToPdf(sheet, fileName, meetingDateVal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), spreadsheetId = ss.getId(), sheetId = sheet.getSheetId(), lastRow = sheet.getLastRow();
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=" + sheetId + "&r1=0&c1=0&r2=" + lastRow + "&c2=8";
  var token = ScriptApp.getOAuthToken(), response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
  var blob = response.getBlob().setName(fileName);
  var props = PropertiesService.getScriptProperties();
  var mmdd = meetingDateVal ? Utilities.formatDate(new Date(meetingDateVal), "Asia/Tokyo", "MMdd") : "";
  var propKey = mmdd ? 'ALLOCATION_PDF_ID_' + mmdd : '';
  var fileId = propKey ? props.getProperty(propKey) : null;
  if (fileId) {
    try { Drive.Files.update({}, fileId, blob); return DriveApp.getFileById(fileId).getUrl(); } catch(e) { fileId = null; }
  }
  var file = DriveApp.getFileById(spreadsheetId), folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  var pdfFile = folder.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (propKey) props.setProperty(propKey, pdfFile.getId());
  return pdfFile.getUrl();
}

function callGeminiAutoAllocation(currentState, maxRoomSize) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('GEMINI_API_KEY');
  var modelName = props.getProperty('GEMINI_MODEL_NAME') || "gemini-2.5-flash"; 
  
  if (!apiKey) throw new Error("Gemini APIキーが設定されていません。");

  var usedInRooms = new Set();
  for (var vNo in currentState.facilAlloc) { 
    if (currentState.facilAlloc[vNo]) usedInRooms.add(String(currentState.facilAlloc[vNo])); 
  }

  var memberPriorities = getMemberPriorities();
  var availableHosts = currentState.hosts.map(String).filter(function(h) { return !usedInRooms.has(h); });
  availableHosts.sort(function(a, b) {
    var pA = memberPriorities[a] !== undefined ? memberPriorities[a] : Infinity;
    var pB = memberPriorities[b] !== undefined ? memberPriorities[b] : Infinity;
    return pA - pB;
  });

  var sortedVisitors = currentState.visitors.filter(function(v) { return !currentState.mergedWith[v.no]; }).sort(function(a, b) {
    var kA = parseInt(String(a.details['メンバーになる確度'] || "0").replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); })) || 0;
    var kB = parseInt(String(b.details['メンバーになる確度'] || "0").replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); })) || 0;
    if (kA !== kB) return kB - kA; 
    
    var typePriority = { "Visitor": 3, "Guest": 2, "Substitute": 1 };
    var tA = typePriority[a.details['種別']] || 0;
    var tB = typePriority[b.details['種別']] || 0;
    return tB - tA; 
  });

  sortedVisitors.forEach(function(v) {
    if (!currentState.facilAlloc[v.no] && availableHosts.length > 0) {
      var h = availableHosts.shift();
      currentState.facilAlloc[v.no] = h;
      usedInRooms.add(h);
    }
    if (currentState.facilAlloc[v.no]) {
      currentState.orienAlloc[v.no] = [currentState.facilAlloc[v.no]];
    }
  });

  currentState.roomAlloc = {};

  var fileId = props.getProperty('MEMBER_BOOK_ID');
  var pdfPart = null;
  if(fileId) {
    try {
      var pdfBlob = DriveApp.getFileById(fileId).getBlob();
      var base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
      pdfPart = { "inlineData": { "mimeType": "application/pdf", "data": base64Pdf } };
    } catch(e) {}
  }

  var maskedVisitors = currentState.visitors.map(function(v) {
    return {
      no: String(v.no), name: v.name, category: v.cat, inviter: v.inviter,
      connectReq: currentState.connectReq[v.no] || "",
      memos: v.details['メモ（非表示）'] + " " + v.details['メモ（ビジターリストに表示）'],
      mergedWith: currentState.mergedWith[v.no] || ""
    };
  });

  var promptText = "あなたはプロのビジネス交流会コーディネーターです。\n" +
    "以下のビジター情報、利用可能な待機メンバー、および提供されたメンバーブックPDFの内容を参考に、最適な「ルームメンバー」を決定してください。\n\n" +
    "【制約条件（厳格に守ること）】\n" +
    "1. 「ファシリテーター」と「オリエンテーション」はシステムで決定済みです。出力には「room」の配列のみを含めてください。\n" +
    "2. 各ルームの総人数がなるべく同じになるよう、各ルームの人数を【厳格に平均化】して「ルームメンバー(room)」を割り振ってください。\n" +
    "3. 1つのルームの総人数は最大「" + maxRoomSize + "名」以下とします。\n" +
    "4. 1人のメンバーが複数のルームに重複して配置されることは【絶対に禁止】です。1人1枠のみです。また、「既存のファシリ配置」にいるメンバーも使えません。\n" +
    "5. 「つなげたいメンバー(connectReq)」に名前がある人を【最優先】で同室に配置してください。\n" +
    "6. 合同ルームになっているビジター（mergedWithに値がある）の room には何も割り当てないでください。\n" +
    "7. 出力する値（room）は、必ず待機メンバーリストにある【メンバーNo（例: 05, 12 など）】のみを出力してください。\n\n" +
    "【データ】\n" +
    "ビジター: " + JSON.stringify(maskedVisitors) + "\n" +
    "待機メンバー: " + JSON.stringify(currentState.pool.map(function(m){return {no:String(m.no), name:m.name}})) + "\n" +
    "既存のファシリ配置（この人達はroomに使えません）: " + JSON.stringify(currentState.facilAlloc) + "\n\n" +
    "出力は必ず以下のJSONフォーマットのみを返してください。\n" +
    "{\n" +
    "  \"allocations\": [\n" +
    "    { \"visitorNo\": \"V01\", \"room\": [\"メンバーNo\", \"メンバーNo\"] }\n" +
    "  ]\n" +
    "}";

  var parts = [{ "text": promptText }];
  if (pdfPart) parts.push(pdfPart); 

  var payload = { "contents": [{ "parts": parts }] };
  
  var response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey, {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  var resData = JSON.parse(response.getContentText());
  if(resData.error) throw new Error(resData.error.message);
  
  var text = resData.candidates[0].content.parts[0].text;
  var jsonStr = text.match(/\{[\s\S]*\}/)[0]; 
  var aiResult = JSON.parse(jsonStr);

  aiResult.allocations.forEach(function(alloc) {
    var vNo = String(alloc.visitorNo);
    if(!currentState.roomAlloc[vNo] && !currentState.mergedWith[vNo]) currentState.roomAlloc[vNo] = [];
    if(alloc.room && !currentState.mergedWith[vNo]) {
      alloc.room.forEach(function(r){
        var rStr = r ? String(r).trim() : "";
        if(rStr !== "" && !usedInRooms.has(rStr)) { 
          currentState.roomAlloc[vNo].push(rStr); 
          usedInRooms.add(rStr); 
        }
      });
    }
  });

  return currentState;
}