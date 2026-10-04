/**
 * ==============================================================================
 * CFO Leadership Pulse - Google Sheets & Email Automation Script
 * ==============================================================================
 * Automatically organizes and updates Google Spreadsheet with:
 * 1. "Players_Summary": Overview dashboard (1 row per player with total scores).
 * 2. "Poker_Responses": Organised wide-table (1 row per player with Hand 1 to 30 columns).
 * 3. "Rally_Responses": Organised wide-table (1 row per player with Q1 to Q15 columns).
 * 4. "Event_Log": Chronological audit trail.
 * 5. Sends automatic scorecard email notifications to: mmsbf26001@stu.xim.edu.in
 * ==============================================================================
 */

// Target email where updates will be sent
const NOTIFICATION_EMAIL = "mmsbf26001@stu.xim.edu.in";

/**
 * Handle GET request (used to verify endpoint is live)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "CFO Leadership Pulse Webhook is online and active!",
    recipient: NOTIFICATION_EMAIL,
    sheets: ["Players_Summary", "Poker_Responses", "Rally_Responses", "Event_Log"],
    format: "One row per player with individual question columns",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST request from the game
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "No POST body received"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Update/Insert in "Players_Summary"
    const summarySheet = getOrCreateSheet(ss, "Players_Summary", getSummaryHeaders(), "#141829", "#f2c14e");
    updatePlayerRow(summarySheet, data);

    // 2. Update/Insert in "Poker_Responses" (One row per player, Hand 1 to 30 as columns)
    if (data.poker && data.poker.responses && data.poker.responses.length > 0) {
      const pokerSheet = getOrCreateSheet(ss, "Poker_Responses", getPokerWideHeaders(), "#0f3b2a", "#6be3a4");
      updatePokerWideRow(pokerSheet, data);
    }

    // 3. Update/Insert in "Rally_Responses" (One row per player, Q1 to Q15 as columns)
    if (data.rally && data.rally.responses && data.rally.responses.length > 0) {
      const rallySheet = getOrCreateSheet(ss, "Rally_Responses", getRallyWideHeaders(), "#4a1d12", "#ff9f4a");
      updateRallyWideRow(rallySheet, data);
    }

    // 4. Append to "Event_Log"
    const logSheet = getOrCreateSheet(ss, "Event_Log", getLogHeaders(), "#2a3050", "#eef0f7");
    appendLog(logSheet, data);

    // 5. Send notification email to the administrator
    sendEmailReport(data, ss.getUrl());

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Player data successfully organized into columns and updated.",
      recipient: NOTIFICATION_EMAIL
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error processing webhook: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Headers for Players_Summary sheet
 */
function getSummaryHeaders() {
  return [
    "Last Updated",
    "Participant ID",
    "Full Name",
    "Email",
    "Phone",
    "Age",
    "Gender",
    "Status",
    "Nickname",
    "Progress State",
    "Poker Completed",
    "Poker Chips",
    "Poker Level",
    "Poker Overall %",
    "Poker Best Streak",
    "Rally Completed",
    "Rally Title",
    "Rally Score",
    "Rally Max",
    "Rally %",
    "Rally CFO Score",
    "Rally EQ Score",
    "Crashes",
    "Coins"
  ];
}

/**
 * Headers for Poker_Responses sheet (Wide table: 1 row per player)
 */
function getPokerWideHeaders() {
  const headers = [
    "Last Updated",
    "Participant ID",
    "Full Name",
    "Email",
    "Nickname",
    "Final Chips",
    "CFO Level",
    "Overall %",
    "Best Streak"
  ];

  for (let h = 1; h <= 30; h++) {
    const num = h < 10 ? "0" + h : "" + h;
    headers.push(`H${num}_Answer`);
    headers.push(`H${num}_Result`);
    headers.push(`H${num}_Stake`);
  }
  return headers;
}

/**
 * Headers for Rally_Responses sheet (Wide table: 1 row per player)
 */
function getRallyWideHeaders() {
  const headers = [
    "Last Updated",
    "Participant ID",
    "Full Name",
    "Email",
    "Nickname",
    "Total Score",
    "Max Score",
    "Percent (%)",
    "Rally Title",
    "CFO Score",
    "EQ Score"
  ];

  for (let q = 1; q <= 15; q++) {
    const num = q < 10 ? "0" + q : "" + q;
    headers.push(`Q${num}_Choice`);
    headers.push(`Q${num}_Score`);
  }
  return headers;
}

/**
 * Headers for Event_Log sheet
 */
function getLogHeaders() {
  return [
    "Timestamp",
    "Event Type",
    "Participant ID",
    "Full Name",
    "Email",
    "Status Details",
    "Raw Payload"
  ];
}

/**
 * Helper to get or create sheet with headers and formatting
 */
function getOrCreateSheet(ss, name, headers, bgColor, fontColor) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    
    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(bgColor || "#141829");
    headerRange.setFontColor(fontColor || "#f2c14e");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(10.5);
    sheet.setFrozenRows(1);
    
    // Auto-fit column widths
    for (let c = 1; c <= headers.length; c++) {
      sheet.setColumnWidth(c, 130);
    }
  }
  return sheet;
}

/**
 * Update or Insert player record in Players_Summary
 */
function updatePlayerRow(sheet, data) {
  const p = data.participant || {};
  const poker = data.poker || {};
  const rally = data.rally || {};
  const pId = p.id || "";

  const rowValues = [
    new Date().toLocaleString(),
    pId,
    p.name || "",
    p.email || "",
    p.phone || "",
    p.age || "",
    p.gender || "",
    p.status || "",
    p.nick || "",
    data.eventType || "UPDATE",
    poker.completed ? "Yes" : "No",
    poker.chips !== undefined ? poker.chips : "",
    poker.level || "",
    poker.overallPct !== undefined ? poker.overallPct + "%" : "",
    poker.bestStreak !== undefined ? poker.bestStreak : "",
    rally.completed ? "Yes" : "No",
    rally.title || "",
    rally.score !== undefined ? rally.score : "",
    rally.maxScore !== undefined ? rally.maxScore : "",
    rally.percent !== undefined ? rally.percent + "%" : "",
    rally.cfoScore !== undefined ? rally.cfoScore : "",
    rally.eqScore !== undefined ? rally.eqScore : "",
    rally.crashes !== undefined ? rally.crashes : "",
    rally.coins !== undefined ? rally.coins : ""
  ];

  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const idRange = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (idRange[i][0] && idRange[i][0].toString() === pId.toString()) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(foundRow, 1, 1, 1).setBackground("#e8f5e9");
  } else {
    sheet.appendRow(rowValues);
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 1).setBackground("#fff8e1");
  }
}

/**
 * Update or Insert player record in Poker_Responses (1 row per player, Hand 1-30 in columns)
 */
function updatePokerWideRow(sheet, data) {
  const p = data.participant || {};
  const poker = data.poker || {};
  const pId = p.id || "";
  const responses = poker.responses || [];

  const rowValues = [
    new Date().toLocaleString(),
    pId,
    p.name || "",
    p.email || "",
    p.nick || "",
    poker.chips !== undefined ? poker.chips : "",
    poker.level || "",
    poker.overallPct !== undefined ? poker.overallPct + "%" : "",
    poker.bestStreak !== undefined ? poker.bestStreak : ""
  ];

  const RS = { w: "Win", p: "Push", l: "Loss" };
  for (let h = 1; h <= 30; h++) {
    const r = responses.find(x => x.hand == h);
    if (r) {
      rowValues.push(r.ansTxt || r.ans || "");
      rowValues.push(RS[r.out] || r.out || "");
      rowValues.push(r.stake !== undefined ? r.stake + "%" : "");
    } else {
      rowValues.push("");
      rowValues.push("");
      rowValues.push("");
    }
  }

  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const idRange = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (idRange[i][0] && idRange[i][0].toString() === pId.toString()) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(foundRow, 1, 1, 1).setBackground("#e8f5e9");
  } else {
    sheet.appendRow(rowValues);
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 1).setBackground("#fff8e1");
  }
}

/**
 * Update or Insert player record in Rally_Responses (1 row per player, Q1-15 in columns)
 */
function updateRallyWideRow(sheet, data) {
  const p = data.participant || {};
  const rally = data.rally || {};
  const pId = p.id || "";
  const responses = rally.responses || [];

  const rowValues = [
    new Date().toLocaleString(),
    pId,
    p.name || "",
    p.email || "",
    p.nick || "",
    rally.score !== undefined ? rally.score : "",
    rally.maxScore !== undefined ? rally.maxScore : 75,
    rally.percent !== undefined ? rally.percent + "%" : "",
    rally.title || "",
    rally.cfoScore !== undefined ? rally.cfoScore : "",
    rally.eqScore !== undefined ? rally.eqScore : ""
  ];

  for (let q = 1; q <= 15; q++) {
    const a = responses.find(x => x.n == q);
    if (a) {
      const choiceStr = a.lab ? `${a.ch || a.l + 1} - ${a.lab}` : (a.ch || a.l + 1 || "");
      rowValues.push(choiceStr);
      rowValues.push(a.s !== undefined ? a.s : "");
    } else {
      rowValues.push("");
      rowValues.push("");
    }
  }

  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const idRange = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (idRange[i][0] && idRange[i][0].toString() === pId.toString()) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(foundRow, 1, 1, 1).setBackground("#e8f5e9");
  } else {
    sheet.appendRow(rowValues);
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 1).setBackground("#fff8e1");
  }
}

/**
 * Append to Event_Log sheet
 */
function appendLog(sheet, data) {
  const p = data.participant || {};
  sheet.appendRow([
    new Date().toLocaleString(),
    data.eventType || "EVENT",
    p.id || "",
    p.name || "",
    p.email || "",
    `Poker: ${data.poker && data.poker.completed ? "Done" : "Pending"} | Rally: ${data.rally && data.rally.completed ? "Done" : "Pending"}`,
    JSON.stringify(data)
  ]);
}

/**
 * Send an attractive HTML summary email
 */
function sendEmailReport(data, sheetUrl) {
  const p = data.participant || {};
  const poker = data.poker || {};
  const rally = data.rally || {};
  const event = data.eventType || "UPDATE";

  let eventTitle = "Player Update";
  let statusBadgeColor = "#2196f3";

  if (event === "REGISTER") {
    eventTitle = "New Player Registered";
    statusBadgeColor = "#ff9800";
  } else if (event === "POKER_COMPLETED") {
    eventTitle = "Hot Seat (Poker) Completed";
    statusBadgeColor = "#4caf50";
  } else if (event === "RALLY_COMPLETED") {
    eventTitle = "Desert Rally Completed";
    statusBadgeColor = "#e91e63";
  } else if (event === "FINISHED") {
    eventTitle = "Both Games Finished & Full Results Recorded";
    statusBadgeColor = "#9c27b0";
  }

  const subject = `[CFO Leadership Pulse] ${eventTitle}: ${p.name || "Player"} (${p.nick || "No Nickname"})`;

  const pokerHandsCount = poker.responses ? poker.responses.length : (poker.handsPlayed || 0);
  const rallyQsCount = rally.responses ? rally.responses.length : 0;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; background-color: #0b0d16; color: #eef0f7; padding: 24px; border-radius: 12px; max-width: 650px; margin: auto; border: 1px solid #2a3050;">
      <div style="text-align: center; border-bottom: 1px solid #2a3050; padding-bottom: 16px; margin-bottom: 20px;">
        <h1 style="color: #f2c14e; margin: 0; font-size: 24px; font-family: Georgia, serif;">CFO Leadership Pulse</h1>
        <p style="color: #9aa1b8; font-size: 13px; margin: 6px 0 0;">Real-time Player Assessment Notification</p>
        <span style="display: inline-block; background-color: ${statusBadgeColor}; color: #ffffff; padding: 4px 12px; border-radius: 20px; font-weight: bold; font-size: 12px; margin-top: 10px;">${eventTitle}</span>
      </div>

      <!-- Player Profile Card -->
      <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 16px; margin-bottom: 18px;">
        <h3 style="color: #f2c14e; margin-top: 0; margin-bottom: 10px; font-size: 16px;">👤 Player Profile</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #eef0f7;">
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8; width: 40%;">Full Name:</td>
            <td style="padding: 5px 0; font-weight: bold;">${p.name || "N/A"}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8;">Nickname:</td>
            <td style="padding: 5px 0; color: #ffd66b; font-weight: bold;">${p.nick || "N/A"}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8;">Email:</td>
            <td style="padding: 5px 0;">${p.email || "N/A"}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8;">Phone:</td>
            <td style="padding: 5px 0;">${p.phone || "N/A"}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8;">Status / Background:</td>
            <td style="padding: 5px 0;">${p.status || "N/A"} (Age: ${p.age || "N/A"}, Gender: ${p.gender || "N/A"})</td>
          </tr>
          <tr>
            <td style="padding: 5px 0; color: #9aa1b8;">Participant ID:</td>
            <td style="padding: 5px 0; font-family: monospace; color: #9aa1b8;">${p.id || "N/A"}</td>
          </tr>
        </table>
      </div>

      <!-- Game 1: Poker Table -->
      <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 16px; margin-bottom: 18px;">
        <h3 style="color: #6be3a4; margin-top: 0; margin-bottom: 10px; font-size: 16px;">♠️ Game 1: The Hot Seat (Poker Table)</h3>
        ${poker.completed ? `
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #eef0f7;">
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8; width: 40%;">Status:</td>
              <td style="padding: 4px 0; color: #6be3a4; font-weight: bold;">Completed ✓</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">Final Chips:</td>
              <td style="padding: 4px 0; font-size: 16px; font-weight: bold; color: #f2c14e;">$${poker.chips || 0}</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">CFO Fit Level:</td>
              <td style="padding: 4px 0; font-weight: bold;">${poker.level || "N/A"}</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">Overall Score %:</td>
              <td style="padding: 4px 0;">${poker.overallPct || "N/A"}%</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">Best Streak:</td>
              <td style="padding: 4px 0;">${poker.bestStreak || 0} wins in a row</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">All 30 Hands Logged:</td>
              <td style="padding: 4px 0; color: #ffd66b;">${pokerHandsCount} columns updated in sheet 'Poker_Responses'</td>
            </tr>
          </table>
        ` : `
          <p style="color: #9aa1b8; font-size: 13px; margin: 0;">Game 1 not yet completed.</p>
        `}
      </div>

      <!-- Game 2: Desert Rally -->
      <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 16px; margin-bottom: 18px;">
        <h3 style="color: #ff9f4a; margin-top: 0; margin-bottom: 10px; font-size: 16px;">🏎️ Game 2: CFO Desert Rally</h3>
        ${rally.completed ? `
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #eef0f7;">
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8; width: 40%;">Status:</td>
              <td style="padding: 4px 0; color: #6be3a4; font-weight: bold;">Completed ✓</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">Rally Rank / Title:</td>
              <td style="padding: 4px 0; font-weight: bold; color: #ffd66b;">${rally.title || "N/A"}</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">Overall Score:</td>
              <td style="padding: 4px 0; font-weight: bold;">${rally.score || 0} / ${rally.maxScore || 75} (${rally.percent || 0}%)</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">CFO Skills vs EQ:</td>
              <td style="padding: 4px 0;">CFO: ${rally.cfoScore || 0} pts | EQ: ${rally.eqScore || 0} pts</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #9aa1b8;">All 15 Qs Logged:</td>
              <td style="padding: 4px 0; color: #ffd66b;">${rallyQsCount} columns updated in sheet 'Rally_Responses'</td>
            </tr>
          </table>
        ` : `
          <p style="color: #9aa1b8; font-size: 13px; margin: 0;">Game 2 not yet completed.</p>
        `}
      </div>

      <!-- Action Button / Link to Google Sheet -->
      <div style="text-align: center; margin-top: 24px; padding-top: 16px; border-top: 1px solid #2a3050;">
        <a href="${sheetUrl}" target="_blank" style="background: linear-gradient(#ffd66b, #f2a93e); color: #2a1a10; font-weight: bold; text-decoration: none; padding: 12px 24px; border-radius: 8px; display: inline-block; font-size: 14px;">
          📊 Open Google Spreadsheet
        </a>
        <p style="color: #9aa1b8; font-size: 12px; margin-top: 14px;">
          Clean 1-row-per-player format with question columns.<br>
          Target Recipient: ${NOTIFICATION_EMAIL}
        </p>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}
