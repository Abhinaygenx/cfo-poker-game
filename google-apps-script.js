/**
 * ==============================================================================
 * CFO Leadership Pulse - Google Sheets & Email Automation Script
 * ==============================================================================
 * Automatically updates Google Spreadsheet with:
 * 1. "Players_Summary": Main dashboard with player profiles and overall scores.
 * 2. "Poker_Responses": Individual sheet logging all 30 Poker questions & user scores.
 * 3. "Rally_Responses": Individual sheet logging all 15 Rally questions & user scores.
 * 4. "Event_Log": Chronological audit trail of all activity.
 * 5. Sends automatic email scorecard notifications to: mmsbf26001@stu.xim.edu.in
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

    // 1. Update/Insert player record in Players_Summary
    const summarySheet = getOrCreateSheet(ss, "Players_Summary", getSummaryHeaders(), "#141829", "#f2c14e");
    updatePlayerRow(summarySheet, data);

    // 2. Record detailed question responses for Game 1 (Poker) if available
    if (data.poker && data.poker.responses && data.poker.responses.length > 0) {
      const pokerSheet = getOrCreateSheet(ss, "Poker_Responses", getPokerHeaders(), "#0f3b2a", "#6be3a4");
      recordPokerResponses(pokerSheet, data);
    }

    // 3. Record detailed question responses for Game 2 (Rally) if available
    if (data.rally && data.rally.responses && data.rally.responses.length > 0) {
      const rallySheet = getOrCreateSheet(ss, "Rally_Responses", getRallyHeaders(), "#4a1d12", "#ff9f4a");
      recordRallyResponses(rallySheet, data);
    }

    // 4. Append event to Event_Log
    const logSheet = getOrCreateSheet(ss, "Event_Log", getLogHeaders(), "#2a3050", "#eef0f7");
    appendLog(logSheet, data);

    // 5. Send notification email to the administrator
    sendEmailReport(data, ss.getUrl());

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Player summary and individual question responses saved successfully.",
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
 * Headers for Poker_Responses sheet
 */
function getPokerHeaders() {
  return [
    "Recorded At",
    "Participant ID",
    "Player Name",
    "Player Email",
    "Hand #",
    "Question Type",
    "Dimension",
    "Question / Statement",
    "User Answer Code",
    "User Answer Text",
    "Target / Key",
    "Distance / Diff",
    "Hand Result",
    "Stake %",
    "Stake Chips",
    "Chips Won / Lost",
    "Chips After Hand",
    "Time Taken (s)",
    "Streak After"
  ];
}

/**
 * Headers for Rally_Responses sheet
 */
function getRallyHeaders() {
  return [
    "Recorded At",
    "Participant ID",
    "Player Name",
    "Player Email",
    "Question #",
    "Section / Category",
    "Question / Statement",
    "User Choice (1-5)",
    "User Answer Label",
    "User Score Awarded (1-5)",
    "Reverse Scored?",
    "Reaction Time (s)"
  ];
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
 * Record individual Poker question responses
 */
function recordPokerResponses(sheet, data) {
  const p = data.participant || {};
  const pId = p.id || "";
  const responses = data.poker.responses || [];
  if (responses.length === 0) return;

  // Remove existing entries for this participant if already present to avoid duplicates
  deleteExistingParticipantRows(sheet, 2, pId);

  const RS = { w: "Win", p: "Push", l: "Loss" };
  const rows = [];
  const now = new Date().toLocaleString();

  responses.forEach(r => {
    rows.push([
      now,
      pId,
      p.name || "",
      p.email || "",
      r.hand || "",
      r.type || "",
      r.dimn || r.dim || "",
      r.q || "",
      r.ans || "",
      r.ansTxt || r.ans || "",
      r.target !== undefined ? r.target : (r.bestA || ""),
      r.diff !== undefined ? r.diff : "",
      RS[r.out] || r.out || "",
      r.stake !== undefined ? r.stake + "%" : "",
      r.st !== undefined ? r.st : "",
      r.dl !== undefined ? r.dl : "",
      r.c1 !== undefined ? r.c1 : "",
      r.tk !== undefined ? r.tk : "",
      r.sk !== undefined ? r.sk : ""
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Record individual Rally question responses
 */
function recordRallyResponses(sheet, data) {
  const p = data.participant || {};
  const pId = p.id || "";
  const responses = data.rally.responses || [];
  if (responses.length === 0) return;

  // Remove existing entries for this participant if already present to avoid duplicates
  deleteExistingParticipantRows(sheet, 2, pId);

  const rows = [];
  const now = new Date().toLocaleString();

  responses.forEach(a => {
    rows.push([
      now,
      pId,
      p.name || "",
      p.email || "",
      a.n || "",
      a.cat || "",
      a.tx || "",
      a.ch !== undefined ? a.ch : (a.l !== undefined ? a.l + 1 : ""),
      a.lab || "",
      a.s !== undefined ? a.s : "",
      a.rev ? "Yes" : "No",
      a.rt !== undefined ? a.rt : ""
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * Helper to delete rows for a participant to keep questions cleanly updated
 */
function deleteExistingParticipantRows(sheet, idCol, targetId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const values = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && values[i][0].toString() === targetId.toString()) {
      sheet.deleteRow(i + 2);
    }
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
    `Poker: ${data.poker && data.poker.completed ? "Done (" + (data.poker.responses ? data.poker.responses.length : 0) + " hands)" : "Pending"} | Rally: ${data.rally && data.rally.completed ? "Done (" + (data.rally.responses ? data.rally.responses.length : 0) + " Qs)" : "Pending"}`,
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
              <td style="padding: 4px 0; color: #9aa1b8;">Questions Logged:</td>
              <td style="padding: 4px 0; color: #ffd66b;">${pokerHandsCount} hands recorded in sheet tab 'Poker_Responses'</td>
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
              <td style="padding: 4px 0; color: #9aa1b8;">Questions Logged:</td>
              <td style="padding: 4px 0; color: #ffd66b;">${rallyQsCount} statements recorded in sheet tab 'Rally_Responses'</td>
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
          Tabs available: <b>Players_Summary</b>, <b>Poker_Responses</b>, <b>Rally_Responses</b>, <b>Event_Log</b><br>
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
