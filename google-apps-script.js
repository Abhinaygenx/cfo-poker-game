/**
 * ==============================================================================
 * CFO Leadership Pulse - Google Sheets & Email Automation Script
 * ==============================================================================
 * Automatically organizes Google Spreadsheet into clean 1-ROW-PER-PLAYER tables:
 * 1. "Players_Summary": Dashboard with player profiles and overall scores.
 * 2. "Poker_Responses": Exactly 1 row per player with questions as individual columns.
 * 3. "Rally_Responses": Exactly 1 row per player with all 15 questions as individual columns.
 * 4. "Event_Log": Chronological audit trail.
 * 5. Sends automatic scorecard email notifications to: mmsbf26001@stu.xim.edu.in
 * ==============================================================================
 */

// Target email where updates will be sent
const NOTIFICATION_EMAIL = "mmsbf26001@stu.xim.edu.in";

// 15 Exact Rally Questions for Column Headers
const RALLY_QUESTIONS = [
  "Q1: I am aware of my emotions as I experience them",
  "Q2: If a profitable shortcut bends the rules slightly, I'd still take it.",
  "Q3: I help other people feel better when they are down",
  "Q4: I double-check figures even when they come from trusted sources.",
  "Q5: I know why my emotions change",
  "Q6: I'd rather miss a big opportunity than expose the company to major uncertainty.",
  "Q7: I use good moods to help myself keep trying in the face of obstacles",
  "Q8: I have control over my emotions",
  "Q9: I am aware of the non-verbal messages other people send",
  "Q10: I stay calm and decisive when the team is under financial pressure.",
  "Q11: I motivate myself by imagining a good outcome to tasks I take on",
  "Q12: By looking at their facial expressions, I recognize the emotions people are experiencing",
  "Q13: I regularly think about where capital would earn the best return.",
  "Q14: When I experience a positive emotion, I know how to make it last",
  "Q15: When I am in a positive mood, solving problems is easy for me"
];

// 10 Hot Seat Poker Questions (5 Ethical Decision Making + 5 Quantitative & Reasoning Puzzles)
const POKER_QUESTIONS = [
  "H01 [Ethical]: $0.2M short of consensus, $0.2M warranty reserve slightly high",
  "H02 [Ethical]: Manager repaid $2,000 padded expenses, asks not to report",
  "H03 [Ethical]: Brother-in-law's firm bidding on funded project",
  "H04 [Ethical]: Unusual contracts with side letter allowing returns",
  "H05 [Ethical]: 30% chance of covenant breach, CEO wants to hold back at bank meeting",
  "H06 [IQ]: Invest $8,000, falls 50%, then rises 75%. Worth now?",
  "H07 [IQ]: 1000 invoices (400 risky, 90 flagged fraud, 60 unflagged fraud). P(flagged|fraud)?",
  "H08 [IQ]: Bat and ball cost $1.10 total. Bat costs $1.00 more than ball. Ball cost?",
  "H09 [IQ]: 5 machines 5 min 5 widgets. Time for 100 machines to make 100 widgets?",
  "H10 [IQ]: Lily pad doubles daily, takes 48 days for entire lake. Days for half lake?"
];

/**
 * Handle GET request (used to verify endpoint is live)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "CFO Leadership Pulse Webhook is online and active!",
    recipient: NOTIFICATION_EMAIL,
    sheets: ["Players_Summary", "Poker_Responses", "Rally_Responses", "Event_Log"],
    format: "1-Row-Per-Player with individual question columns",
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

    // 1. Update/Insert in "Players_Summary" (1 row per player)
    const summarySheet = getOrCreateSheet(ss, "Players_Summary", getSummaryHeaders(), "#141829", "#f2c14e");
    updatePlayerRow(summarySheet, data);

    // 2. Update/Insert in "Poker_Responses" (1 row per player with 30 question columns)
    if (data.poker && data.poker.responses && data.poker.responses.length > 0) {
      const pokerSheet = getOrCreateSheet(ss, "Poker_Responses", getPokerHeaders(), "#0f3b2a", "#6be3a4");
      updatePokerRow(pokerSheet, data);
    }

    // 3. Update/Insert in "Rally_Responses" (1 row per player with 15 question columns)
    if (data.rally && data.rally.responses && data.rally.responses.length > 0) {
      const rallySheet = getOrCreateSheet(ss, "Rally_Responses", getRallyHeaders(), "#4a1d12", "#ff9f4a");
      updateRallyRow(rallySheet, data);
    }

    // 4. Append to "Event_Log"
    const logSheet = getOrCreateSheet(ss, "Event_Log", getLogHeaders(), "#2a3050", "#eef0f7");
    appendLog(logSheet, data);

    // NOTE: Per-signup/per-event emails are intentionally disabled to avoid inbox spam.
    // All participant data is recorded in the sheets above.
    // A single consolidated digest email is sent at the end of the day via sendDailySummaryEmail().

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Player data successfully organized into single-row question columns and updated in Google Sheet.",
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
 * Headers for Poker_Responses sheet (1 row per player, 10 question columns)
 */
function getPokerHeaders() {
  const headers = [
    "Last Updated",
    "Participant ID",
    "Full Name",
    "Email",
    "Nickname",
    "Final Chips ($)",
    "CFO Fit Level",
    "Overall Score %",
    "Best Streak"
  ];
  return headers.concat(POKER_QUESTIONS);
}

/**
 * Headers for Rally_Responses sheet (1 row per player, 15 question columns)
 */
function getRallyHeaders() {
  const headers = [
    "Last Updated",
    "Participant ID",
    "Full Name",
    "Email",
    "Nickname",
    "Total Score",
    "Max Score",
    "Percent (%)",
    "CFO Score",
    "EQ Score",
    "Rally Title"
  ];
  return headers.concat(RALLY_QUESTIONS);
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
 * Helper to get or create sheet with headers and formatting.
 * If the sheet exists but has outdated column headers (e.g. old 30-hand columns),
 * it wipes the old structure and responses, applying the new clean 10-hand structure.
 */
function getOrCreateSheet(ss, name, headers, bgColor, fontColor) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    formatHeaderRow(sheet, headers.length, bgColor, fontColor);
  } else {
    const curCols = sheet.getLastColumn();
    if (curCols !== headers.length) {
      // Column structure changed (e.g. from 30 questions to 10 questions)
      // Wipe old outdated responses to prevent mismatched column alignments
      sheet.clear();
      sheet.appendRow(headers);
      formatHeaderRow(sheet, headers.length, bgColor, fontColor);
    }
  }
  return sheet;
}

function formatHeaderRow(sheet, colCount, bgColor, fontColor) {
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground(bgColor || "#141829");
  headerRange.setFontColor(fontColor || "#f2c14e");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  headerRange.setWrap(true);
  sheet.setFrozenRows(1);
  for (let c = 1; c <= colCount; c++) {
    sheet.setColumnWidth(c, c <= 5 ? 140 : 220);
  }
}

/**
 * Utility to immediately reset and remove old responses in Poker_Responses sheet.
 * Can be run manually from Google Apps Script editor.
 */
function resetPokerResponsesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Poker_Responses");
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet("Poker_Responses");
  }
  const headers = getPokerHeaders();
  sheet.appendRow(headers);
  formatHeaderRow(sheet, headers.length, "#0f3b2a", "#6be3a4");
  Logger.log("✅ Poker_Responses sheet wiped and reset with 10 question columns.");
}

/**
 * Update or Insert player record in Players_Summary (1 row per player)
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

  upsertRow(sheet, 2, pId, rowValues);
}

/**
 * Update or Insert player record in Poker_Responses (1 row per player, 10 question columns)
 */
function updatePokerRow(sheet, data) {
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
  for (let h = 1; h <= POKER_QUESTIONS.length; h++) {
    const r = responses.find(x => x.hand == h);
    if (r) {
      const outcome = RS[r.out] || r.out || "";
      const delta = r.dl !== undefined ? (r.dl > 0 ? "+" + r.dl : "" + r.dl) : "";
      const ansText = r.ansTxt || (r.ans !== undefined ? "Option " + r.ans : "");
      rowValues.push(`${ansText} [${outcome} ${delta}]`);
    } else {
      rowValues.push("");
    }
  }

  upsertRow(sheet, 2, pId, rowValues);
}

/**
 * Update or Insert player record in Rally_Responses (1 row per player, 15 question columns)
 */
function updateRallyRow(sheet, data) {
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
    rally.cfoScore !== undefined ? rally.cfoScore : "",
    rally.eqScore !== undefined ? rally.eqScore : "",
    rally.title || ""
  ];

  for (let q = 1; q <= 15; q++) {
    const a = responses.find(x => x.n == q);
    if (a) {
      const label = a.lab || (a.ch !== undefined ? "Choice " + a.ch : "");
      const score = a.s !== undefined ? a.s + " pts" : "";
      rowValues.push(`${label} (${score})`);
    } else {
      rowValues.push("");
    }
  }

  upsertRow(sheet, 2, pId, rowValues);
}

/**
 * Helper to update row if participant ID exists, or append if new (ensuring strictly 1 row per player)
 */
function upsertRow(sheet, idColIndex, targetId, rowValues) {
  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const idRange = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (idRange[i][0] && idRange[i][0].toString() === targetId.toString()) {
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
              <td style="padding: 4px 0; color: #9aa1b8;">Responses Layout:</td>
              <td style="padding: 4px 0; color: #ffd66b;">Saved as 10 question columns in 1 single row</td>
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
              <td style="padding: 4px 0; color: #9aa1b8;">Responses Layout:</td>
              <td style="padding: 4px 0; color: #ffd66b;">Saved as 15 question columns in 1 single row</td>
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
          Clean 1-Row-Per-User layout with questions as columns.<br>
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

/**
 * ==============================================================================
 * DAILY SUMMARY / DIGEST EMAIL (ONE EMAIL PER DAY AT END OF DAY)
 * ==============================================================================
 * Gathers all responses & player signups recorded during the day and sends a
 * single consolidated executive summary email to NOTIFICATION_EMAIL.
 */
function sendDailySummaryEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUrl = ss.getUrl();
  const summarySheet = ss.getSheetByName("Players_Summary");
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayFormatted = Utilities.formatDate(now, tz, "EEE, MMM d, yyyy");
  const todayShort = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  if (!summarySheet || summarySheet.getLastRow() <= 1) {
    Logger.log("No player data available to send daily summary.");
    return;
  }

  const values = summarySheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);

  // Filter rows updated or active today
  const todayRows = [];
  const allRows = [];

  rows.forEach(r => {
    const rawDate = r[0]; // Last Updated
    let isToday = false;
    if (rawDate) {
      try {
        const d = new Date(rawDate);
        if (!isNaN(d.getTime())) {
          const rowDateStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
          if (rowDateStr === todayShort) isToday = true;
        } else {
          // Fallback string matching
          const str = String(rawDate);
          const tDay = Utilities.formatDate(now, tz, "M/d/yyyy");
          const tDay2 = Utilities.formatDate(now, tz, "d/M/yyyy");
          if (str.indexOf(tDay) !== -1 || str.indexOf(tDay2) !== -1) isToday = true;
        }
      } catch (e) {}
    }
    allRows.push(r);
    if (isToday) todayRows.push(r);
  });

  // If no rows were strictly timestamped today, include the most recent 10 rows for context
  const displayRows = todayRows.length > 0 ? todayRows : allRows.slice(-10);

  // Compute stats
  let totalToday = todayRows.length;
  let pokerDoneToday = 0;
  let rallyDoneToday = 0;
  let bothDoneToday = 0;

  displayRows.forEach(r => {
    const pokerDone = String(r[10]).toLowerCase() === "yes";
    const rallyDone = String(r[15]).toLowerCase() === "yes";
    const bothDone = String(r[20]).toLowerCase() === "yes";
    if (pokerDone) pokerDoneToday++;
    if (rallyDone) rallyDoneToday++;
    if (bothDone) bothDoneToday++;
  });

  const subject = `📊 Daily CFO Assessment Digest: ${todayFormatted} (${displayRows.length} Participants)`;

  // Generate table rows HTML
  const rowsHtml = displayRows.map((r, idx) => {
    const pokerDone = String(r[10]).toLowerCase() === "yes";
    const rallyDone = String(r[15]).toLowerCase() === "yes";
    const bothDone = String(r[20]).toLowerCase() === "yes";
    const rowBg = idx % 2 === 0 ? "#141829" : "#0e1120";

    return `
      <tr style="background-color: ${rowBg}; border-bottom: 1px solid #2a3050; font-size: 13px;">
        <td style="padding: 10px; color: #eef0f7; font-weight: bold;">
          ${r[2] || "Unknown"}<br>
          <span style="color: #ffd66b; font-weight: normal; font-size: 11px;">${r[8] || ""} · ${r[7] || ""}</span>
        </td>
        <td style="padding: 10px; color: #9aa1b8; font-size: 12px;">
          ${r[3] || "N/A"}<br>
          <span style="font-size: 11px; color: #626987;">${r[4] || ""} (Age ${r[5] || ""})</span>
        </td>
        <td style="padding: 10px; text-align: center;">
          ${pokerDone ? `<span style="color: #6be3a4; font-weight: bold;">✓ $${r[11]}</span><br><span style="font-size: 11px; color: #9aa1b8;">${r[12] || ""} (${r[13]}%)</span>` : `<span style="color: #626987;">—</span>`}
        </td>
        <td style="padding: 10px; text-align: center;">
          ${rallyDone ? `<span style="color: #ff9f4a; font-weight: bold;">✓ ${r[17]}/${r[18]}</span><br><span style="font-size: 11px; color: #9aa1b8;">${r[16] || ""} (${r[19]}%)</span>` : `<span style="color: #626987;">—</span>`}
        </td>
        <td style="padding: 10px; text-align: center;">
          ${bothDone ? `<span style="background: #1f8a5b; color: #fff; padding: 3px 8px; border-radius: 99px; font-size: 11px; font-weight: bold;">COMPLETED 🏆</span>` : `<span style="background: #2a3050; color: #9aa1b8; padding: 3px 8px; border-radius: 99px; font-size: 11px;">In Progress</span>`}
        </td>
      </tr>
    `;
  }).join("");

  const htmlBody = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 680px; margin: 0 auto; background-color: #0b0d16; color: #eef0f7; padding: 24px; border-radius: 12px; border: 1px solid #2a3050;">
      
      <!-- Header -->
      <div style="text-align: center; border-bottom: 1px solid #2a3050; padding-bottom: 18px; margin-bottom: 20px;">
        <span style="letter-spacing: 0.18em; font-size: 11px; color: #f2c14e; text-transform: uppercase; font-weight: bold;">Daily Consolidated Digest · 1-Row Per Player</span>
        <h1 style="color: #ffffff; margin: 6px 0 4px; font-size: 22px;">CFO Leadership Pulse Summary</h1>
        <p style="color: #9aa1b8; margin: 0; font-size: 13px;">${todayFormatted} · End-of-Day Executive Report</p>
      </div>

      <!-- KPI Summary Cards -->
      <div style="display: table; width: 100%; margin-bottom: 20px;">
        <div style="display: table-cell; width: 25%; padding: 4px;">
          <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 12px; text-align: center;">
            <div style="font-size: 11px; color: #9aa1b8; text-transform: uppercase;">Participants</div>
            <div style="font-size: 22px; font-weight: bold; color: #ffd66b; margin-top: 4px;">${displayRows.length}</div>
          </div>
        </div>
        <div style="display: table-cell; width: 25%; padding: 4px;">
          <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 12px; text-align: center;">
            <div style="font-size: 11px; color: #9aa1b8; text-transform: uppercase;">Hot Seat</div>
            <div style="font-size: 22px; font-weight: bold; color: #6be3a4; margin-top: 4px;">${pokerDoneToday}</div>
          </div>
        </div>
        <div style="display: table-cell; width: 25%; padding: 4px;">
          <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 12px; text-align: center;">
            <div style="font-size: 11px; color: #9aa1b8; text-transform: uppercase;">Desert Rally</div>
            <div style="font-size: 22px; font-weight: bold; color: #ff9f4a; margin-top: 4px;">${rallyDoneToday}</div>
          </div>
        </div>
        <div style="display: table-cell; width: 25%; padding: 4px;">
          <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; padding: 12px; text-align: center;">
            <div style="font-size: 11px; color: #9aa1b8; text-transform: uppercase;">Both Done 🏆</div>
            <div style="font-size: 22px; font-weight: bold; color: #5fe39a; margin-top: 4px;">${bothDoneToday}</div>
          </div>
        </div>
      </div>

      <!-- Participant Table -->
      <div style="background-color: #141829; border: 1px solid #2a3050; border-radius: 8px; overflow: hidden; margin-bottom: 22px;">
        <table style="width: 100%; border-collapse: collapse; text-align: left;">
          <thead>
            <tr style="background-color: #1a2035; color: #ffd66b; font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; border-bottom: 1px solid #2a3050;">
              <th style="padding: 10px;">Participant</th>
              <th style="padding: 10px;">Email / Details</th>
              <th style="padding: 10px; text-align: center;">Hot Seat</th>
              <th style="padding: 10px; text-align: center;">Rally</th>
              <th style="padding: 10px; text-align: center;">Status</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>

      <!-- Action Button / Link to Google Sheet -->
      <div style="text-align: center; margin-top: 24px; padding-top: 18px; border-top: 1px solid #2a3050;">
        <a href="${sheetUrl}" target="_blank" style="background: linear-gradient(#ffd66b, #f2a93e); color: #2a1a10; font-weight: bold; text-decoration: none; padding: 12px 28px; border-radius: 8px; display: inline-block; font-size: 14px; box-shadow: 0 4px 14px rgba(242, 169, 62, 0.3);">
          📊 Open Google Spreadsheet (All Rows &amp; Columns)
        </a>
        <p style="color: #9aa1b8; font-size: 12px; margin-top: 14px; line-height: 1.5;">
          This is an automated consolidated digest sent once daily at the end of the day.<br>
          Target Recipient: <b style="color: #ffd66b;">${NOTIFICATION_EMAIL}</b> · Total All-Time Records: ${allRows.length}
        </p>
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log(`Daily summary email successfully sent to ${NOTIFICATION_EMAIL}`);
}

/**
 * Helper to set up automatic daily trigger (runs at 11:00 PM every day)
 * Run this function once from the Apps Script editor to automate.
 */
function setupDailyTrigger() {
  removeDailyTriggers();
  ScriptApp.newTrigger("sendDailySummaryEmail")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
  Logger.log("✅ Automatic daily summary trigger registered for 11:00 PM every day.");
}

/**
 * Remove existing daily triggers
 */
function removeDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendDailySummaryEmail") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log("Cleaned up existing daily summary triggers.");
}

/**
 * Test function to verify the daily summary email instantly
 */
function testDailySummaryEmail() {
  sendDailySummaryEmail();
}
