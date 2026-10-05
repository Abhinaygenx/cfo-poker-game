# CFO Leadership Pulse - Live Google Sheets & Daily Summary Email Integration

This system automatically records players who participate in the **CFO Leadership Pulse** assessment (**The Hot Seat / CFO Poker Table** and **CFO Desert Rally**) into a Google Spreadsheet, keeps their progress updated in real time across 4 dedicated tabs, and sends **one consolidated daily summary email at the end of the day** to **`mmsbf26001@stu.xim.edu.in`** (no annoying per-signup spam).

---

## 🚀 How It Works

1. **Player Registration & Real-Time Sheet Storage (`REGISTER` event):**
   - When a participant enters their name, email, phone, age, gender, status (Working / PG / UG), and nickname, their record is instantly appended to the Google Sheet (`Players_Summary`).
   - Individual emails are **not** sent per signup to avoid inbox clutter.

2. **Game 1 Completion (`POKER_COMPLETED` or `RALLY_COMPLETED`):**
   - When the player finishes the first game, their detailed stats and question responses are saved to the Google Sheet.
   - The end screen prompts **"Play Next Game ▶"** and automatically launches the unplayed game.

3. **Game 2 Completion & Final Summary:**
   - When the player completes the second game, their profile shows **"🏆 Both Games Completed! View Final Summary ▶"**.
   - Clicking redirects to the combined executive scorecard showing cross-game competencies, CFO readiness tier, and radar analysis.

4. **One Consolidated Daily Email (End of Day):**
   - Every night (between 11 PM and midnight, or trigger time), the script compiles all participants who played during the day, their scores, completion status, and a direct link to the Google Sheet, sending **one single executive summary email** to `mmsbf26001@stu.xim.edu.in`.

---

## 📋 2-Minute Setup Guide (Google Sheets & Apps Script)

You do **not** need any paid database or third-party service. This uses Google's native, free Apps Script engine.

### Step 1: Create a Google Spreadsheet
1. Open [Google Sheets](https://sheets.new).
2. Name your spreadsheet: **`CFO Leadership Pulse - Players Data`**.

### Step 2: Paste the Apps Script Code
1. In your new Google Sheet, click **Extensions > Apps Script** in the top menu.
2. In the code editor, delete any existing code.
3. Open [`google-apps-script.js`](./google-apps-script.js) from this folder, copy all of it, and paste it into the editor.
4. Click the **Save** (💾) icon.

### Step 3: Deploy as a Web App
1. At the top right of the Apps Script window, click the blue **Deploy** button > **New deployment**.
2. Click the gear icon (**Select type**) on the left and select **Web app**.
3. Configure the settings:
   - **Description**: `CFO Game Webhook`
   - **Execute as**: `Me` (your Google account)
   - **Who has access**: `Anyone` *(Crucial: allows the game frontend to submit player responses)*
4. Click **Deploy**.
5. When prompted, click **Authorize access**, choose your Google account, click *Advanced*, and click *Go to Untitled project (unsafe)* (this is standard for your own script), then click **Allow**.
6. Copy the **Web app URL** (it ends in `/exec`, e.g., `https://script.google.com/macros/s/AKfycb.../exec`).

### Step 4: Connect the URL to the Game
1. Open [`index.html`](./index.html) in your browser (Google Chrome or Microsoft Edge recommended).
2. Click the **⚙️ Cloud Setup** button at the top right of the screen (or in the footer).
3. Paste your Web App URL into the input field and click **Save & Test Connection**.
4. That's it! Your game is now live and connected.

### Step 5: Activate the End-of-Day Daily Summary Email Trigger (One-Time)
1. Go back to your **Apps Script editor**.
2. In the toolbar function dropdown (next to "Debug" / "Run"), select **`setupDailyTrigger`**.
3. Click **Run**.
4. That's all! Google Apps Script will now automatically run every night (between 11 PM and midnight) to send one single summary email to `mmsbf26001@stu.xim.edu.in`.
5. *(Optional test)*: Select **`testDailySummaryEmail`** and click **Run** anytime to send an instant test summary email to your inbox.


---

## 🗂️ Google Spreadsheet Tabs Created Automatically

The script automatically maintains **4 dedicated tabs** in your Google Spreadsheet:

### 1. `Players_Summary` (Overview Dashboard)
Stores one consolidated row per player, keeping their status, chip totals, and overall rankings updated in real time:
- **Last Updated**, **Participant ID**, **Full Name**, **Email**, **Phone**, **Age**, **Gender**, **Status (Working/PG/UG)**, **Nickname**, **Progress State**
- **Poker Stats:** Completed (Yes/No), Chips ($), Level, Overall %, Best Streak
- **Rally Stats:** Completed (Yes/No), Title, Score (out of 75), %, CFO Score, EQ Score, Crashes, Coins

---

### 2. `Poker_Responses` (Individual Poker Questions & Scores)
Logs all 10 hands/questions answered by every user in **The Hot Seat (Poker)** (5 Ethical Decision dilemmas + 5 Quantitative & Reasoning puzzles):
| Column | Description |
|---|---|
| **Recorded At** | Timestamp when game completed |
| **Participant ID** | Player's unique ID |
| **Player Name** | Full name |
| **Player Email** | Email address |
| **Hand #** | Hand index (1 to 10) |
| **Question Type** | CFO Dilemma / IQ puzzle |
| **Dimension** | Tested competency (Ethics, Governance, Financial Math, Cognitive Logic) |
| **Question / Statement** | Full text of the prompt / dilemma |
| **User Answer Code** | Option chosen (1 to 5) |
| **User Answer Text** | Complete text of user's chosen answer |
| **Target / Key** | Recommended CFO decision benchmark |
| **Distance / Diff** | Difference / distance from ideal response |
| **Hand Result** | **Win**, **Push**, or **Loss** |
| **Stake %** | Percentage of pot wagered (10% to 100% ALL IN) |
| **Stake Chips** | Bet amount in chips ($) |
| **Chips Won / Lost** | Profit (+) or loss (-) on the hand |
| **Chips After Hand** | Remaining pot balance after hand |
| **Time Taken (s)** | Seconds taken before locking in answer |
| **Streak After** | Consecutive win streak |

---

### 3. `Rally_Responses` (Individual Rally Questions & Scores)
Logs all 15 questions answered by every user in **CFO Desert Rally**:
| Column | Description |
|---|---|
| **Recorded At** | Timestamp when rally completed |
| **Participant ID** | Player's unique ID |
| **Player Name** | Full name |
| **Player Email** | Email address |
| **Question #** | Question sequence (1 to 15) |
| **Section / Category** | Category (CFO Strategy, Risk, Analysis, Leadership, Ethics / EQ Perception, Regulation, Utilization) |
| **Question / Statement** | Full statement text displayed over the race gates |
| **User Choice (1-5)** | Lane driven through (1 to 5) |
| **User Answer Label** | Strongly disagree, Disagree, Neutral, Agree, Strongly agree |
| **User Score Awarded (1-5)** | Scored points earned (1 to 5 points, accounting for reverse scoring) |
| **Reverse Scored?** | Yes / No |
| **Reaction Time (s)** | Time taken to steer and pass gate |

---

### 4. `Event_Log` (Audit Trail)
Records every single raw event payload with timestamp for auditability.

---

## 📧 Consolidated Daily Summary Email Format

Sent once daily at the end of the day to **`mmsbf26001@stu.xim.edu.in`**:
- **Date & Assessment Headline**: Executive digest of the day's testing cohort.
- **Key Metrics**: Total participants today, fully completed both games, and average performance.
- **Roster Table**: Name, Email, Status, Hot Seat level & chips, Desert Rally title & score, and final completion status.
- **One-Click Sheet Access**: Direct button linking to the live Google Spreadsheet.

---

## 🎮 Gameplay & User Experience Enhancements

1. **Seamless Progression Between Games:**
   - Finishing Game 1 displays a prominent **"Play Next Game ▶"** button and seamlessly loads Game 2.
   - When both games are complete, it shows **"🏆 Both Games Completed! View Final Summary ▶"** to inspect overall executive calibration.
2. **Desert Rally Cockpit Question Console:**
   - Moved from the top edge to the lower console below the racing car so the driver never has to look away from oncoming traffic.
   - Highlights the 5 Likert lanes with real-time feedback on your current lane.
3. **Optimized Traffic Density:**
   - Civilian car frequency reduced by 25% for a smoother, balanced challenge that keeps the focus on decision-making without cheap collisions.
4. **How to Play Guide in Desert Rally:**
   - Added a vertical "How to Play" tab matching the Poker style, providing complete controls, gate mechanics, scoring, and tips.

