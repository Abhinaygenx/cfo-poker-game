# CFO Leadership Pulse - Live Google Sheets & Email Integration

This system automatically records players who participate in the **CFO Leadership Pulse** assessment (**The Hot Seat / CFO Poker Table** and **CFO Desert Rally**) into a Google Spreadsheet, keeps their progress updated in real time, and sends instant email reports to **`mmsbf26001@stu.xim.edu.in`**.

---

## 🚀 How It Works

1. **Player Registration (`REGISTER` event):**
   - When a participant enters their name, email, phone, age, gender, status (Working / PG / UG), and nickname, their record is instantly appended to the Google Sheet.
   - An email alert is automatically dispatched to `mmsbf26001@stu.xim.edu.in` with the player's profile.

2. **Game 1 Completion (`POKER_COMPLETED` event):**
   - As soon as the player finishes the 30 hands in the Poker game, the script **finds that player's row** in the Google Sheet and updates it with:
     - Final Chip count
     - CFO Leadership level (e.g., *Chief Financial Officer, VP Finance*)
     - Overall Score percentage
     - Best streak & hands played
   - An updated scorecard email is sent to `mmsbf26001@stu.xim.edu.in`.

3. **Game 2 Completion (`RALLY_COMPLETED` event):**
   - When the player completes the Desert Rally, the script updates their row with:
     - Rally title (e.g., *Rally Champion, Podium Finisher*)
     - Overall Score (points & percentage)
     - CFO Skills score vs EQ score breakdown
     - Crashes & coins collected
   - An updated scorecard email is sent to `mmsbf26001@stu.xim.edu.in`.

4. **Final Assessment (`FINISHED` event):**
   - When the user reviews their final scorecard, their status is marked as *Completed* and a comprehensive final report is emailed to `mmsbf26001@stu.xim.edu.in`.

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
Logs all 30 hands/questions answered by every user in **The Hot Seat (Poker)**:
| Column | Description |
|---|---|
| **Recorded At** | Timestamp when game completed |
| **Participant ID** | Player's unique ID |
| **Player Name** | Full name |
| **Player Email** | Email address |
| **Hand #** | Hand index (1 to 30) |
| **Question Type** | Statement / IQ puzzle / CFO Dilemma |
| **Dimension** | Tested competency (Integrity, People, Vision, Rigor, Risk, Reasoning) |
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

## 📧 Email Alerts Format

Emails sent to `mmsbf26001@stu.xim.edu.in` include:
- A colored status badge (`New Player Registered`, `Hot Seat Completed`, `Desert Rally Completed`, or `Fully Completed`).
- Complete player profile details.
- Real-time performance breakdown for both games.
- A direct button link to open the live Google Spreadsheet with one click.
