# EnviroVac Trainer Evaluations

A website trainers use on a phone, tablet or computer to complete skills evaluations.
**Save & Send** saves a PDF of the evaluation to SharePoint, adds a row to the Evaluation Log,
and emails envirovactraining@envirovac.us.

Site: **https://envirovac-vecta.github.io/trainer-evaluations/**

| File | What it does |
|---|---|
| `index.html` | the site (Evaluations, Current / Past / All Training Cycles tabs) |
| `history.json` | cohort history imported once from the Trainer Tracker (`tools/import_tracker.py`) |
| `forms.json` | the 10 evaluations: sections, items, rating scale |
| `trainers.txt` | trainer names for the picker, one per line |
| `build.py` | builds the site (trainer list, flow links, employee list) and merges the daily HSI cohort report into `history.json` |
| `.github/workflows/publish.yml` | publishes on every change and each morning |
| `vendor/` | PDF library (jsPDF + AutoTable) |

## Changing things
- **Trainer list:** edit `trainers.txt` on GitHub (pencil icon) and commit. Live in about a minute.
- **An evaluation's wording or items:** edit `forms.json` (or ask Claude to).

## Secrets (Settings → Secrets and variables → Actions)
| Secret | What |
|---|---|
| `EVAL_SUBMIT_URL` | HTTP URL of the "Evaluations - Save and Send" flow |
| `EVAL_LOG_URL` | HTTP URL of the "Evaluations - Read log" flow (shows **Submitted** on the Training Cycle tabs) |
| `EVAL_COHORT_URL` | HTTP URL of the "Cohort Report - Send file" flow (today's HSI Cohort Daily Report) |
| `COMMENT_SAVE_URL` | HTTP URL of the "Comments - Save" flow (trainer comments on the Training Cycle tabs) |
| `COMMENT_READ_URL` | HTTP URL of the "Comments - Read" flow (shows the comments) |
| `TRAINER_REPORT_URL` | optional - HTTP URL of the "Trainer Report - Save file" flow: the build sends a colored **Training Cycle Report.xlsx** (one row per person, a column per training) there every day, for the trainers in Project Polarity. Built by `tools/trainer_report.py`. |
| `EVAL_DATA_URL` | optional - HTTP URL of a flow that sends `HSI Data.xlsx` (Active employees for the picker) |

After adding or changing a secret: Actions → **Publish Evaluations Site** → **Run workflow**.

## Power Automate flow: "Evaluations - Save and Send"
1. **Trigger: When an HTTP request is received.** Who can trigger: Anyone. Leave the schema empty.
2. **Parse JSON.** Content: `triggerBody()` (Expression tab). Schema: click *Use sample payload to generate schema*
   and paste the sample below.
3. **Create file** (SharePoint). Site, and Folder Path =
   `/Shared Documents/Training Files to Upload to HSI/Evaluation PDFs/` + `officeFolder` (from Parse JSON).
   The office folder is created automatically the first time.
   - File Name: `fileName` (from Parse JSON) - "Employee - Evaluation - Office - Date.pdf"
   - File Content (Expression): `base64ToBinary(body('Parse_JSON')?['fileContent'])`
4. **Add a row into a table** (Excel Online (Business)). Point it at `Evaluation_Log.xlsx` in SharePoint, table `Evaluations`,
   and map each column from Parse JSON (Submitted At → submittedAt, Form ID → formId, Evaluation → formTitle,
   PDF File → fileName, and so on).
5. **Send an email (V2)** (or *Send an email from a shared mailbox (V2)* to send from the training mailbox).
   - To: `envirovactraining@envirovac.us`
   - Subject: `trainer` submitted `formTitle` for `employee`
   - Body: date, office, total/average score, items needing work, `needsList`, `notes`
   - Attachments (optional): Name `fileName`, Content (Expression) `base64ToBinary(body('Parse_JSON')?['fileContent'])`
6. **Response.** Status code `200`. Headers: `Access-Control-Allow-Origin` = `*`.

## Power Automate flow: "Evaluations - Read log"
1. **Trigger: When an HTTP request is received.** Who can trigger: Anyone.
2. **List rows present in a table** (Excel Online (Business)): `Evaluation_Log.xlsx`, table `Evaluations`.
3. **Response.** Status `200`. Headers: `Access-Control-Allow-Origin` = `*`. Body: `value` from step 2.
Copy the trigger URL into the `EVAL_LOG_URL` secret.

## Training Cycle tabs (called cohorts in HSI and the Trainer Tracker)
- **Intake Form** column (first column, for everyone): a date = an intake form sent from this site, or the date modified of the person's PDF in SharePoint > Completed Level Up Intake Forms (`intakeDate` in `history.json`); Done = the Trainer Tracker's "Intake Form Entered" says Complete.
- Each morning the build reads the **Cohort Daily Report** (HSI Envirovac Cohort Status Report, pulled by the
  `hsi_cohort` bot) and updates `history.json`: completions, incomplete/delinquent items, cohort, and anyone new
  to a cohort. People stay in the history after they drop off the report, so past cohorts are kept.
  The workflow commits the updated `history.json` back to the repo ("Daily cohort history update").
- **Current Training Cycles:** people whose training cycle hasn't ended. **Past Training Cycles:** ended, or no longer employed. **All Training Cycles:** everyone.
- Filters: find a person, Employees, Office, Training and Status (pick several in each), Training Cycle, Training cycle period, Show. Picking trainings shows only those columns and counts complete/missing on just those trainings.
- **Status** (Complete, Submitted, Incomplete): shows people with at least one shown training in a picked status and fades the other cells. Incomplete includes Delinquent and ?. Pair it with Training to see, for example, everyone still incomplete on Operator Training.
- **Export to Excel** downloads exactly what the table shows (same filters): one row per person with every shown training's status, plus a Filters sheet listing what was applied.
- Each training shows Complete (date), **Submitted** (an evaluation was sent from this site, waiting for HSI),
  Incomplete, Delinquent, or ? (person isn't in the HSI cohort report). Rows shaded light red: not recommended for promotion.
- Evaluations count toward: Operator Training, Driver Equipment Training (Part 1/2), Advance Mechanics Training,
  Basic Mechanical Troubleshooting, Basic Set Up, Lancing Automation Training, 24/40 Hour Hazwoper (hands on),
  Tech I to Tech II and Tech II to Tech III Skill Assessments.

Sample payload for step 2:
```json
{
  "formId": "operator-training",
  "formTitle": "Operator Training Skills Assessment – Vacuum Truck & Hydroblasting Operations",
  "trainer": "Trainer Name",
  "employee": "Employee Name",
  "employeeId": "ABC123",
  "office": "008-Tuscaloosa",
  "location": "Mill site",
  "evaluationDate": "2026-10-14",
  "submittedAt": "2026-10-14T16:00:00.000Z",
  "totalScore": 35,
  "maxScore": 50,
  "averageScore": 3.5,
  "itemsNeedingWork": 2,
  "needsList": "Safety Fundamentals: JHSA Review; Troubleshooting: Loss of Vacuum",
  "sectionScores": [{"section": "Safety Fundamentals", "rating": 4}],
  "notes": "",
  "officeFolder": "008-Tuscaloosa",
  "fileName": "Employee Name - Operator Training Skills Assessment - 008-Tuscaloosa - 2026-10-14.pdf",
  "fileContent": "JVBERi0xLjMK"
}
```

## How trainers use it
1. Pick their name and the employee (start typing to filter), check the date and location.
2. Pick the Training Cycle (Intake (all levels), Tech I to Tech II, Tech II to Tech III, Tech III to Jr. Operator, Jr. Operator to Sr. Operator,
   Sr. Operator to Crew Leader). Picking an employee jumps to their current training cycle. Then pick the evaluation.
3. Rate **every item** 1–5: 1–2 Needs Work, 3 Meets Standard, 4 Proficient, 5 Mastered. Bullets under an item are
   talking points. Items rated 1–2 are listed as "needing work" in the email, log and PDF. Add comments per section.
4. Trainer and employee sign with a finger. **Save & Send**.

Unsent work is saved on the device automatically, so a dropped connection or closed tab doesn't lose it.
**Download PDF** works any time, even without a connection to the flow.

## Intake form
The **Intake (all levels)** tab holds the Employee Intake Form for Initial Evaluation: current employee level, then every training from Tech I to Supervisor with Previously Completed, Previously Observed, Observed Date and a note. It sends like the evaluations (PDF to the office folder, Evaluation Log row with Total Score = trainings completed or observed, email).

## Trainer comments
**Initial Comments:** the intake form's comment box is **Initial Comments** and is required. When an intake form is sent, that text is also saved to the Comments table as the employee's initial comment (Comment column starts with `Initial Comment: `), shown first and marked in orange. The **Initial Comments** tab lists everyone in the training cycle that started in September, to catch up their initial comment (Show: missing / has one / everyone).

**Comments tab:** pick any employee to see the needs identified in their evaluations and their comments, and add a comment; below is a list of recent comments from everyone (click one to open that employee).
**Needs identified in evaluations** (also under each employee on the Training Cycle tabs and in Export to Excel) come from the Evaluation Log's needs list column (the items rated 1-2 on each evaluation).

Click an employee's name on a Training Cycle tab: under Completed / Submitted / Missing is **Comments** (newest first, with date
and trainer) and a box to add one. Comments are saved to the `Comments` table and show for everyone. **Show > Has comments**
filters to people with comments; Export to Excel has a Comments column. The site is public (unlisted), so anyone with the
link can read the comments.

Table `Comments` (Excel, e.g. a new sheet in Evaluation_Log.xlsx), columns: Date, Employee, Employee ID, Office, Training Cycle, Trainer, Comment.

**Comments - Save:** When an HTTP request is received (Anyone; schema from the sample below) > Add a row into a table
(map each column: Date -> date, Employee -> employee, Employee ID -> employeeId, Office -> office, Training Cycle -> trainingCycle,
Trainer -> trainer, Comment -> comment) > Response 200 with header `Access-Control-Allow-Origin: *`.

**Comments - Read:** When an HTTP request is received (Anyone) > List rows present in a table (Comments) > Response 200,
header `Access-Control-Allow-Origin: *`, Body: `value` from List rows.

Sample payload:
```json
{"date": "2026-10-21 11:38", "employee": "Gregory Partlow", "employeeId": "EOAUH5AVJ", "office": "009-Columbus",
 "trainingCycle": "Jr. Operator to Sr. Operator", "trainer": "Noah", "comment": "Completed Part 1 today."}
```
