# EnviroVac Trainer Evaluations

A website trainers use on a phone, tablet or computer to complete skills evaluations.
**Save & Send** saves a PDF of the evaluation to SharePoint, adds a row to the Evaluation Log,
and emails envirovactraining@envirovac.us.

Site: **https://envirovac-vecta.github.io/trainer-evaluations/**

| File | What it does |
|---|---|
| `index.html` | the site (Evaluations, Current Cohorts, Past Cohorts tabs) |
| `history.json` | cohort history imported once from the Trainer Tracker (`tools/import_tracker.py`) |
| `forms.json` | the 10 evaluations: sections, items, rating scale |
| `trainers.txt` | trainer names for the picker, one per line |
| `build.py` | builds the site (adds trainer list, flow link, employee list) |
| `.github/workflows/publish.yml` | publishes on every change and each morning |
| `vendor/` | PDF library (jsPDF + AutoTable) |

## Changing things
- **Trainer list:** edit `trainers.txt` on GitHub (pencil icon) and commit. Live in about a minute.
- **An evaluation's wording or items:** edit `forms.json` (or ask Claude to).

## Secrets (Settings → Secrets and variables → Actions)
| Secret | What |
|---|---|
| `EVAL_SUBMIT_URL` | HTTP URL of the "Evaluations - Save and Send" flow |
| `EVAL_LOG_URL` | HTTP URL of the "Evaluations - Read log" flow (shows **Submitted** on the cohort tabs) |
| `EVAL_DATA_URL` | optional - HTTP URL of a flow that sends `HSI Data.xlsx` (Active employees for the picker) |

After adding or changing a secret: Actions → **Publish Evaluations Site** → **Run workflow**.

## Power Automate flow: "Evaluations - Save and Send"
1. **Trigger: When an HTTP request is received.** Who can trigger: Anyone. Leave the schema empty.
2. **Parse JSON.** Content: `triggerBody()` (Expression tab). Schema: click *Use sample payload to generate schema*
   and paste the sample below.
3. **Create file** (SharePoint). Site + folder for the PDFs.
   - File Name: `fileName` (from Parse JSON)
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

## Cohort tabs
- **Current Cohorts:** people whose cohort period hasn't ended. **Past Cohorts:** ended, or no longer employed.
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
  "fileName": "2026-10-14 Employee Name - Operator Training Skills Assessment - Trainer Name.pdf",
  "fileContent": "JVBERi0xLjMK"
}
```

## How trainers use it
1. Pick their name and the employee (start typing to filter), check the date and location.
2. Pick the evaluation. Rate each section 1–5, tap **Needs?** on any item that needs more work, add comments.
3. Trainer and employee sign with a finger. **Save & Send**.

Unsent work is saved on the device automatically, so a dropped connection or closed tab doesn't lose it.
**Download PDF** works any time, even without a connection to the flow.
