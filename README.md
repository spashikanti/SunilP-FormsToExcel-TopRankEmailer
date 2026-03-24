# SunilP-FormsToExcel-TopRankEmailer
Automated workflow for collecting evaluation responses, calculating the highest‑ranked instructor, and sending a notification email after:

✔ All submissions are received **before the deadline**, OR  
✔ The deadline has arrived (even if responses are incomplete).

This project integrates **Microsoft Forms → Excel → Office Scripts → Power Automate** to create a reliable, enterprise‑ready evaluation automation solution.

---
## ⬇️ Download

**Get the latest solution package (ZIP) from GitHub Releases:**

👉 [📦 Download Latest Release](https://github.com/spashikanti/SunilP-FormsToExcel-TopRankEmailer/releases/latest)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
![Power Platform](https://img.shields.io/badge/Platform-Power%20Platform-blue)
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/spashikanti/SunilP-FormsToExcel-TopRankEmailer?logo=github)](https://github.com/spashikanti/SunilP-FormsToExcel-TopRankEmailer/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/spashikanti/SunilP-FormsToExcel-TopRankEmailer/total?logo=github)](https://github.com/spashikanti/SunilP-FormsToExcel-TopRankEmailer/releases)

---

# 📘 Features

### ✅ Collect responses from Microsoft Forms  
Each form submission is automatically written into an Excel table.

### ✅ Admin-configurable evaluation logic  
Admin sets in Excel:
- **Expected number of submissions**
- **Deadline date/time**

### ✅ Intelligent flow outcome logic  
Office Script outputs one of three workflow statuses:

| Status             | Meaning |
|--------------------|---------|
| **Complete**       | All expected submissions received before deadline |
| **DeadlineReached**| Deadline passed; send results with available data |
| **InProgress**     | Waiting for more submissions and before deadline |

### ✅ Email automation  
Power Automate sends the email to the highest‑ranked instructor when:
- All submissions are in, OR  
- Deadline is reached (whichever happens first)

---

# 🧱 Solution Architecture
Microsoft Forms
↓
Excel Table (RatingsData)
↓
Office Script (Evaluation Status + Top Instructor)
↓
Power Automate Flow
↓
Conditional Email to Highest-Ranked Instructor

---

# 📊 Excel Structure

## Worksheet 1 — `RatingsData`
This sheet stores Microsoft Forms submissions.

| Column Name        | Description |
|--------------------|-------------|
| ParticipantEmail   | Email from form |
| Instructor         | Instructor evaluated |
| Score              | Numeric rating |
| Timestamp          | Submission time |

---

## Worksheet 2 — `Settings`
Admin-controlled configuration.

| Setting              | Value Example |
|----------------------|---------------|
| ExpectedSubmissions  | 20 |
| DeadlineDate         | 2026-03-31 23:59 |

⚠ **Admin may edit these values anytime.**

---

# 🧩 Office Script

Script name:  
### **`GetEvaluationStatusAndTopInstructor.ts`**

This script:
- Reads actual submissions  
- Reads expected submissions  
- Checks whether the deadline has passed  
- Computes the top instructor  
- Returns workflow status for Power Automate  

See script in `/office-scripts/GetEvaluationStatusAndTopInstructor.ts`.

---

# 🔁 Power Automate Flow

JSON Export located under:  
📁 `/power-automate/FlowDefinition.json`

### Flow Logic:
1. Trigger: *When a new form response is submitted*  
2. Add row to Excel table  
3. Run Office Script  
4. Condition block checks:
   - `status == "Complete"`  
   - `status == "DeadlineReached"`  
   - Otherwise: Do nothing  
5. Sends an email to the top-ranked instructor  
6. Includes their score and evaluation summary

---

# 🖼 Flow Diagram

![Flow Diagram](./media/Forms_To_Excel_TopRankEmailer_Architecture.png)

---

# ▶️ How to Deploy

### **Step 1 — Upload Excel file**
Upload `excel/SampleRatings.xlsx` to:
- OneDrive for Business or  
- SharePoint  

### **Step 2 — Import Office Script**
Copy the code into:
Excel → Automate Tab → New Script

### **Step 3 — Import Power Automate Flow**
Use:
Power Automate → My Flows → Import → Import Package

Reconnect Excel & Script resources.

### **Step 4 — Test**
Submit sample Forms data and observe:
- Script output  
- Status value  
- Email behavior

---

# 🧪 Sample Workflow Scenarios

### ✔ Scenario A — All 20 submissions received before 2026‑03‑31  
Flow status = **Complete**  
→ Email sent immediately

### ✔ Scenario B — Only 12 submissions received by deadline  
Flow status = **DeadlineReached**  
→ Email still sent (with best available data)

### ✔ Scenario C — Ongoing, before deadline  
Flow status = **InProgress**  
→ No email sent

---

# 🛠 Optional Enhancements

- Email CC to program coordinator  
- Multi-instructor averaging  
- Multi-form consolidation  
- Teams notification instead of email  
- Storage in Dataverse instead of Excel  
- Power BI dashboard integration  

---

## 🤝 Community & Contribution
As a **Power Platform Super User**, I built this to solve the deployment bottlenecks of heavy UI kits identified in the community forums. If you have ideas for new patterns (e.g., "wave" vs "pulse"), feel free to open a Pull Request!

---

# 🙌 Author
**Sunil Pashikanti**  
Community Contributor — Power Automate / Excel / Office Scripts

**Created by [Sunil Kumar Pashikanti](https://sunilpashikanti.com)** *Principal Architect | Power Platform Super User* **Blog:** [sunilpashikanti.blogspot.com](http://sunilpashikanti.blogspot.com)
