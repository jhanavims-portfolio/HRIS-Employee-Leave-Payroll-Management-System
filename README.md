<h1 align="center">📊 HRIS: Employee Leave & Payroll Management System</h1>

<p align="center"><b>An Excel-based HR system that tracks leave and automatically flags Loss of Pay (LOP)</b><br>
Microsoft Excel · Excel Tables · XLOOKUP · Conditional Formatting</p>

> ### 👀 Want to see the system?
> **Open the `screenshots` folder in the `main` branch of this repository** to see the Employee Master and the Leave Tracker & Payroll sheets. The full workbook is in the repository files.

---

## 🎯 At a glance

| 👥 Employees | ✅ Active | 📉 Turnover | 💸 LOP days flagged | ⚠️ Employees with LOP |
|:---:|:---:|:---:|:---:|:---:|
| **25** | **20** | **20%** | **13** | **6 of 25** |

---

## 📌 Project overview

A mini Human Resource Information System (HRIS) built in Excel. It keeps employee records in one place, tracks monthly leave, and calculates **Loss of Pay (LOP) days** and the **Net Remaining Leave Balance** automatically, so HR does less manual record-keeping.

## 🔄 How it works

| Step | Stage | What happens |
|:---:|---|---|
| 1 | 👤 Employee Master | Employee records stored centrally |
| 2 | 📅 Leave Tracker | Casual, sick and unexcused leave entered per employee |
| 3 | 🔗 Lookup | Name and department pulled automatically from the master using the Emp ID |
| 4 | 🧮 Calculation | Excess leave, LOP days and net balance calculated by formula |
| 5 | 💰 Payroll impact | LOP days show which employees face a salary deduction |

---

## 📂 Sheet 1: Employee Master

25 employee records across 5 departments.

| Field | Purpose |
|---|---|
| Emp ID | Unique identifier |
| Name | Employee name |
| Department | Operations, Engineering, Marketing, HR, Sales |
| Designation | Job role |
| Date of Joining | Tenure |
| Reporting Manager | Manager's Emp ID |
| Salary Band | Band A, B or C |
| Status | Active or Terminated |
| Performance Rating | Rating out of 5 |

**Summary block on the sheet:**

| Metric | Value |
|---|:---:|
| Total headcount | 25 |
| Active employees | 20 |
| Terminated employees | 5 |
| Turnover rate | 20% |

**Headcount by department:** Operations 8 · Engineering 6 · Marketing 4 · HR 4 · Sales 3

---

## 📂 Sheet 2: Leave Tracker & Payroll

Each employee gets a monthly allowance of **2 casual days and 1 sick day**. Leave beyond the allowance, plus unexcused absences, becomes Loss of Pay.

### The formulas

| Column | Formula logic |
|---|---|
| Name, Department | `XLOOKUP(Emp ID, EmployeeMaster[Emp ID], EmployeeMaster[Name or Department])` |
| Total Leaves Taken | `SUM(Casual Taken : Unexcused Absences)` |
| Casual LOP | `MAX(0, Casual Taken − Casual Allowed)` |
| Sick LOP | `MAX(0, Sick Taken − Sick Allowed)` |
| **LOP Days** | `Unexcused Absences + Casual LOP + Sick LOP` |
| **Net Remaining Leave Balance** | `(Casual Allowed + Sick Allowed) − (Casual Taken + Sick Taken)` |

A negative net balance means the employee used more leave than allowed.

### Worked example: Priya Nair (EMP004)

| Item | Allowed | Taken | Result |
|---|:---:|:---:|:---:|
| Casual leave | 2 | 4 | Casual LOP = **2** |
| Sick leave | 1 | 2 | Sick LOP = **1** |
| Unexcused absences | n/a | 1 | **1** |
| **Total LOP days** | | | **4** |
| **Net remaining balance** | | | **−3** |

---

## 📈 Results from the sample data

| Metric | Result |
|---|:---:|
| Casual leave taken | 39 days |
| Sick leave taken | 14 days |
| Unexcused absences | 4 days |
| Total leave taken | 57 days |
| Total LOP days | **13** (7 casual + 2 sick + 4 unexcused) |
| Employees with LOP | **6 of 25** |
| Employees with a negative balance | 5 |

**LOP days by department**
```
Operations   █████  5
Marketing    ████   4
Sales        ██     2
Engineering  █      1
HR           █      1
```

**Highest LOP:** Priya Nair and Amitav Sen, 4 days each.

---

## ⚙️ Features

✅ Central employee master  ·  ✅ Automatic name and department lookup  ·  ✅ Casual and sick leave tracking  ·  ✅ Automatic LOP calculation  ·  ✅ Net leave balance  ·  ✅ Excel tables with structured references  ·  ✅ Conditional formatting (data bars and colour scales)

## 🛠️ Tools used

| Tool | Used for |
|---|---|
| Microsoft Excel | Building the system |
| Excel Tables | Structured data and formulas that grow with new rows |
| XLOOKUP | Pulling employee details from the master (needs Excel 2021 or Microsoft 365) |
| SUM and MAX | Leave totals, LOP and balance |
| Conditional formatting | Highlighting high values at a glance |

---

## 💡 Business value

- Less manual HR record-keeping
- One source of truth for employee data
- Consistent LOP calculation with no manual errors
- Quick view of who is over their leave allowance
- Clear starting point for payroll deductions

## 🚀 Future improvements

📊 HR analytics dashboard · 📈 Power BI integration · 💰 LOP salary-deduction amounts · 📧 Leave approval notifications · 👥 Employee self-service portal · 📅 Attendance tracking · 📄 Payslip generation · ✅ Data validation dropdowns

---

## 📁 What's in this repository

| Item | Contents |
|---|---|
| Excel workbook | The complete HRIS system (2 sheets) |
| `screenshots/` | Employee Master and Leave Tracker & Payroll views |

## 📸 Screenshots

### Screenshot 1: Employee Master


### Screenshot 2: Leave Tracker and Payroll


---

## 👤 Author
**M.S. Jhanavi** · MBA (HR), KL University

## ⭐ Conclusion
This project shows how Microsoft Excel can support everyday HR operations. It combines employee records, leave tracking and automatic LOP calculation in one mini HRIS, and it identifies which employees face a payroll impact from unpaid leave.
