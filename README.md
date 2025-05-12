# Task Schedule Dashboard (Google Sheets + GAS)

This project is a Google Apps Script (GAS)-powered dashboard for visualizing and managing workday availability, meeting load, and task scheduling capacity directly within Google Sheets.

---

## 📌 Features

* 🗓️ Auto-generates monthly calendar sheets (6 months ahead)
* 📊 Visual stats per day:

  * Total working hours
  * Meeting count and total time
  * Task-available time
  * Fill rate (%)
* 🎨 Color-coded availability:

  * Red = 80%+ busy
  * Yellow = 50–79% busy
  * Green = Less than 50%
* 🧠 Automatically excludes:

  * Weekends (gray)
  * Public holidays (gray from Google Calendar)
  * Personal vacation (light blue, editable in-sheet)

---

## ⚙️ Setup

### 1. Install clasp

```bash
npm install -g @google/clasp
```

### 2. Authenticate

```bash
clasp login
```

### 3. Create or clone your script project

```bash
clasp create --type sheets --title "TaskScheduleDashboard"
```

Or use:

```bash
clasp clone <YOUR_SCRIPT_ID>
```

### 4. Push local scripts to Apps Script

```bash
clasp push
```

---

## 📂 Project Structure

```
task-schedule-dashboard/
├── .clasp.json
├── .gitignore
├── README.md
└── src/
    ├── UiSetup.gs
    ├── SheetGenerator.gs
    ├── StatsUpdater.gs
    ├── ColoringRules.gs
    └── StylingEnhancer.gs
```

---

## 📌 Customization

* Edit `Param` sheet to adjust fill rate thresholds
* Add vacation days to `MyVacation` sheet
* Modify coloring logic in `StylingEnhancer.gs` if needed

---

## 🛡️ Notes

> ⚠️ **Timezone Disclaimer**: This dashboard assumes all data is handled in JST (Asia/Tokyo). If you are operating in another timezone or using a calendar with a different region setting, you may experience incorrect date/time calculations. For multi-region support, align timezones using `Session.getScriptTimeZone()` or make timezone configurable.

* Requires Google Workspace account for Calendar access
* Only confirmed or tentative events are counted
* All-day events are ignored *unless* they occupy 9–18 business hours

---

## 🧪 Example Use Case

Track when you're overwhelmed with meetings, find optimal days for focused work, or use it in 1:1s to communicate availability transparently.

---

MIT License
