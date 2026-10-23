# SAP Plant Maintenance: IP10 Single Cycle Plan Automation

This project provides a lightweight automation utility using **VBScript** and a **CSV data file** to streamline the process of starting Single Cycle maintenance plans in **SAP Plant Maintenance (PM)** via transaction **IP10**.

---

## What is SAP Transaction IP10?

**IP10** is the standard SAP GUI transaction code used for **Maintaining Maintenance Plan Schedules**. 

In SAP Plant Maintenance, a **Maintenance Plan** defines *when* and *how often* maintenance tasks should be performed on technical objects (such as equipment or functional locations). 
* **Single Cycle Plans:** These are maintenance plans tied to a single time-based or performance-based cycle (e.g., every 30 days, or every 500 operating hours).
* **Starting a Plan ("Start in Cycle"):** Before a maintenance plan can automatically generate scheduling calls (maintenance orders or service notifications), it must be explicitly **started**. Starting a plan initializes its cycle from a designated start date, calculating the initial call horizon and scheduled dates.

Manually entering dozens or hundreds of plans in IP10 can be tedious and time-consuming. This script automates that repetitive data-entry workflow.

---

## Prerequisites & Requirements

1. **SAP GUI Installed:** Ensure the SAP Logon Pad is installed and configured on your machine.
2. **Active SAP Session:** You must be logged into your SAP system (though the script will attempt to navigate using `/nIP10`).
3. **Correct File Placement:** Both the VBScript file (`_IP10_PM_Single_Cycle_Plan.vbs`) and the CSV data file (`DATA_IP10_Single_Cycle_Plan.csv`) must reside in the **same directory**.

---

## File Structure

* **`_IP10_PM_Single_Cycle_Plan.vbs`**: The automation script that reads row data, opens SAP, navigates to IP10, enters the plan number, sets the start date via function keys (`{F9}` for start options), and saves.
* **`DATA_IP10_Single_Cycle_Plan.csv`**: The input file containing the list of maintenance plans and their respective start dates.

---

## CSV Data File Format (`DATA_IP10_Single_Cycle_Plan.csv`)

The CSV file uses a **semicolon (`;`)** delimiter. The first line is treated as a header and is automatically skipped by the script.

### Column Layout:
1. **`MaintenancePlan`**: The Maintenance Plan number (e.g., `10000234`).
2. **`StartDate`**: The start date for the cycle, formatted according to your SAP user profile date settings (e.g., `DD.MM.YYYY` or `MM/DD/YYYY`).

### Example CSV Content:
```csv
MaintenancePlan;StartDate
10000234;01.11.2026
10000235;15.11.2026
10000236;01.12.2026
```

---

## How to Use the Script

1. **Prepare Your Data:**
   * Open `DATA_IP10_Single_Cycle_Plan.csv` in a text editor or spreadsheet program (ensure you save it back as comma/semicolon-separated text with `.csv` extension).
   * Add your maintenance plan numbers and desired start dates.

2. **Launch SAP:**
   * Open your SAP GUI session and leave it open (the script will activate the window containing `"SAP"` in its title).

3. **Run the Script:**
   * Double-click on `_IP10_PM_Single_Cycle_Plan.vbs`.
   * A prompt will appear telling you to ensure SAP is ready. Click **OK**.
   * The script will loop through each row in your CSV, execute the IP10 transaction sequence, start the plan with the given date, and save it.

4. **Completion:**
   * Once all rows are processed, a final message box will notify you: *"All rows have been processed."*

---

## Troubleshooting & Tips

* **Window Focus:** The script relies on `shell.AppActivate "SAP"`. Make sure your active SAP window title includes the word "SAP" so it can successfully find the session.
* **Timing & Sleep Delays:** The script includes `WScript.Sleep 500` (0.5 seconds) pauses between keystrokes to prevent desynchronization. If your SAP server is running slow, you may want to slightly increase these sleep durations in the `.vbs` file.
* **Error Logs / Dry Runs:** Always test the script on a single test plan in a QA/Development SAP client before running it across a large production batch.