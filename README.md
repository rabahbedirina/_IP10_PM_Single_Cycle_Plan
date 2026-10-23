# SAP Plant Maintenance: IP10 Single Cycle Plan Automation

**Current Version:** `v2.0.0`

**Author:** Rabah BEDIRINA

**GitHub:** [rabahbedirina](https://github.com/rabahbedirina/)

This project provides a robust, production-ready automation utility using **VBScript** and a **CSV data file** to streamline the process of starting Single Cycle maintenance plans in **SAP Plant Maintenance (PM)** via transaction **IP10**.

## What is SAP Transaction IP10?

**IP10** is the standard SAP GUI transaction code used for **Maintaining Maintenance Plan Schedules**.

In SAP Plant Maintenance, a **Maintenance Plan** defines *when* and *how often* maintenance tasks should be performed on technical objects (such as equipment or functional locations).

* **Single Cycle Plans:** These are maintenance plans tied to a single time-based or performance-based cycle (e.g., every 30 days, or every 500 operating hours).

* **Starting a Plan ("Start in Cycle"):** Before a maintenance plan can automatically generate scheduling calls (maintenance orders or service notifications), it must be explicitly **started**. Starting a plan initializes its cycle from a designated start date, calculating the initial call horizon and scheduled dates.

Manually entering dozens or hundreds of plans in IP10 can be tedious and time-consuming. This script automates that repetitive data-entry workflow safely and securely.

## What's New in Version 2.0.0?

* **Strict Header & Date Validation:** Automatically checks that CSV headers match `MaintenancePlan;StartDate` and ensures start dates strictly follow the `DD.MM.YYYY` format using Regular Expressions and `DateSerial` verification.

* **Audit Trail & Logging:** Generates a detailed tab-delimited log file (`DATA_processing.log`) capturing row-by-row execution status (`SENT`, `DRY RUN`, or `ERROR`), timestamps, and error details.

* **Dry Run Mode:** Allows you to validate your CSV file syntax and dates without sending any keystrokes or modifying data in SAP.

* **Safe Character Escaping:** Automatically escapes special characters in plan numbers or dates to prevent SAP `SendKeys` injection errors.

## Prerequisites & Requirements

1. **SAP GUI Installed:** Ensure the SAP Logon Pad is installed and configured on your machine.

2. **Active SAP Session:** You must be logged into your SAP system.

3. **File Placement:** By default, both the VBScript file and the CSV data file (`DATA_IP10_Single_Cycle_Plan.csv`) should reside in the **same directory**.

## File Structure

* **`_IP10_PM_Single_Cycle_Plan.vbs`**: The core automation script with built-in error handling, logging, and validation.

* **`DATA_IP10_Single_Cycle_Plan.csv`**: The input file containing maintenance plans and start dates.

* **`DATA_processing.log`**: Automatically generated log file tracking execution results per row.

## CSV Data File Format (`DATA_IP10_Single_Cycle_Plan.csv`)

The CSV file uses a **semicolon (`;`)** delimiter.

* **Required Header:** `MaintenancePlan;StartDate`

* **Date Format:** Must be exactly `DD.MM.YYYY` (e.g., `01.01.2026`).

### Example CSV Content:

```
MaintenancePlan;StartDate
12345678;01.01.2026
87654321;15.01.2026

```

## How to Use the Script

### Method 1: Live Run (GUI Mode)

1. **Prepare Your Data:** Open `DATA_IP10_Single_Cycle_Plan.csv` in a text editor or spreadsheet program and add your target plans and dates using the `MaintenancePlan;StartDate` format.

2. **Launch SAP:** Open your SAP GUI session and leave it open. The script will maximize and activate the SAP window.

3. **Execute:** Double-click `_IP10_PM_Single_Cycle_Plan.vbs`.

4. **Confirm:** Click **OK** on the prompt and **do not touch your mouse or keyboard** while the script runs.

5. **Review Log:** Once completed, inspect `DATA_processing.log` to confirm successful processing.

### Method 2: Dry Run (Command Line / Safety Check)

To test your CSV file structure and date formatting without interacting with SAP, open your terminal (`cmd`) in the script directory and run:

```
cscript _IP10_PM_Single_Cycle_Plan_2.vbs /dryrun DATA_IP10_Single_Cycle_Plan.csv DATA_processing.log

```

## Troubleshooting & Tips

* **Window Focus:** The script looks for an SAP window containing `"SAP"` in its title. Ensure your SAP session is active.

* **Log Analysis:** If a row fails due to an invalid date format or missing plan number, check `DATA_processing.log` to see the exact row number and failure reason.

* **Timing & Delays:** The script includes a 500ms delay (`STEP_DELAY`) between steps. If your network or SAP instance is slow, you can adjust `STEP_DELAY` at the top of the `.vbs` script.

* **Testing:** Always test the script on a single test plan in a QA/Development SAP client before running it across a large production batch.