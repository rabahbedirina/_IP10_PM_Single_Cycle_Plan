' ==============================================================================
' Script Name: _IP10_PM_Single_Cycle_Plan.vbs
' Description: Automated SAP PM IP10 Single Cycle Plan Execution
' Version:     2.0.0
' Author:      Rabah BEDIRINA
' GitHub:      https://github.com/rabahbedirina/
' ==============================================================================

Option Explicit

Const DATE_PATTERN  = "^\d{2}\.\d{2}\.\d{4}$"
Const STEP_DELAY    = 500
Const STOP_ON_ERROR = True

Dim shell, fso, file, logFile, dateRegex
Dim csvPath, logPath, csvLine, fields, header, headerValid
Dim MaintenancePlan, StartDate
Dim rowNumber, rowError, dryRun, aborted, logErr

Set shell = CreateObject("WScript.Shell")

csvPath = shell.CurrentDirectory & "\DATA_IP10_Single_Cycle_Plan.csv"
logPath = shell.CurrentDirectory & "\DATA_processing.log"
dryRun = False
aborted = False

If WScript.Arguments.Count > 0 Then
    If LCase(WScript.Arguments(0)) = "/dryrun" And WScript.Arguments.Count = 3 Then
        dryRun = True
        csvPath = WScript.Arguments(1)
        logPath = WScript.Arguments(2)
    Else
        WScript.Echo "Usage: cscript DATA_IP10_Single_Cycle_Plan.vbs [/dryrun <csvPath> <logPath>]"
        WScript.Quit 1
    End If
End If


Set fso = CreateObject("Scripting.FileSystemObject")
Set dateRegex = New RegExp
dateRegex.Pattern = DATE_PATTERN

If Not fso.FileExists(csvPath) Then
    ShowMessage "CSV file not found:" & vbCrLf & csvPath, dryRun, vbCritical
    WScript.Quit 1
End If

On Error Resume Next
Err.Clear
Set file = fso.OpenTextFile(csvPath, 1, False)
If Err.Number <> 0 Then
    logErr = Err.Description
    On Error GoTo 0
    ShowMessage "Could not open CSV file:" & vbCrLf & csvPath & vbCrLf & logErr, dryRun, vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

On Error Resume Next
Err.Clear
Set logFile = fso.OpenTextFile(logPath, 8, True)
If Err.Number <> 0 Then
    logErr = Err.Description
    On Error GoTo 0
    file.Close
    ShowMessage "Could not open log file:" & vbCrLf & logPath & vbCrLf & logErr, dryRun, vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

logFile.WriteLine "Run started: " & CStr(Now)
logFile.WriteLine "CSV row" & vbTab & "Status" & vbTab & "MaintenancePlan" & vbTab & _
                  "StartDate" & vbTab & "Details"

If file.AtEndOfStream Then
    LogRow 1, "ERROR", "", "", "CSV file is empty"
    CloseAndQuit "The CSV file is empty."
End If

header = file.ReadLine
header = Replace(header, ChrW(239) & ChrW(187) & ChrW(191), "")
fields = Split(header, ";")
headerValid = False
If UBound(fields) = 1 Then
    If LCase(Trim(fields(0))) = "maintenanceplan" And _
       LCase(Trim(fields(1))) = "startdate" Then
        headerValid = True
    End If
End If
If Not headerValid Then
    LogRow 1, "ERROR", "", "", "Header must be: MaintenancePlan;StartDate"
    CloseAndQuit "Invalid header. Expected: MaintenancePlan;StartDate"
End If

rowNumber = 1
If Not dryRun Then
    MsgBox "Open SAP and make sure you're ready." & vbCrLf & _
           "Do not touch the keyboard or mouse while the script runs." & vbCrLf & _
           "Click OK to start.", vbInformation, "SAP Automation"
    WScript.Sleep STEP_DELAY
    shell.SendKeys "% x"
    WScript.Sleep STEP_DELAY
End If

Do Until file.AtEndOfStream Or aborted
    csvLine = file.ReadLine
    rowNumber = rowNumber + 1

    If Trim(csvLine) <> "" Then
        fields = Split(csvLine, ";")

        If UBound(fields) <> 1 Then
            LogRow rowNumber, "ERROR", "", "", "Expected exactly 2 semicolon-delimited fields"
            If STOP_ON_ERROR Then aborted = True
        Else
            MaintenancePlan = Trim(fields(0))
            StartDate = Trim(fields(1))

            If MaintenancePlan = "" Then
                LogRow rowNumber, "ERROR", MaintenancePlan, StartDate, "Maintenance plan is empty"
                If STOP_ON_ERROR Then aborted = True
            ElseIf Not IsValidDate(StartDate, dateRegex) Then
                LogRow rowNumber, "ERROR", MaintenancePlan, StartDate, _
                       "Start date must be a valid date in dd.mm.yyyy format"
                If STOP_ON_ERROR Then aborted = True
            ElseIf dryRun Then
                LogRow rowNumber, "DRY RUN", MaintenancePlan, StartDate, _
                       "Row validated; SAP actions were skipped"
            ElseIf Not ProcessSapRow(MaintenancePlan, StartDate, rowError) Then
                LogRow rowNumber, "ERROR", MaintenancePlan, StartDate, rowError
                If STOP_ON_ERROR Then aborted = True
            Else
                LogRow rowNumber, "SENT", MaintenancePlan, StartDate, _
                       "Keystrokes sent; SAP save was not independently verified"
            End If
        End If
    End If
Loop

If aborted Then
    logFile.WriteLine "Run aborted at CSV row " & rowNumber & ": " & CStr(Now)
Else
    logFile.WriteLine "Run finished: " & CStr(Now)
End If

file.Close
logFile.Close
If Not dryRun Then
    If aborted Then
        MsgBox "Stopped at CSV row " & rowNumber & ". Check the log:" & vbCrLf & logPath, _
               vbExclamation, "SAP Automation"
    Else
        MsgBox "All rows have been processed. Check the log:" & vbCrLf & logPath, _
               vbInformation, "SAP Automation"
    End If
End If

Set file = Nothing
Set logFile = Nothing
Set dateRegex = Nothing
Set fso = Nothing
Set shell = Nothing

Sub ShowMessage(ByVal message, ByVal useConsole, ByVal icon)
    If useConsole Then
        WScript.Echo message
    Else
        MsgBox message, icon, "SAP Automation"
    End If
End Sub

Sub CloseAndQuit(ByVal message)
    logFile.WriteLine "Run aborted: " & CStr(Now)
    logFile.Close
    file.Close
    ShowMessage message, dryRun, vbCritical
    WScript.Quit 1
End Sub

Sub LogRow(ByVal csvRow, ByVal status, ByVal maintenancePlanValue, _
           ByVal startDateValue, ByVal details)
    logFile.WriteLine CStr(csvRow) & vbTab & CleanLogField(status) & vbTab & _
                      CleanLogField(maintenancePlanValue) & vbTab & _
                      CleanLogField(startDateValue) & vbTab & _
                      CleanLogField(details)
End Sub

Function CleanLogField(ByVal value)
    value = Replace(CStr(value), vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    CleanLogField = value
End Function

Function IsValidDate(ByVal value, ByVal regex)
    Dim dayPart, monthPart, yearPart, parsedDate

    IsValidDate = False
    If Not regex.Test(value) Then Exit Function

    dayPart = CInt(Left(value, 2))
    monthPart = CInt(Mid(value, 4, 2))
    yearPart = CInt(Right(value, 4))

    On Error Resume Next
    Err.Clear
    parsedDate = DateSerial(yearPart, monthPart, dayPart)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    IsValidDate = Day(parsedDate) = dayPart And _
                  Month(parsedDate) = monthPart And Year(parsedDate) = yearPart
End Function

' Escape characters that SendKeys treats as special.
Function EscapeKeys(ByVal s)
    Dim c, k, result
    result = ""
    For k = 1 To Len(s)
        c = Mid(s, k, 1)
        If InStr("+^%~(){}[]", c) > 0 Then
            result = result & "{" & c & "}"
        Else
            result = result & c
        End If
    Next
    EscapeKeys = result
End Function

Function SendKeysChecked(ByVal keys, ByVal stepDescription, ByRef errorDetail)
    SendKeysChecked = False
    On Error Resume Next
    Err.Clear
    shell.SendKeys keys
    If Err.Number <> 0 Then
        errorDetail = stepDescription & " failed: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    WScript.Sleep STEP_DELAY
    SendKeysChecked = True
End Function

Function ProcessSapRow(ByVal maintenancePlanValue, ByVal startDateValue, _
                       ByRef errorDetail)
    ProcessSapRow = False
    WScript.Sleep STEP_DELAY

    If Not SendKeysChecked("^/", "Opening SAP command field", errorDetail) Then Exit Function
    If Not SendKeysChecked("/nIP10", "Entering transaction IP10", errorDetail) Then Exit Function
    If Not SendKeysChecked("{ENTER}", "Opening transaction IP10", errorDetail) Then Exit Function

    If Not SendKeysChecked(EscapeKeys(maintenancePlanValue), _
                           "Entering maintenance plan", errorDetail) Then Exit Function
    If Not SendKeysChecked("{ENTER}", "Opening plan", errorDetail) Then Exit Function

    If Not SendKeysChecked("{F9}", "Opening start date", errorDetail) Then Exit Function
    If Not SendKeysChecked(EscapeKeys(startDateValue), _
                           "Entering start date", errorDetail) Then Exit Function
    If Not SendKeysChecked("{ENTER}", "Confirming start date", errorDetail) Then Exit Function

    If Not SendKeysChecked("^s", "Saving maintenance plan", errorDetail) Then Exit Function

    ProcessSapRow = True
End Function
