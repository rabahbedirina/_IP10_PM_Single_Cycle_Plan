Option Explicit

Dim shell, fso, file
Dim csvPath, csvLine, fields
Dim MaintenancePlan, StartDate
Dim TAB_COUNT
Dim i


Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

csvPath = shell.CurrentDirectory & "\DATA_IP10_Single_Cycle_Plan.csv"
If Not fso.FileExists(csvPath) Then
    MsgBox "CSV file not found:" & vbCrLf & csvPath, vbCritical, "Error"
    WScript.Quit
End If

Set file = fso.OpenTextFile(csvPath, 1)

' Skip CSV header
If Not file.AtEndOfStream Then
    file.ReadLine
End If

MsgBox "Open SAP and make sure you re ready." & vbCrLf & _
       "Click OK to start.", _
       vbInformation, "SAP Automation"
Do Until file.AtEndOfStream

    csvLine = file.ReadLine

    If Trim(csvLine) <> "" Then

        fields = Split(csvLine, ";")

        If UBound(fields) >= 1 Then

            MaintenancePlan = Trim(fields(0))
            StartDate = Trim(fields(1))
         
            shell.AppActivate "SAP"
            WScript.Sleep 500
            shell.SendKeys "^/"
            WScript.Sleep 500
            shell.SendKeys "/nIP10"
            shell.SendKeys "{ENTER}"
            WScript.Sleep 500
            shell.SendKeys MaintenancePlan
            shell.SendKeys "{ENTER}"            
            WScript.Sleep 500            
            shell.SendKeys "{F9}"        
            shell.SendKeys StartDate            
            shell.SendKeys "{ENTER}"
            WScript.Sleep 500         
            shell.SendKeys "^s"
            WScript.Sleep 500

        End If

    End If

Loop

file.Close

MsgBox "All rows have been processed.", _
       vbInformation, "SAP Automation"

Set file = Nothing
Set fso = Nothing
Set shell = Nothing
