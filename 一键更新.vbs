' ============================================================
'  Zhongxibu Dashboard - One-Click Update + Auto Push to GitHub
'  Double-click this file to: Read Excel -> Generate HTML -> Push to GitHub
'  Managers just visit: https://naichaniuiu.github.io/zhongxibu-dashboard/
' ============================================================

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

pythonPath = "C:\Users\wm881\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
gitExe = """C:\Users\wm881\.workbuddy\vendor\PortableGit\bin\git.exe"""

' Log file
logFile = strPath & "\update_log.txt"
Set log = fso.CreateTextFile(logFile, True)
log.WriteLine "=== Dashboard Update: " & Now & " ==="
log.WriteLine ""

' Step 1: Check Excel file
excelPath = "C:\Users\wm881\Downloads\业绩 欠款看板.xlsx"
If Not fso.FileExists(excelPath) Then
    log.Close
    MsgBox "ERROR: Excel file not found!" & vbCrLf & vbCrLf & "Please download the Excel file to:" & vbCrLf & excelPath, vbCritical, "Update Failed"
    WScript.Quit 1
End If
log.WriteLine "[OK] Excel file found"

' Step 2: process_data_v2.py
log.WriteLine "[1/4] Processing Excel data..."
WshShell.Run "cmd /c """ & pythonPath & """ """ & strPath & "\process_data_v2.py"" >> """ & logFile & """ 2>&1", 0, True

' Step 3: extract_customers.py
log.WriteLine "[2/4] Extracting customer data..."
WshShell.Run "cmd /c """ & pythonPath & """ """ & strPath & "\extract_customers.py"" >> """ & logFile & """ 2>&1", 0, True

' Step 4: gen_modal_dashboard.py
log.WriteLine "[3/4] Generating dashboard HTML..."
WshShell.Run "cmd /c """ & pythonPath & """ """ & strPath & "\gen_modal_dashboard.py"" >> """ & logFile & """ 2>&1", 0, True

' Check if index.html was generated
If Not fso.FileExists(strPath & "\index.html") Then
    log.Close
    MsgBox "ERROR: Dashboard HTML generation failed!" & vbCrLf & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If
log.WriteLine "[OK] index.html generated"

' Step 5: Git push
log.WriteLine "[4/4] Pushing to GitHub..."
gitCmd = gitExe & " add -A && " & gitExe & " commit -m ""Update dashboard " & Date & """ && " & gitExe & " push origin main"
WshShell.Run "cmd /c cd /d """ & strPath & """ && " & gitCmd & " >> """ & logFile & """ 2>&1", 0, True
log.WriteLine "[OK] Pushed to GitHub"

log.WriteLine ""
log.WriteLine "=== Update completed: " & Now & " ==="
log.Close

' Success message
result = MsgBox("Dashboard updated successfully!" & vbCrLf & vbCrLf & _
    "Managers can view at:" & vbCrLf & _
    "https://naichaniuiu.github.io/zhongxibu-dashboard/" & vbCrLf & vbCrLf & _
    "Note: GitHub Pages may take 1-2 minutes to refresh.", _
    vbInformation, "Update Done")

' Open the dashboard in browser
WshShell.Run "https://naichaniuiu.github.io/zhongxibu-dashboard/"
