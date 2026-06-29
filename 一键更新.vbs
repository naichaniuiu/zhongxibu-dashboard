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
gitExe = "C:\Users\wm881\.workbuddy\vendor\PortableGit\bin\git.exe"
logFilePath = strPath & "\update_log.txt"

' Step 1: Check Excel file
excelPath = "C:\Users\wm881\Downloads\Òµ¼¨ Ç·¿î¿´°å.xlsx"
If Not fso.FileExists(excelPath) Then
    MsgBox "ERROR: Excel file not found!" & vbCrLf & vbCrLf & "Please download the Excel file to:" & vbCrLf & excelPath, vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Run all commands in sequence and log output
cmdLine = "cmd /c cd /d """ & strPath & """ && " & _
    "echo === Dashboard Update: %date% %time% === > """ & logFilePath & """ && " & _
    "echo [1/4] Processing Excel data... >> """ & logFilePath & """ && " & _
    """"" & pythonPath & """ """ & strPath & "\process_data_v2.py"" >> """ & logFilePath & """ 2>&1 && " & _
    "echo [2/4] Extracting customer data... >> """ & logFilePath & """ && " & _
    """"" & pythonPath & """ """ & strPath & "\extract_customers.py"" >> """ & logFilePath & """ 2>&1 && " & _
    "echo [3/4] Generating dashboard HTML... >> """ & logFilePath & """ && " & _
    """"" & pythonPath & """ """ & strPath & "\gen_modal_dashboard.py"" >> """ & logFilePath & """ 2>&1 && " & _
    "echo [4/4] Pushing to GitHub... >> """ & logFilePath & """ && " & _
    """"" & gitExe & """ add -A >> """ & logFilePath & """ 2>&1 && " & _
    """"" & gitExe & """ commit -m ""Update dashboard %date%"" >> """ & logFilePath & """ 2>&1 && " & _
    """"" & gitExe & """ push origin main >> """ & logFilePath & """ 2>&1 && " & _
    "echo === Update completed === >> """ & logFilePath & """"

resultCode = WshShell.Run(cmdLine, 0, True)

' Check if index.html exists
If Not fso.FileExists(strPath & "\index.html") Then
    MsgBox "ERROR: Dashboard HTML generation failed!" & vbCrLf & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Show success message
MsgBox "Dashboard updated successfully!" & vbCrLf & vbCrLf & _
    "Managers can view at:" & vbCrLf & _
    "https://naichaniuiu.github.io/zhongxibu-dashboard/" & vbCrLf & vbCrLf & _
    "Note: GitHub Pages may take 1-2 minutes to refresh.", _
    vbInformation, "Update Done"

' Open the dashboard in browser
WshShell.Run "https://naichaniuiu.github.io/zhongxibu-dashboard/"
