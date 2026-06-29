' ============================================================
'  Zhongxibu Dashboard - One-Click Update + Auto Push to GitHub
'  Double-click this file to: Read Excel -> Generate HTML -> Push to GitHub
'  Managers just visit: https://naichaniuniu.github.io/zhongxibu-dashboard/
' ============================================================

Dim WshShell, fso, strPath, q
Dim pythonPath, gitExe, logFilePath, excelPath
Dim pyCmd, gitAddCmd, gitCommitCmd, gitPushCmd, diffCmd
Dim hasChanges

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

q = Chr(34)
pythonPath = "C:\Users\wm881\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
gitExe = "C:\Users\wm881\.workbuddy\vendor\PortableGit\bin\git.exe"
logFilePath = strPath & "\update_log.txt"

' Step 1: Check Excel file
excelPath = "C:\Users\wm881\Downloads\??? ?????.xlsx"
If Not fso.FileExists(excelPath) Then
    MsgBox "ERROR: Excel file not found!" & vbCrLf & vbCrLf & "Please download the Excel file to:" & vbCrLf & excelPath, vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Helper function to run a command and log output
' cmdArgs must already be properly quoted if it contains spaces
Function RunCommand(stepName, cmdArgs)
    Dim fullCmd, exitCode
    fullCmd = "cmd /c echo [" & stepName & "] >> " & q & logFilePath & q & " 2>&1 && " & cmdArgs & " >> " & q & logFilePath & q & " 2>&1"
    exitCode = WshShell.Run(fullCmd, 0, True)
    RunCommand = exitCode
End Function

' Initialize log file
WshShell.Run "cmd /c echo === Dashboard Update: %date% %time% === > " & q & logFilePath & q, 0, True

' Step 2: Process Excel data
pyCmd = q & pythonPath & q & " " & q & strPath & "\process_data_v2.py" & q
If RunCommand("1/4 Processing Excel data", pyCmd) <> 0 Then
    MsgBox "ERROR: Failed to process Excel data." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Step 3: Extract customer data
pyCmd = q & pythonPath & q & " " & q & strPath & "\extract_customers.py" & q
If RunCommand("2/4 Extracting customer data", pyCmd) <> 0 Then
    MsgBox "ERROR: Failed to extract customer data." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Step 4: Generate dashboard HTML
pyCmd = q & pythonPath & q & " " & q & strPath & "\gen_modal_dashboard.py" & q
If RunCommand("3/4 Generating dashboard HTML", pyCmd) <> 0 Then
    MsgBox "ERROR: Failed to generate dashboard HTML." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Step 5: Git add, commit, push
gitAddCmd = q & gitExe & q & " add -A"
gitCommitCmd = q & gitExe & q & " commit -m " & q & "Update dashboard " & Date & q
gitPushCmd = q & gitExe & q & " push origin main"
diffCmd = q & gitExe & q & " diff --cached --exit-code"

If RunCommand("4/4 Git add", gitAddCmd) <> 0 Then
    MsgBox "ERROR: Git add failed." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Check if there are staged changes
hasChanges = WshShell.Run("cmd /c " & diffCmd, 0, True)
' Note: git diff --cached --exit-code returns 0 if no changes, 1 if there are changes

If hasChanges = 0 Then
    WshShell.Run "cmd /c echo No changes to commit. Skipping push. >> " & q & logFilePath & q, 0, True
    MsgBox "Dashboard generated, but no changes to commit." & vbCrLf & vbCrLf & _
        "The data may be the same as yesterday." & vbCrLf & _
        "Managers can view at:" & vbCrLf & _
        "https://naichaniuniu.github.io/zhongxibu-dashboard/", _
        vbInformation, "Update Done"
    WScript.Quit 0
End If

If RunCommand("Git commit", gitCommitCmd) <> 0 Then
    MsgBox "ERROR: Git commit failed." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

If RunCommand("Git push", gitPushCmd) <> 0 Then
    MsgBox "ERROR: Git push failed." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

WshShell.Run "cmd /c echo === Update completed === >> " & q & logFilePath & q, 0, True

' Check if index.html exists
If Not fso.FileExists(strPath & "\index.html") Then
    MsgBox "ERROR: Dashboard HTML generation failed!" & vbCrLf & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Show success message
MsgBox "Dashboard updated successfully!" & vbCrLf & vbCrLf & _
    "Managers can view at:" & vbCrLf & _
    "https://naichaniuniu.github.io/zhongxibu-dashboard/" & vbCrLf & vbCrLf & _
    "Note: GitHub Pages may take 1-2 minutes to refresh.", _
    vbInformation, "Update Done"

' Open the dashboard in browser
WshShell.Run "https://naichaniuniu.github.io/zhongxibu-dashboard/"
