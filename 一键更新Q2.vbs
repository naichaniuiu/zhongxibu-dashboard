' ============================================================
'  Zhongxibu Dashboard Q2 - One-Click Update + Auto Push to GitHub
'  Double-click this file to: Read Excel -> Generate HTML -> Push to GitHub
'  Managers just visit: https://naichaniuiu.github.io/zhongxibu-dashboard/
' ============================================================

Dim WshShell, fso, strPath, q
Dim pythonPath, gitExe, logFilePath
Dim pyCmd, gitAddCmd, gitCommitCmd, gitPushCmd, diffCmd
Dim hasChanges

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

q = Chr(34)
pythonPath = "C:\Users\wm881\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
gitExe = "C:\Users\wm881\.workbuddy\binaries\PortableGit\versions\1.2.0\cmd\git.exe"
logFilePath = strPath & "\update_log.txt"

' Step 1: Check Q2 data file
Dim dataFile
dataFile = "D:\业绩 欠款看板 Q2.xlsx"
If Not fso.FileExists(dataFile) Then
    MsgBox "ERROR: Data file not found!" & vbCrLf & vbCrLf & _
        "Please make sure the following file exists:" & vbCrLf & _
        dataFile, vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Helper function to run a command and log output
Function RunCommand(stepName, cmdArgs)
    Dim fullCmd, exitCode
    fullCmd = "cmd /c echo [" & stepName & "] >> " & q & logFilePath & q & " 2>&1 && " & cmdArgs & " >> " & q & logFilePath & q & " 2>&1"
    exitCode = WshShell.Run(fullCmd, 0, True)
    RunCommand = exitCode
End Function

' Initialize log file
WshShell.Run "cmd /c echo === Dashboard Q2 Update: %date% %time% === > " & q & logFilePath & q, 0, True

' Step 2: Generate Q2 dashboard HTML
pyCmd = q & pythonPath & q & " " & q & strPath & "\gen_q2_dashboard.py"
If RunCommand("1/3 Generating Q2 dashboard HTML", pyCmd) <> 0 Then
    MsgBox "ERROR: Failed to generate dashboard HTML." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Step 3: Copy generated HTML to index.html
Dim srcHtml, dstHtml
srcHtml = strPath & "\中西部大区26财年Q2数据看板_弹窗下钻版.html"
dstHtml = strPath & "\index.html"
If Not fso.FileExists(srcHtml) Then
    MsgBox "ERROR: Generated HTML not found!" & vbCrLf & vbCrLf & _
        "Expected at:" & vbCrLf & srcHtml, vbCritical, "Update Failed"
    WScript.Quit 1
End If
fso.CopyFile srcHtml, dstHtml, True
WshShell.Run "cmd /c echo [2/3] Copied HTML to index.html >> " & q & logFilePath & q, 0, True

' Step 4: Git add, commit, push
gitAddCmd = q & gitExe & q & " add -A"
gitCommitCmd = q & gitExe & q & " commit -m " & q & "Update Q2 dashboard " & Date & q
gitPushCmd = q & gitExe & q & " push origin main"
diffCmd = q & gitExe & q & " diff --cached --exit-code"

If RunCommand("3/3 Git add", gitAddCmd) <> 0 Then
    MsgBox "ERROR: Git add failed." & vbCrLf & "Check update_log.txt for details.", vbCritical, "Update Failed"
    WScript.Quit 1
End If

' Check if there are staged changes
hasChanges = WshShell.Run("cmd /c " & diffCmd, 0, True)

If hasChanges = 0 Then
    WshShell.Run "cmd /c echo No changes to commit. Skipping push. >> " & q & logFilePath & q, 0, True
    MsgBox "Dashboard generated, but no changes to commit." & vbCrLf & vbCrLf & _
        "The data may be the same as yesterday." & vbCrLf & _
        "Managers can view at:" & vbCrLf & _
        "https://naichaniuiu.github.io/zhongxibu-dashboard/", _
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

' Show success message
MsgBox "Q2 Dashboard updated successfully!" & vbCrLf & vbCrLf & _
    "Managers can view at:" & vbCrLf & _
    "https://naichaniuiu.github.io/zhongxibu-dashboard/" & vbCrLf & vbCrLf & _
    "Note: GitHub Pages may take 1-2 minutes to refresh.", _
    vbInformation, "Update Done"

' Open the dashboard in browser
WshShell.Run "https://naichaniuiu.github.io/zhongxibu-dashboard/"
