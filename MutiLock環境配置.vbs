DIM ws
Set ws = CreateObject("WScript.Shell")
ws.Run "MutiLock.exe Mode=SetLoginInfo iniFile=MutiLock.Ini"
Set ws=nothing