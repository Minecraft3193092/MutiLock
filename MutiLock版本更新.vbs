Const OverWriteExisting = True
DIM ws,fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile "LoaderUpdate.exe", "_LoaderUpdate.exe", OverWriteExisting

Set ws = CreateObject("WScript.Shell")
ws.Run "_LoaderUpdate.exe Mode=UpdateVer iniFile=MutiLock.Ini"



