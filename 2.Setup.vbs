Dim domain, user
Dim shell

domain = "@hunghau.vn"
user = CreateObject("WScript.Network").UserName
Set shell = CreateObject("WScript.Shell")

shell.Run "C:\Soft\Apps\UnikeyNT.exe"

shell.Run "cmd /c echo " & user & domain & " | clip"
MsgBox "Mail Copied to Clipboard "& user & domain, 1, True

shell.Run "odopen://launch"
shell.CurrentDirectory = "C:\Soft\Apps\signatureHHH\"
shell.Run "run.exe", 1, True
shell.Run "Outlook.exe" , 1, True

shell.Run "http://datacenter.hunghau.vn/" , 1, True
shell.Run "https://wework.base.vn/" , 1, True
shell.Run "https://e.hunghau.org/" , 1, True

For Each file in CreateObject("Scripting.FileSystemObject").GetFolder("C:\Soft\Apps").Files
If LCase(Right(file.Name, 4)) = ".reg" Then
shell.Run "regedit /s """ & file.Path & """", 0, True
End If
Next

MsgBox "Done...!"

' @echo off
' set domain=@hunghau.vn
' set user=%USERNAME%
' echo %user%%domain% | clip

' start http://datacenter.hunghau.vn/
' start https://wework.base.vn/
' start https://e.hunghau.org/

' start "" "C:\Soft\Apps\UnikeyNT.exe"
' start "" "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"
' start "" "C:\Soft\Support\signatureHHH\run.exe"
' start "" "odopen://launch"

' for %%f in (C:\Soft\Apps\*.reg) do regedit /s "%%f"
' timeout 3