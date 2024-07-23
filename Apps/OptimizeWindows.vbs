Set Shell = CreateObject("WScript.Shell")

MsgBox ("1. Control Panel"& vbNewLine &"   1.1. Auto Play = OFF"& vbNewLine &	"   1.2. User Account = not recommended"& vbNewLine &"   1.3. Open Ultimate Power"& vbNewLine &"+ Open CMD with Administrator and run code: "& vbNewLine &"powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61"& vbNewLine &"2. Setting"& vbNewLine &"   2.1. System -> About -> Advenced system setting"& vbNewLine &". Performance"& vbNewLine &"+ Adjust for best performance" & vbNewLine &"+ Show thumbnails instead of icons"& vbNewLine &"+ Smooth edges of screen fonts"	& vbNewLine &". Startup and Recovery = 0"& vbNewLine &"   2.2. Personalization -> Transparency Effect = OFF"& vbNewLine &"   2.3. Apps delete apps"& vbNewLine &"   2.4. Privacy and Background Apps = OFF"& vbNewLine &"   2.5. Game Mode = OFF"& vbNewLine &"3. Task Manager"& vbNewLine &"Start Up = Disable Apps"& vbNewLine &"4. MS config"& vbNewLine &"Boot -> choose No GUI"& vbNewLine &"5. Restart or Shutdows")

Shell.Run "control /name Microsoft.AutoPlay", 1, True
Shell.Run "UserAccountControlSettings.exe", 1, True
Shell.Run "cmd /c powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61", 0
Shell.Run "SystemPropertiesPerformance.exe", 1, True
Shell.Run "control.exe sysdm.cpl", 1, True
Shell.Run "ms-settings:appsfeatures"
Shell.Run "taskmgr", 1, True
Shell.Run "msconfig", 1, True
MsgBox("Done") , 1, True

' Shell.Run "shutdown /r /t 10", 0
' WScript.Sleep(5000)

