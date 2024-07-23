msgbox "Developed: Nguyen Lam Truong"& vbNewLine &"Phone: (+84) 85 3714 852", vblInformation, "Contact"

Set objClipboard = CreateObject("WScript.Shell")
objClipboard.Run "cmd /c echo (+84) 85 3714 852|clip", 2, True