;RoS-BOT path
Local $rosPath = "Your RoS-Bot path"
;Change directory to BoS-BOT folder
FileChangeDir ( $rosPath )
;run RoS-BOT
Run ( "RoS-BoT.exe" )
;get RoS-BOT memory address
Local $hWnd = WinWait("[CLASS:WindowsForms10.Window.8.app.0.141b42a_r6_ad1]", "", 15)
;Sleep 3s
Sleep(5000)
;focus RoS-BOT.exe
WinActivate($hWnd)
;Sleep 1s
Sleep(1000)
;Click Start botting！
ControlClick("[CLASS:WindowsForms10.Window.8.app.0.141b42a_r6_ad1]", "", "[CLASS:WindowsForms10.BUTTON.app.0.141b42a_r6_ad1; INSTANCE:8]")
