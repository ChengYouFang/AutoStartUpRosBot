;RoS-BOT path
Local $rosPath = "Your RoS-Bot path"
;Change directory to BoS-BOT folder
FileChangeDir ( $rosPath )
;run RoS-BOT
Run ( "RoS-BoT.exe" )
;get RoS-BOT memory address
Local $hWnd = WinWait("[CLASS:WindowsForms10.Window.8.app.0.9585cb_r6_ad1]", "", 20)
;Sleep 3s
Sleep(3000)
;focus RoS-BOT.exe
WinActivate($hWnd)
;Sleep 1s
Sleep(1000)
;Click Start botting！
ControlClick("[CLASS:WindowsForms10.Window.8.app.0.9585cb_r6_ad1]", "", "[CLASS:WindowsForms10.BUTTON.app.0.9585cb_r6_ad1; INSTANCE:8]")
;Sleep 23s
Sleep(23000)
;Close IME
Send("{ENTER}")
Send("^{SPACE}")