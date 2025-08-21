#SingleInstance,Force
#Persistent
#NoEnv
; SetBatchLines, -1
#HotkeyInterval 1000
#MaxHotkeysPerInterval 210
#ClipboardTimeout 7500
#HotkeyModifierTimeout
#maxthreadsperhotkey, 1
SetTitleMatchMode, 2
#winactivateForce
; SetControlDelay -1
CoordMode, ToolTip, Window
SetNumlockState, on
SetscrolllockState, off
CrLf=`r`n
SetNumlockState Alwayson
setcapslockstate alwaysoff
SetWorkingDir, %A_ScriptDir%
SetscrolllockState, alwaysoff
AutoTrim, On
SetKeyDelay, -1 , 0
RegexSampleGUID := "i)(?P<SampleGUID>\b[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}\b)"
RegexSaomething := "i)(?P<SampleGUID>\b23[0-1][0-9]{5}|S[0-9]{8}-[0-9]{3}\b)"
RegexBatch := "i)(?<!Ct#)(?P<Batch>\d{3}-\d{4}\b)"
QuickPaste := "20230801000000000"
EnterEachLine := A_ScriptDir "RawFiles\EnterEachLine.ahk"
; SettingsFile := "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\RawFiles\Settings.ini"
	iniRead, ResizeExcel, Settings.ini, Config, ResizeExcel
	stringupper,ResizeExcel,ResizeExcel

full_command_line := DllCall("GetCommandLine", "str")

Menu, Tray, add, EnterEachLine, ShowEnterEachLine
Menu, Tray, add, Edit TextFile, EditEnterEachLine
Menu, Tray, Default, EnterEachLine

if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
	try
	{
		if A_IsCompiled
			Run *RunAs "%A_ScriptFullPath%" /restart
		else
			Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
	}
	ExitApp
}


OnClipboardChange("clipChange")
SetTimer,activeCheck, 300
return
EditEnterEachLine:
	Run , Edit A_ScriptDir "\TEXTtoENTER.txt"
return


ShowEnterEachLine:
!^v::Run, %EnterEachLine%
return

activeCheck:
	ListLines, OFF

	If winactive("Counter usage detected ahk_exe EXCEL.exe") {
		winactivate,
		sleep 100

		sendinput, {enter}
		; mousemove, 222, 138
		; click
		; sleep 1000
		return
	}
	else If winactive("Reason For Change ahk_exe eln.exe") {
		sendinput, {tab 2}Testing{tab}{enter}
		return
	}
	else If winactive("Reason For Change ahk_exe EXCEL.EXE") {
		sendinput, {tab 2}Testing{Click 173 334}
		return
	}
	else If winactive("Reason for Change ahk_exe EXCEL.EXE") {
		sendinput, {tab 2}Testing{Click 173 334}
		return
	}
	else If winactive("Smart Builder ahk_exe EXCEL.EXE ahk_pid 9036") {
		sendinput, {right}
		winwaitclose
		return
	}
	else If winexist("Smart Builder ahk_exe EXCEL.EXE", "Do you want to delete the fetched data?") {
		winactivate,
		sendinput, {n}
		winwaitclose
		return
	}
	else If winactive("Smart Builder ahk_exe EXCEL.exe", "Are you sure you want to delete the selection?") {
		winactivate,
		sleep 100

		sendinput, {enter}
		; mousemove, 222, 138
		; click
		; sleep 1000
		return
	}
	else if winactive("Book ahk_class XLMAIN ahk_exe EXCEL.exe") && !(ResizeExcel = "NO"){
		WinGetPos, bWinX, bWinY, bWinW, bWinH
		screenwidth := A_ScreenWidth - 10
		screenHeight := A_ScreenHeight - 40
		if (bwinw > Screenwidth)
			WinRestore, ahk_class XLMAIN ahk_exe EXCEL.EXE
		WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, 0, 40,,%screenHeight%
		return
	}
	else
		return


Ralt::return

^+c::copyTOClipboard2()

clipChange(type) {
	global
	rSampleGUID :=
	rbatch :=
	if SampleGUID
		PreviousSampleGUID := SampleGUID
	SampleGUID :=
	sleep 75
	; If winactive("NuGenesis LMS"){
	clipboardParse := Clipboard
	clipwait,1
	SampleGUID := TRIM(RegExMatch(ClipboardParse, RegexSampleGUID, r))? rSampleGUID : SampleGUID
	; Batch:=TRIM(RegExMatch(ClipboardParse, RegexBatch, r))? rBatch : Batch
	tooltip,  %rSampleGUID% `t %rSampleGUID%, 600,0
	sleep 900
	tooltip,
	if (SampleGUID=PreviousSampleGUID) && (SampleGUID)
		FileAppend, %SampleGUID%`n, %A_ScriptDir% PriorSampleGUIDs.txt
	; }
	; If winactive("TIBCO Jaspersoft") || winactive("Dataset and Query Dialog"){
	; 	copyTOClipboard2()
	; 	return
	; }
	else if InStr(Clipboard, "<<QuIT>>", true, 1, 1) {
		Clipboard :=
		exitApp
		Return
	}
	return
}



; ~RWin::Send {Blind}{vkFF}
; ~LWin::Send {Blind}{vkFF}



#ifwinactive, TIBCO Jaspersoft
+enter::
	ControlClick,ToolbarWindow3254, TIBCO Jaspersoft,,,, x10 y10
	sleep 300
	; send, {down}{enter}
return
; +enter::ControlClick ,Button47, TIBCO Jaspersoft
; +Enter::click 648, 241
F8::click 648, 241
F9::controlclick, Button47, TIBCO Jaspersoft
F7::PriorSampleGUIDsMenu(1)
^c::copyTOClipboard2()

#ifwinactive, Parameter: SAMPLEGUID
F9::
+^v::
	controlclick, Button1, Parameter: SAMPLEGUID
	; click 372,79
	sleep 500
	; winwaitactive,
	; controlclick, Edit1, Parameter: SAMPLEGUID
	send, {tab}
	sleep 100
	send, ^{v}
	sleep 200
	send, {tab}{enter}
	sleep 200
	click 256, 473
	winwaitactive, TIBCO Jaspersoft,,5
	sleep 200
	sleep 200
	ifwinactive, TIBCO Jaspersoft
	{
		; ControlClick,ToolbarWindow3254, TIBCO Jaspersoft,,,, x80 y10
		sleep 300
		ControlClick,ToolbarWindow3254, TIBCO Jaspersoft,,,, x10 y10
		; sendinput, {down}{enter}
	}
	; controlclick, Buttton6, Parameter: SAMPLEGUID
	; click 253, 472
return

#Ifwinactive, ahk_exe eln.exe
F1::sendinput, %batch%
F2::Sendinput, %SampleGUID%
; F3::Sendinput, %SampleGUID%

+F10::sendinput, +{V}{Tab 4}Form Contro{Tab 2}Y

#IfWinactive, Dataset and Query Dialog
^x::
^c::copyTOClipboard2()

copyTOTextfile() {
	Send, {ctrldown}{c}{ctrlup}
	sleep 100
	Ifwinnotexist, Untitled - Notepad
		Run, notepad.exe
	else
		Winactivate, Untitled - Notepad
	WinWaitActive, Untitled - Notepad
	Sendinput, ^{a}^{v}
	Sleep, 200
	Send, ^{a}^{c}
	Winminimize, Untitled - Notepad
	Sleep, 100
	return
	Return
}

copyTOClipboard2() {
	; Copy selected text
	Send, ^c
	Sleep, 100
	; tooltip,%Clipboard%
	; sleep 100
	; Tooltip
	; Store the copied text into a variable
	ClipWait, 1 ; Wait for the clipboard to have data (1 second timeout)
	if (ErrorLevel) {
		MsgBox, Error: No text was copied to the clipboard!
		return
	}
	copiedText := Clipboard

	; Optionally process the text if needed (e.g., trim, modify)
	; For now, simply re-copy the text to the clipboard
	Clipboard := copiedText

	IfWinNotExist, Untitled - Notepad
	{
		Run, notepad.exe
		; Sleep, 300 ; Allow Notepad to launch if necessary
	WinWaitActive, Untitled - Notepad,,1
		WinSet, Transparent, 220, Untitled - Notepad
		WinMove, Untitled - Notepad, , A_ScreenWidth * 0.85, A_ScreenHeight * 0.75, A_ScreenWidth * 0.15, A_ScreenHeight * 0.25
Return
	}

	else
		WinActivate, Untitled - Notepad
	; sleep 250
	; Paste text into Notepad in the background
	WinWaitActive, Untitled - Notepad,,0.5
	Send, ^{a}^{v}
	; Confirm that the clipboard has been updated
	; ClipWait, 1
	Winminimize, Untitled - Notepad
	if (ErrorLevel) {
		MsgBox, Error: Failed to update clipboard!
		return
	}

	return
}

; #ifwinactive, Microsoft Visual Basic for Applications
; F10::F5
; 		return

; #ifwinactive, ahk_exe EXCEL.EXE
; ;	F2::Sendinput, %SampleGUID%{enter}
; ;F3::Sendinput, %SampleGUID%{enter}
; ;F3::Sendinput, %SampleGUID%{enter}
; F10::
; 	if !winactive("Microsoft Visual Basic for Applications")
; 		winactivate, Microsoft Visual Basic for Applications
; 	sendinput, {F5}
; 	winactivate, ahk_class XLMAIN
; return
; F7::
; 	WinGetPos, WinX, WinY, WinW, WinH
; 	targetWidth := A_ScreenWidth / 3 ; Calculate one third of the screen width
; 	targetX := A_ScreenWidth - targetWidth ; Calculate the X position to start the window
; 	WinRestore, ahk_class XLMAIN ahk_exe EXCEL.EXE  ; Restore the window to normal state if it is maximized
; 	WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, targetX, WinY, targetWidth, WinH
; return
; F6::
; 	WinGetPos, WinX, WinY, WinW, WinH
; 	targetWidth := A_ScreenWidth / 2 ; Calculate one half of the screen width
; 	;targetX := A_ScreenWidth - targetWidth  ; Calculate the X position to start the window
; 	WinRestore, ahk_class XLMAIN ahk_exe EXCEL.EXE  ; Restore the window to normal state if it is maximized
; 	WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, 0, WinY, targetWidth, WinH
; return
; F8::
; 	CoordMode, Mouse, Screen
; 	WinGetPos, WinX, WinY, WinW, WinH
; 	MouseGetPos, rmx, rmY,
; 	if rmx < 400
; 		rmx := 400
; 	if rmy< 600
; 		rmy := 600
; 	WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, %Winx%, %winY%, %rMx%, %rmy%
; 	sleep 400
; 	CoordMode, Mouse, Window
; return

; #ifwinactive

; #v::Send, ^{v}
; !v::Send, ^v
; F1::Sendinput, %Batch%
; F2::Sendinput, %SampleGUID%
; F3::Sendinput, %SampleGUID%
; +F1::
; Send, ^c
; sleep 300
; Batch:=Clipboard
; tooltip,%Clipboard%
; sleep 500
; Tooltip
; return
; +F2::
; Send, ^c
; sleep 300
; SampleGUID:=Clipboard
; tooltip,%Clipboard%
; sleep 500
; Tooltip
; return
; +F3::
; Send, ^c
; sleep 300
; SampleGUID:=Clipboard
; tooltip,%Clipboard%
; sleep 500
; Tooltip
; return





PriorSampleGUIDsMenu(ShowMenu := "") {
	global
	; PreviousSampleGUIDs:=[]
	try Menu, SampleGUIDMenu, DeleteAll
	FileRead, PriorSampleGUIDs, PriorSampleGUIDs.txt
	PriorSampleGUIDs := Trim(StrReplace(PriorSampleGUIDs, "`n`n", ""))
	Loop, parse, PriorSampleGUIDs, "`n"
	{
		Menu, SampleGUIDMenu, Add, %a_LoopField%, PriorSampleGUIDsMenubutton,
		; PreviousSampleGUIDs.Push(A_loopfield)
	}
	Menu, SampleGUIDMenu, Add,
	Menu, SampleGUIDMenu, Add, Edit List, EditListButton

	if ShowMenu
		Try Menu,SampleGUIDMenu,show
	return


	PriorSampleGUIDsMenubutton:
	winactivate, ahk_exe Jaspersoft Studio.exe
	Clipboard := A_ThisMenuItem
	sleep 200
	if (winactive("TIBCO Jaspersoft") || winactive("Parameter: SAMPLEGUID")) {
		ifwinactive, TIBCO Jaspersoft
		{
			ControlClick,ToolbarWindow3254, TIBCO Jaspersoft,,,, x75 y10
			sleep 200
			controlclick, Button47, TIBCO Jaspersoft
			sleep 200
			winwaitactive, Parameter: SAMPLEGUID,,2
			sleep 200
		}
		sleep 200
		ifwinactive, Parameter: SAMPLEGUID
		{
			controlclick, Button1, Parameter: SAMPLEGUID
			sleep 200
			send, {tab}
			sleep 100
			send, ^v
			sleep 300
			send, {tab}{enter}
			click 256, 473
			sleep 100
			winwaitactive, TIBCO Jaspersoft,,2
			sleep 200
			ControlClick,ToolbarWindow3254, TIBCO Jaspersoft,,,, x10 y10
			return
			; controlclick, Buttton6, Parameter: SAMPLEGUID
		}
	}
	else
		sendinput % Trim(A_ThisMenuItem)
	return
	; sendinput % A_ThisMenuItem
	; Clipboard:= Trim(A_ThisMenuItem)
	try Menu, SampleGUIDMenu, DeleteAll
	return

	EditListButton:
	Run, Edit A_ScriptDir "\PriorSampleGUIDs.txt"
	return


}

