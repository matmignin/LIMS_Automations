

#IfWinActive,
  ; #If (A_PriorHotKey = "numpaddiv" && A_TimeSincePriorHotkey < 400)
  ;   numpaddiv::CloseWindow() ;4 finger swipe down
  #If (A_PriorHotKey = "Numlock" || A_PriorHotkey = "Mbutton" && A_TimeSincePriorHotkey < 500) ;4 finger swipe down
  F6::send, {F6} 
  F7::send, {F7} 
  F8::send, {F8} 
  F9::send, {F9} 
  numpadsub::send, {NumpadSub}
  numpadadd::send, {NumpadAdd}
  numpadmult::send, {numpadmult} 
  numpaddiv::send, {numpadDiv}  ;4down clear filter
  lbutton::send, {Lbutton}
  Numlock::Send, {Numlock}
  mbutton::send, {mbutton}
  wheelup::send, {WheelUp}
  ; rbutton::send % "{shiftDown}{click}{shiftup}" ;Mouse_RbuttonUP()
  #If
 TrouchPad:
#If getkeystate("lbutton","p")
  space::click
  F19::send, {altdown}{ctrldown}{4}{ctrlup}{altup}
  .::VS_Code_WindowInfo()
  ; Lwin::^x
  v::send, {shiftdown}{altdown}{ctrldown}{v}{ctrlup}{altup}{shiftup}
  ;  F20::Wheel_Paste()
  /::varbar.reset()
  e::send,{LWinDown}{e}{lwinup}
  o::OpenApp.Outlook()
  d::LMS.Orient()
  ;  f::OpenApp.firefox()
  w::OpenApp.Workbook()
#If
#IfWinActive,
;  [TrackPad]
numpadsub::#left
numpadadd::#right
numpadMult::#up
  Numlock::4_tap() ;send, {altDown}{lwindown}{Ctrldown}{o}{CtrlUp}{lwinup}{altup}
numpaddiv::#down
return
; #If (A_PriorHotKey = "NumpadDiv" AND A_TimeSincePriorHotkey < 450) ;4 finger swipe down
; ;   F6::send, #{left} ;4left
; ;   F7::send, #{right}-
; ;   F8::send, ^{home}
; ;   F9::send, {altDown}{Ctrldown}{tab}{CtrlUp}{altup}
; ;   numpadsub::send, +#{left} ;4left
; ;   numpadadd::send, +#{right}
; ;   numpadmult::send, {ShiftDown}{altDown}{up}{altup}{ShiftUp} ;4up
; ;   ; numpadmult::send, +#{up} ;4up
; ;   lbutton::send, ^{click}
; ;   ; Numlock::Send, +#{right}
; ;   mbutton::send, {shiftDown}{click}{shiftup}
; ;   ; rbutton::send % "{shiftDown}{click}{shiftup}" ;Mouse_RbuttonUP()
; ;   #If
;   numpaddiv::
;   TT("Down")
; ;   sleep 450
;   if (A_TimeSincePriorHotkey > 450)
;     send, {altDown}{lwindown}{Shiftdown}{down}{ShiftUp}{lwinup}{altup}
;     sleep 200
;   return



; #If (A_PriorHotKey = "Numpadmult" AND A_TimeSincePriorHotkey < 450) ;4 finger swipe down
;   F6::send, #{left} ;4left
;   F7::send, #{right}
;   F8::send, ^{home}
;   F9::send, {altDown}{Ctrldown}{tab}{CtrlUp}{altup}
;   numpadsub::send, +#{left} ;4left
;   numpadadd::send, +#{right}
;   ; numpadmult::
;   numpaddiv::CloseWindow() ;4down clear filter
;   lbutton::send, ^{click}
;   ; Numlock::Send, +#{right}
;   mbutton::send, {shiftDown}{click}{shiftup}
;   ; rbutton::send % "{shiftDown}{click}{shiftup}" ;Mouse_RbuttonUP()
;   #If
;   numpadmult::
;   TT("up")
;   sleep 450
;   if (A_TimeSincePriorHotkey > 450)
;     send, {altDown}{lwindown}{Shiftdown}{up}{ShiftUp}{lwinup}{altup}
;     sleep 200
;   return
  
  
  
  
  
4_Tap(){
     TT(A_ThisHotkey,450,,,2)
     Input, Akey, T0.45 ,{NumpadAdd}{Numpadsub}{Numpadmult}{numpad0}{numpad5}{numpad4}{numpad3}{Numpad2}{Numpad1}{numpad9}{numpad8}{numpad7}{numpad6}{Numpaddiv}{space}{Mbutton}{F6}{F7}{F8}{F9}{F19}{F20}{Lbutton}{Mbutton}{numlock}
      If InStr(ErrorLevel, "EndKey:")
      {
        If InStr(ErrorLevel, "NumpadAdd")
          send, {altdown}{ctrldown}{right}{ctrlup}{altup}
        If InStr(ErrorLevel, "NumpadSub")
          send, {altdown}{ctrldown}{left}{ctrlup}{altup}
        If InStr(ErrorLevel, "NumpadMult")
          send, {altdown}{ctrldown}{up}{ctrlup}{altup}
        If InStr(ErrorLevel, "NumpadDiv")
          CloseWindow()
        If InStr(ErrorLevel, "F6")
          send, {altDown}{lwindown}{o}{lwinup}{altup}
        If InStr(ErrorLevel, "F7")
          send, {altDown}{lwindown}{F7}{lwinup}{altup}
        If InStr(ErrorLevel, "F9")
          send, {ctrldown}{w}{ctrlup}
        If InStr(ErrorLevel, "F8")
          send, {ctrldown}{`}{ctrlup}
        exit
      }
      If ErrorLevel = TimeOut
      {
          send, {altDown}{lwindown}{Ctrldown}{o}{CtrlUp}{lwinup}{altup}
          return
      }
      return
}
  
  
  

  
  
#IfWinActive, ahk_exe Code.exe

  ; $numpadMult:: send, {ShiftDown}{altDown}{up}{altup}{ShiftUp}  ;4up
 numpaddiv::send, {altDown}{lwindown}{Shiftdown}{down}{altup}{ShiftUp}{lwinup}
 numpadmult::send, {altDown}{lwindown}{Shiftdown}{UP}{ShiftUp}{altup}{lwinup}




  Mbutton:: 
     TT(A_ThisHotkey,600,,,2)
     Input, Akey, T0.6 ,{NumpadAdd}{Numpadsub}{Numpadmult}{numpad0}{numpad5}{numpad4}{numpad3}{Numpad2}{Numpad1}{numpad9}{numpad8}{numpad7}{numpad6}{Numpaddiv}{space}{Mbutton}{F6}{F7}{F8}{F9}{F19}{F20}{Lbutton}{numlock}
      If InStr(ErrorLevel, "EndKey:")
      {
        If InStr(ErrorLevel, "NumpadAdd")
          send, {altdown}{ctrldown}{right}{ctrlup}{altup}
        If InStr(ErrorLevel, "NumpadSub")
          send, {shiftdown}{altdown}{F3}{altup}{shiftup}
        If InStr(ErrorLevel, "NumpadMult")
          send, {altdown}{ctrldown}{up}{ctrlup}{altup}
        If InStr(ErrorLevel, "Numpaddiv") {
          send, {shiftdown}{altdown}{ctrldown}{v}{ctrlup}{altup}{shiftup}
          TT("column select",3000)
        }
        If InStr(ErrorLevel, "F9")
          send, {pgup}
        If InStr(ErrorLevel, "F6")
          send, {altDown}{lwindown}{o}{lwinup}{altup}
        If InStr(ErrorLevel, "F7")
          send, {ctrldown}{f}{ctrlup}
        If InStr(ErrorLevel, "F8")
        send, {pgdn}
          
        exit
      }
      If ErrorLevel = TimeOut
      {
          ; TT("Previous ",800)
          send, {altdown}{down}{altup}
          exit
      }
      return
  ; Mbutton::sendinput, {shiftdown}{altdown}{ctrldown}{s}{ctrlup}{altup}{shiftup} ;search
  F6::send, {altDown}{left}{altup}
  F7::send, {altDown}{right}{altup}
  F8::ToggleDefinition()
  F9::send, {ShiftDown}{altDown}{up}{altup}{ShiftUp}
  numpadadd::send, {altDown}{lwindown}{Shiftdown}{right}{ShiftUp}{lwinup}{altup}
  numpadsub::send, {altDown}{lwindown}{Shiftdown}{left}{ShiftUp}{lwinup}{altup}
  ; numpaddiv::send, {altDown}{lwindown}{Shiftdown}{down}{ShiftUp}{lwinup}{altup}
  
  

; $numlock::send, {altDown}{lwindown}{Ctrldown}{o}{CtrlUp}{lwinup}{altup}
  

#ifwinactive

 ~lbutton::return
 
 
CloseWindow(){
	global
	if WinActive("Inbox - mmignin@vitaquest.com - Outlook") ; || winactive("ahk_exe OUTLOOK.EXE")
		Return
	else if winexist("PDF Preview - \\Remote"){
		winactivate
		clk(944, 11)
		sleep 400
	}
	; else if winactive("Composition - \\Remote"){
	; 	clk(841, 895)
	; 	sleep 50
	; 	clk(946, 896)
	; 	sleep 400
	; }
	else If winactive("NuGenesis LMS - \\Remote")
		LMS.FilterClear()
	; else If WinActive("ahk_exe WFICA32.EXE")
	; {
	; 	send,{esc}
	; 	sleep 400
	; 	return
	; }
	else if WinActive("ahk_exe firefox.exe") || winactive("ahk_exe msedge.exe") || winactive("ahk_exe Code.exe") {
		send, ^{w}
		sleep 400
		return
	}
	else if WinActive("ahk_exe explorer.exe") || winactive("ahk_exe OUTLOOK.EXE") || winactive("OneNote for Windows 10") {
		sendinput, !{F4}
		sleep 400
		return
	}
	; else if winactive("ahk_exe Code.exe")
	; {
	; 	sendinput, ^{F4}
	; 	sleep 400
	; 	return
	; }
	else if winactive("Settings ahk_class ApplicationFrameWindow"){
		winclose
		sleep 400
		return
	}
	; else if winactive("ahk_exe mstsc.exe") 
	else if winactive("ahk_exe EXCEL.EXE"){
		Send, {ctrldown}{down}{ctrlup}
		sleep 400
		return
	}
	else if winactive("Adobe Acrobat Reader"){
		send, ^q
		sleep 400
	}
	else
		PostMessage, 0x112, 0xF060,,, A     ; ...so close window     
	return
}
