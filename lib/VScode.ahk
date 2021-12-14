return
#Ifwinactive,ahk_exe Code.exe  ;;___VSCODE___
	; Mbutton::sendinput, +{F9}
	<^h::[
	<^l::]
	F17::Enter
	Numlock::ReloadScript()
	F13 & numlock::									send % tt("`n Toggle Column Selection `n ") "^+{\}"
	Media_Next::							sendinput, {altdown}{ctrldown}{lwin down}{]}{lwin up}{ctrlup}{altup} ;debug next
	Media_Play_Pause::				sendinput, {altdown}{ctrldown}{lwin down}{\}{lwin up}{ctrlup}{altup} ;debug stat
	Media_Prev::							sendinput, {altdown}{ctrldown}{lwin down}{[}{lwin up}{ctrlup}{altup} ;debug prev
	; Lbutton & F13::
	; 	send, ^c
	; 	tt(Clipboard,1000,100,-400,,160)
	; 	return
	#h::send, !{F2}
	#p::send, +!{h}
	#k::send, ^+{h}
	numpadsub::               sendinput, {shiftdown}{altdown}{ctrldown}{,}{ctrlup}{altup}{shiftup} ;open file
	numpadadd::               sendinput, {shiftdown}{altdown}{ctrldown}{.}{ctrlup}{altup}{shiftup}
	numpadmult::              send, ^{F9} ;switch workplace
	numpaddot::               numpaddot
	^numpaddot::              Sendinput,{ctrldown}{w}{ctrlup}
	; Numpadsub::               Sendinput, ^{d} ;go to Deffinition
	; Numpadadd::               Sendinput, !^{d} ;go to reference
;;	---modifiers---
	$lwin::return
		Lctrl & Appskey::         return
	RShift::										sendinput,+!{i}
	Rshift & appsKey::   			Return
	Lshift & appsKey::   			Return
	; Lwin & Appskey::          return
	Lalt & Appskey::          return
	rshift & space::          Sendinput,{shiftdown}{altdown}{ctrldown}{s}{ctrlup}{altup}{shiftup}
	<+space::                 Sendinput,{shiftdown}{altdown}{ctrldown}{e}{ctrlup}{altup}{shiftup}
;;		---F Keys---
 ; go to file
	F2::FindMatchingwindows()
	F15::F15
	; F3::Run, windowSpy.ahk,C:\Program Files\AutoHotkey\
	; F13 & s::Sendinput,{shiftdown}{altdown}{`;}{altup}{shiftup}
;send, {click}^{v} ;{F19}
	; F13 & s::						 sendinput, {right}{ctrldown}{left}{shiftdown}{right}{ctrlup}{shiftup} ;selectWord
	F13 & Lalt::sendinput, ^{click}
	F13 & tab::								Sendinput,{shiftdown}{altdown}{lwindown}{1}{lwinup}{altup}{shiftup}
	F13 & `::sendinput, {ctrldown}{F8}{ctrlup}
	F13::F13
;;		---F19 and F20---
	F20 & /:: 			 					Sendinput,{ctrldown}{f}{ctrlup}%wintitle%
	<^f19::                   Sendinput,{shiftdown}{ctrldown}{tab}{ctrlup}{shiftup}
	F20 & h::                 Sendinput,{shiftdown}{altdown}{lwindown}{left}{lwinup}{altup}{shiftup}
	F20 & k::                 Sendinput,{shiftdown}{altdown}{lwindown}{up}{lwinup}{altup}{shiftup}
	F20 & backspace::         delete
	F19 & -::                 Sendinput,{ctrldown}{-}{ctrlup}
	F19 & =::                 Sendinput,{ctrldown}{=}{ctrlup}
	F19 & y::                 Sendinput,{ctrldown}{w}{ctrlup}
	F19 & n::                 Sendinput,{shiftdown}{lwindown}{j}{lwinup}{shiftup}
	F19 & u::                 Sendinput,{shiftdown}{lwindown}{k}{lwinup}{shiftup}
	F19 & i::                 Sendinput,{F9}
	F19 & o::                 Sendinput,+!{F9}
	F19 & p::                 Sendinput,^{F9}
	F19 & Enter::							Sendinput,{ctrldown}{enter}{ctrlup}
	F19 & j::                 Sendinput,+!{j}
	F19 & k::                 Sendinput,+!{k}
	F19 & l::                 Sendinput,{shiftdown}{ctrldown}{pgdn}{ctrlup}{shiftup}
	F19 & h::                 Sendinput,{shiftdown}{ctrldown}{pgup}{ctrlup}{shiftup}
	f19 & `::                 Sendinput,~
	f19 & r::                 Sendinput,%process%
	f19 & c::                 Sendinput,%mouseposition%
	f19 & t::                 Sendinput,%wintitle%
	f19 & w::                 Sendinput,%wininfo%
	f19 & /::                 Sendinput,{shiftdown}{altdown}{lwindown}{m}{lwinup}{altup}{shiftup} ;navigate bookmarks
	; F13 & lshift::						enter
	; F13::F13

;;		---TAB---
VScodeTabsMenu(){
		try Menu,VScodeTabs, deleteAll
		Menu, VScodeTabs,add, V&quest, SelectVScodeTab
		Menu, VScodeTabs,add, VAR&BAR, SelectVScodeTab
		Menu, VScodeTabs,add, &VIM, SelectVScodeTab
		Menu, VScodeTabs,add, &Excel, SelectVScodeTab
		Menu, VScodeTabs,add, &LMS, SelectVScodeTab
		Menu, VScodeTabs,add, &temp, SelectVScodeTab
		Menu, VScodeTabs,add, &Test, SelectVScodeTab
		Menu, VScodeTabs,add, &VScode, SelectVScodeTab
		Menu, VScodeTabs,add, &OpenApp, SelectVScodeTab
		Menu, VScodeTabs,add, &Menu, SelectVScodeTab
		Menu, VScodeTabs,add, &Functions, SelectVScodeTab
		Menu, VScodeTabs,add, &clip, SelectVScodeTab
		Menu, VScodeTabs,add, &HotStrings, SelectVScodeTab
		Menu, VScodeTabs,add, &KEYS, SelectVScodeTab
		Menu, VScodeTabs,add, P&ad, SelectVScodeTab
		menu, VsCodeTabs,add
		Menu, VScodeTabs,add, &Previous, SelectVScodeTab
		Menu, VScodeTabs,add, E&xit, SelectVScodeTab
		Menu, VscodeTabs,Show
		return

SelectVScodeTab:
	If (A_ThisMenuItem Contains E&xit) {
		Menu,VScodeTabs, deleteAll
		return
		}
	If (A_ThisMenuItem Contains &Previous) {
		send, {F9}
		sleep 50
		send, {down}
	}
	else {
		ThisMenuItem:= StrReplace(A_ThisMenuItem, "&","")
		send, {F9}
		sleep 50
		send % ThisMenuItem ".ahk"
	}
	sleep 200
	send, {enter}
	Menu,VScodeTabs, deleteAll
	return
}


class Vim {

	find(){
		global
		Sendinput,{shiftdown}{altdown}{ctrldown}{f}{shiftup}{altup}{ctrlup}
		input, letter, V,{enter}{return}{up}{down}{left}{right}
		Sendinput,{enter}
		return
	}


	ChangeSelection(){
		Clipsave:=ClipboardAll
		clipboard:=
		Sendinput,^c
		clipwait, 0.25
		if errorlevel
			Sendinput,{shiftdown}{altdown}{ctrldown}{s}{ctrlup}{altup}{shiftup}
		Sendinput,{shiftdown}{ctrldown}{r}{shiftup}{ctrlup}
		sleep 200
		Sendinput,^v
		sleep 300
		Clipboard:=ClipSave
		return
	}
	Yank(Precommand:="",Cut:=""){
		global
		clipboard:=
		if Precommand
			Sendinput % PreCommand
		Send, ^{c}
		clipwait, 0.25
			; if errorlevel
				; send, {home}+{end}^{c}
		if Cut
			Sendinput, {Backspace}
		else
			send, {esc}{F3}
		sleep 20
		return
	}

	Paste(){
		global

		KeyWait, p, U T1
		if errorlevel {
			Clipboard := StrReplace(Clipboard, A_tab, "")
			Clipboard := StrReplace(Clipboard, A_space A_space, A_Space)
			Clipboard:=RegExReplace(Clipboard, "\R+\R", "`r`n")     ; remove empty lines
			Send,^v
		}
		else
			Send,{end}{enter}^{v}
		return
	}
	}
	FindMatchingwindows(){
		global
	  winGetTitle, CurrentLMSwindow, ahk_exe WFICA32.EXE
  ; Pop(SubStr(CurrentLMSwindow, 1, 20))
  	winactivate, ahk_exe Code.exe
    send, ^{f}
    sleep 200
    sendinput % SubStr(CurrentLMSwindow, 1, 25)
	}

ReloadScript(){
	global
	varbar.SaveVariables()
	Send, !s
	winSet, Transparent, 155, ahk_exe Code.exe
	sleep 100
	winSet, Transparent, off, ahk_exe Code.exe
	; try	run, VQuest.ahk, C:\Users\mmignin\Documents\VQuest
if A_DebuggerName
		ControlSend, , {Numlock}, ahk_exe Code.exe
	else
		reload
	}