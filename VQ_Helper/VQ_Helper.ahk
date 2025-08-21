	#SingleInstance,Force
	#Persistent
	#NoEnv
	Process, Priority,, High
	; SetBatchLines, -1
	; Thread, NoTimers
	#HotkeyInterval 1000
	#MaxHotkeysPerInterval 210
	#InstallKeybdHook
	#InstallMouseHook
	#ClipboardTimeout 7500
	#HotkeyModifierTimeout
	#maxthreadsperhotkey, 2
	SetTitleMatchMode, 2
	#winactivateForce
	listLines On
	FormatTime, The_Time,, hh:mm
	SetNumlockState, on
	SetscrolllockState, off
	CrLf=`r`n
	SetNumlockState Alwayson
	setcapslockstate alwaysoff
	CoordMode, mouse, Window
	CoordMode, Tooltip, screen
	; CoordMode, Tooltip, Screen
	SetWorkingDir, %A_ScriptDir%

	SetscrolllockState, alwaysoff
	AutoTrim, On
	IdleThreshold := 12 * 60 * 60 * 1000 ;4 hours
	Fileread, AllBatches, AllBatches.txt
	Fileread, AllProducts, AllProducts.txt
	iniRead, CompileTime, Settings.ini, Config, CompileTime

	ReadIniFiles()
	tooltipNumber=6
	HotkeysTip := "
(
F5    :: copyLabelCopy
^+v   :: Paste(`;)
1c`;  	:: 1 capsule
2t`;  	:: two (2) capsules
3s`;  	:: three (3) scoops
00`;  	:: #00 capsule / 0.917`" x 0.336`"
3s`;  	:: three (3) scoops
00`;  	:: #00 capsule / 0.917`" x 0.336`"
00e`;  	:: #00 elongated capsule / 0.995`" x 0.336`"
etc
)"


; !F10 :: AddsampleLog x5
; !+F10  :: AddDataFromClipboard
; ^+w   :: get_window_info
; ^+e   :: get_mouse_info
; ^F7	  :: Copy Clipped excel
; +F5   :: copyLabelCopyIng
; +F6/7 :: Copy All Products/Batches

; !F<n> :: ChangePxercision( n )
; !+v   :: Paste(A_Space)
; +F1   ::  GetAllProducts( )
; +F2   ::  GetAllBatches( )
; +^1  ::  GetAllProducts(`;)
; +!1  ::  GetAllProducts(``n)
; +^2  ::  GetAllBatches(`;)

VariableFilePath := A_ScriptDir "\ClippedExcelData.txt"

FileGetTime, InitialFileTime, %VariableFilePath%

prefix:=
	Menu,Tray,NoStandard
	Menu, Tray, Add, All Products, AllProductsMsgbox
	Menu, Tray, Add, All Batches, AllBatchesMsgbox
	Menu, Tray, add, &Final Label Copy, ShowFinalLabelCopy
	Menu, Tray, add, &Scan Label Copy, ShowScanLabelCopy
	Menu, Tray, add, &Global Vision, ShowGlobalVision
	; Menu, Tray, Add, Whole Batches, ShowWholeBatches
	; Menu, Tray, add, Show EditBox, ShowEditBox
	;  Menu, Tray, add, Add Sample Log, Add15SampleLog
	;Menu, Tray, add, Delete Prior Codes, DeletePriorCodes
	; Menu, Tray, add, Delete Whole Batches, DeleteWholeBatches
	Menu, Tray, Add,
	Menu, Tray, Add, Save List via Clipboard, SaveClipboardToList
	; Menu, Tray, Add, Get Requirements, GetRequirements
	; Menu, TestSubMenu, Add, run Script, TestCode
	; Menu, TestSubMenu, Add, Test1, TestCode1
	; Menu, TestSubMenu, Add, Test2, TestCode2
	; Menu, TestSubMenu, Add, Test3, TestCode3
	; Menu, MarkerSubMenu, Add, PasteMarker, SetPasteMarker
	; Menu, MarkerSubMenu, Add, MouseMarker, SetMouseMarker
	; Menu, MarkerSubMenu, Add, TriggerMarker, TriggerMarker
	Menu, Tray, Add, windowSpy, windowSpy
	; Menu, TestSubMenu, Add, run Script, TestCode
	; Menu, Tray, add, TestCode, :TestSubMenu ;testCode
	; Menu, Tray, add, Marker, :MarkerSubMenu
	;Menu, Tray, Add, Show Variables, ShowVariables
	;Menu, Tray, Add, ListLines, ListLines
	Menu, Tray, Add, mmignin, mmigninFolder
	; Menu, Tray, Add, Copy GUID, CopyGUID
	if (DisableAutoScroll != "YES")
			DisableAutoScroll := "NO"

; --- Build tray menu ---
		Menu, Tray, Add, Disable AutoScroll, Toggle_DisableAutoScroll
; Reflect current state with a checkmark
		if (DisableAutoScroll := "YES")
    Menu, Tray, Check, Disable AutoScroll
else
    Menu, Tray, Uncheck, Disable AutoScroll
	Menu, Tray, Add, Hotkeys, ShowHotkeys
	Menu, Tray, Add, &Reload, ReloadSub
	Menu, Tray, Add, Exitsub, Exitsub
	Menu, Tray, Add, Settings, SettingsFile
	Menu, Tray, tip, %HotkeysTip% %HotkeysTip%
	Menu, Tray, Default, &Reload


	Key:= {}
	RegexProduct:="i)(?<=[\w\d]{3})?(?P<Product>\b[defghijklmn]\d{3}\b)"
	RegexBatch:= "i)(?<!Ct#)(?P<Batch>\d{3}-\d{4})"
	RegexLot:= "i)(?P<Lot>\b\d{4}[A-Z]\d[A-Z]?-?[A-Z]?[0-9]?|\bBulk\b|\bG\d{7}[A-Z]?\b|\bVC\d{6}[ABCDEFGH]?|\bV[A-Z]\d{5}[A-Z]\d?\b|\d{5}\[A-Z]{3}\d\b)"


	RegexCoated:= "i)(\d{4}\w\d\w?.|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d\s|coated: |ct#|ct\s?|coated\s?|ct#/s)(?P<Coated>\d{3}-\d{4})"
	regexMaxLimit:="i)(<| - |NMT )(?P<maxLimit>[0-9,.]*)\s\w*"
	regexMaxLimit:="i)(<|NLT )?(?P<minLimit>(?<!NMT )[0-9.,]*)"
	regexunit:="i)(?P<unit>\w*)"
	RegexRequirements:="iO)(?<Prefix>(NLT |NMT |<))?(?<LowerLimit>([,|\d]*.?[\d]*))( - (?<UpperLimit>[,|\d]*.?[\d]*))? (?<Unit>(mg RAE|mcg RAE|mcg DFE|mg|mcg|g|IU|CFU\\g|ppm|ppb))"
	RegexSampleGUID:="i)(?P<SampleguID>\b[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}\b)"
	PreviousSampleguIDsFile:=A_ScriptDir "\PriorSampleguIDs.txt"
	; RegexCombined := "iO)(?<=[\w\d]{3})?(?P<Product>[abcdefghijklmn]\d{3}\b)|(?<!Ct#)(?P<Batch>\d{3}-\d{4}\b)|(?P<Lot>\b\d{4}\w\d\w?|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d)|(\d{4}\w\d\w?.|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d\s|coated: |ct#|ct\s?|coated\s?|ct#/s)(?P<Coated>\d{3}-\d{4})"
	RegexCombined := "i)(?<=[\w\d]{3})?(?P<Product>[defghijklm]\d{3}\b)|(?<!Ct#)(?P<Batch>\d{3}-\d{4}\b)|(?P<Lot>\b\d{4}\w\d\w?|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d)|(\d{4}\w\d\w?.|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d\s|coated: |ct#|ct\s?|coated\s?|ct#/s)(?P<Coated>\d{3}-\d{4})"



	OnClipboardChange("clipChange")
	PasteTime:=A_TickCount
	; CodeFile:= "\\10.1.2.118\users\vitaquest\mmignin\%A_ScriptDir% Code.txt"
	CodeFile:= A_ScriptDir" \Code.txt"
	OnExit("ClipBar.SaveVariables")
	SetTimer,activeCheck, %ActiveTimerCheck%
	SetTimer,FileCheck, 1000
	LMS.Orient()
	ClipBar.Show()
	sleep 200
try Menu, Tray, Icon, %AppIconPath%
GUI, ClipBar:default
	return


Toggle_DisableAutoScroll:
{
    ; flip YES/NO
    DisableAutoScroll := (DisableAutoScroll := "YES") ? "NO" : "YES"

    ; update checkmark
    Menu, Tray, ToggleCheck, Disable AutoScroll

    ; save back to INI
    IniWrite, %DisableAutoScroll%, %A_ScriptDir% \Settings.ini, Config, DisableAutoScroll

    ; optional quick feedback
    TrayTip, Settings saved, DisableAutoScroll = %DisableAutoScroll%, 2
}
return

SaveClipboardToList:

		if instr(Clipboard, "111Skin Limited",true,1,1)
	 		WorkTab.SaveCustomerList()
		else if instr(Clipboard, "5-HTP",true,1,1)
			ProductTab.SaveIngredientList()
return

;;----------------------File Check-------------------------
FileCheck:
    FileGetTime, CurrentFileTime, %VariableFilePath%
    if (CurrentFileTime <> InitialFileTime) {
        InitialFileTime := CurrentFileTime
				FileContents := FileOpen(VariableFilePath, "r")
				sleep 200
				Clip.VariableFileClip(FileContents.ReadLine())
    }

	; UpdateVariableFile()
Return
;; [[                    Active Check                    ]]


activeCheck:
ListLines, OFF


if (A_TimeIdlePhysical > IdleThreshold)
	ExitApp

	If winexist("Delete Attribute ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
		winactivate,
		sleep 200
		sendinput, {enter}
		; mousemove, 245, 137
		sleep 1000
	}
	else if winexist("Delete specification ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
		MsgBox, 4, , Do you want to continue? (Press YES or NO),5
		IfMsgBox No
			exit
		IfMsgBox timeout
		{
			winactivate
			sendinput, {n}
		return
	}
	sleep 300
	return
}

else if MouseIsOver("ClipBar"){
	ClipBar_x1:=Clipbar_x+310
	ClipBar_x2:=Clipbar_x+50
	ClipBar_x6:=Clipbar_x+265
	; CoordMode, mouse, Screen
	mousegetpos,MouseClipbarx, MouseClipbarY,,winControl
	; CoordMode, mouse, Window
	; ControlGetFocus,winControl,ClipBar
		TT("Double Click To Copy",3000,80,MouseClipbarY,2,200,"M")
	; if (winControl="Edit1"){
	; 	; try menu, AllProductsMenu, DeleteAll
	; 	; loop, parse, AllProducts, " "
	; 	; {
	; 		; Menu, AllProductsMenu, Add, %a_LoopField%, allproductsmenubutton
	; 		; if (A_LoopField = product)
	; 			; try Menu, AllProductsMenu, Check, %a_LoopField%,
	; 	; }
	; 	; try Menu,AllProductsMenu,show
	; }
	; else if (winControl="Edit2"){
	; 	TT("Double Click To Copy",1000,1,0,2,250,"M")
	; 	; try menu, AllBatchesMenu, DeleteAll
	; 	; loop, parse, AllBatches, " "
	; 	; {
	; 		; Menu, AllBatchesMenu, Add, %a_LoopField%, allbatchesmenubutton
	; 		; if (A_LoopField = batch)
	; 			; try Menu, AllBatchesMenu, Check, %a_LoopField%,
	; 	; }
	; 	; try Menu,AllBatchesMenu,show
	; 	; GetAllBatches(" ")
	; 	; sleep 1000
	; }

	; else if (winControl="Edit3"){
	; 	TT("Double Click To Copy",1000,1,0,2,250,"M")
	; 	return
	; 	; TT("Whole Batches")
	; 	;  GetAllWholeBatches("`n")
	; }
	; else if (winControl="Edit4"){
	; 	TT("Double Click To Copy",1000,1,0,2,250,"M")
	; 	; sleep 1000
	; 	; tt(sampleguid,7000,100,20,7,220,"M")
	; 	; Clipboard:=SAMPLEGUID
	; 	; ControlsetText, Edit4,%SAMPLEGUID%,ClipBar
	; 	return
	; }
	; ; else if (winControl="Edit5")
	; 	; return
	; else if (winControl="Edit6") {
	; 	TT("Double Click To Copy",1000,1,0,2,250,"M")
	; 	; BothGeneralBoxes:= "`n" GeneralBox2
	; 	; tt(BothGeneralBoxes,7000,100,2,6,220,"M")
	; }
	; else if (winControl="Edit7"){
	; 	TT("Double Click to Copy",1000,1,0,2,250,"M")
	; 	; tt(sampleguid,7000,100,20,7,220,"M")

	; }
	return
}
else if winactive("Error ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
	ControlSend,, {enter}, Error
	sleep 200
	if winExist("Register new samples") && Product{
		winactivate,
		Sendinput, {click 185, 103,2}%Product%{enter}
	}
	sleep 1000
	; RNSW:=
	return
}
else if winactive("Register new samples") && !(DisableAutoScroll="YES"){
	; if RNSW
		; return
	WinGetPos, RNSX, RNSY, RNSW, RNSH, Register new samples
	if (RNSW < 900){
		clk(231,133,,,,0) ; click register using sample templates
		winwaitclose,,,1
		return
	}
	else if (RNSW > 900){
	if BarMoved
		return
		clipbar.SaveVariables()
		; msgbox, yoyo
		GUI, ClipBar:Show, x%RNSX% y%RNSY% w%ClipBar_w% h%ClipBar_H% Noactivate, ClipBar
		;ClipBar.show(100,500)
		BarMoved:=1
		; winwaitclose,,,10
		return
	}
	else
		return
	return
}
else If winexist("Release: ahk_exe eln.exe")  && !(DisableAutoScroll="YES"){
	winactivate
	click 128,146
	return
}
; else if winactive("Book ahk_class XLMAIN ahk_exe EXCEL.exe"){
; 	WinGetPos, bWinX, bWinY, bWinW, bWinH
; 	screenwidth:=A_ScreenWidth - 50
; 	screenHeight:=A_ScreenHeight - 40
; 	if (bwinw > Screenwidth)
; 		WinRestore, ahk_class XLMAIN ahk_exe EXCEL.EXE
; 		WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, 0, 40,,%screenHeight%
; 	return
; }
else if (winactive("Result Editor ahk_exe eln.exe") && !SimpleClip && !WindowMoved  && !(DisableAutoScroll="YES")){
	WindowMoved:=1
	WinGetPos, bWinX, bWinY, bWinW, bWinH
	if (bWinH < 800) {
		; WinRestore, Result Editor ahk_exe eln.exe
		WinMove, Result Editor ahk_exe eln.exe,, , ,,853
		clk2(362,463,1,2) ;click lower limmit box
	}
	WinWaitClose,,,10
	WindowMoved:=
	return
}
;;-----EXCEL Handles
; else if (winexist("ahk_class XLMAIN ahk_exe EXCEL.EXE") && !ExcelConnected){
; 	try XL := ComObjActive("Excel.Application") ;handle to running application
; Catch {
; 					MsgBox % "no existing Excl ojbect:  Need to create one"
; 			XL := ComObjCreate("Excel.Application")
; 			XL.Visible := 1 ;1=Visible/Default 0=hidden
; 			}
; 			; XL.Visible := 1 ;1=Visible/Default 0=hidden
; 			ExcelConnected:=1
; 	Return
; }




else If winactive("Edit Formulation ahk_exe eln.exe") && !WindowMoved && !(DisableAutoScroll="YES"){
	clk(458,477,,2)
	sleep 400
	WindowMoved:=1
	WinWaitClose,,,10
	WindowMoved:=
	return
}
else If winactive("Edit Formulation ahk_exe eln.exe") && !WindowMoved && !(DisableAutoScroll="YES"){
	; sleep 200
	; ifwinactive,
	; clk2(469,533,0,2)
	clk(469,533,,2)
	clk(259,246,0,0)
	; clk2(259,246,0,0)
	sleep 400
	WindowMoved:=1
	WinWaitClose
	WindowMoved:=
	return
}
else If winactive("Test Definition Editor ahk_exe eln.exe") && !WindowMoved && !(DisableAutoScroll="YES"){
	; sleep 200
	clk(469,530,,1)
	; ifwinactive,
	; clk2(469,530,0)
	; clk2(259,487,0,0)
	sleep 50
	WindowMoved:=1
	; WinWaitClose
	; WindowMoved:=
	return
}
else If winactive("Reason For Change") && !WindowMoved && !(DisableAutoScroll="YES"){
	send, {tab 2}
	sleep 400
	WindowMoved:=1
	WinWaitClose
	WindowMoved:=
	return
}
else If winactive("Select methods tests") && !WindowMoved && !(DisableAutoScroll="YES"){
	sleep 600
	; clk2(233, 67,1,2)
	Click, 233, 67, 2 ; select search bar
	WindowMoved:=1
	; WinWaitClose
	; WindowMoved:=
	return
}
; else If winactive("Edit test (Field") && !WindowMoved{
; 	; sleep 200
; 	; ifwinactive,
; 	; clk2(471, 536,0,2)
; 	; clk2(238, 545,0,0)
; 	WindowMoved:=1
; 	WinWaitClose
; 	WindowMoved:=
; 	return
; }
else If winactive("Delete Test ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
	winactivate,
	sleep 100

	sendinput, {enter}
	; mousemove, 222, 138
	; click
	; sleep 1000
	return
}
else If winactive("Delete results ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
	winactivate,
	sleep 100
	sendinput, {enter}
	; mousemove, 222, 138
	; click
	sleep 1000
	return
}
else If winactive("Delete ingredients ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
	winactivate,
	sleep 200
	sendinput, {enter}
	; mousemove, 222, 138
	sleep 1000
	return
}
else if winactive("Lock specification ahk_exe eln.exe") && !(DisableAutoScroll="YES")
{
	sendinput, {n}
	sleep 9000
	return
}
else if winActive("Information ahk_exe eln.exe") && !(DisableAutoScroll="YES"){
	winactivate,
	send, {enter}
}
else if Winactive("NuGenesis LMS"){
	; LMS.DetectTab()
	if BarMoved
	{
		clipbar.SaveVariables()
		Clipbar.Show()
		BarMoved:=
	}
	windowmoved:=
}
; listlines, On
	; #maxthreadsperhotkey, 2
return




Block_Input:
; Thread, Priority, 1
	sleep 100
		NAdd:=
		nsub:=
		Mousemoved:=
	BlockInput, off
	settimer, Block_Input, Off
Return





#include NuGenesis.ahk
	#Include KEYS.ahk
	#include CLIPBAR.ahk
		#include Markers.ahk


copyLabelCopyDoc(SaveText:="",showtooltip:=""){
	Global Product,RegexIngredients

	firstLetter:=SubStr(Product,1,1)
FilePattern := "\\netapp\Label Copy Final\" firstLetter "000-" firstLetter "999\*" product "*.docx"
Loop, %FilePattern%, 1, 0
		oW:=ComObjGet(A_LoopFileLongPath)
		; sleep 1000
		sleep 200
		oW.Range.FormattedText.Copy
		clipwait,4,0
		if errorlevel
			msgbox, didnt  find labelcopy
	LabelCopyText:=Clipboard

	Ingredients:= RegexMatch(LabelCopyText, RegexIngredients,ri)

		 ;Clipboard:=LabelCopyText

	If showTooltip
		tt(LabelCopyText,1000)
	If SaveText
		{
		Try FileDelete, %A_ScriptDir% \LabelCopies\%Product%.txt
		FileAppend,  %labelcopytext%, %A_ScriptDir% \LabelCopies\%Product%.txt
		}
	; Clipboard:=listofIngredients
		; MsgBox % riIngredients
	Return LabelCopyText
}





GetRequirements(){
	Global
	RegexRequirements:="iO)(?<Prefix>(NLT |NMT |<|>))?(?<LowerLimit>([,|\d]*.?(?<PercisionAmount>[\d]*)))( - (?<UpperLimit>[,|\d]*.?[\d]*))? (?<Unit>(mg RAE|mcg RAE|mcg DFE|mg|mcg|g|IU|CFU\\g|ppm|ppb))"
	; Parse:="15.0 - 24.8 mg"
	; Parse:="NLT 50.0 mcg"
	; Parse:="NMT 50.04 mcg"
	; Parse:="1,000.00 - 1,650.00 mcg RAE"
	; Parse:="<2.0 mg"
	Parse:=Clipboard

If (RegexMatch(Parse, RegexRequirements, subpat)){
	; if (RegexMatch(Parse, RegexRequirements, r)){
					Prefix:=Subpat["Prefix"]
					MinLimit:=Subpat["LowerLimit"]
					MaxLimit:=Subpat["UpperLimit"]
					Units:=Subpat["Unit"]
					Percision:=StrLen(Trim(Subpat["PercisionAmount"]))
					; Prefix:=rPrefix
					; LowerLimit:=rLowerLimit
					; UpperLimit:=rUpperLimit
					if (Prefix = "NMT ")
					{
						Maxlimit:=MinLimit
						MinLimit:=""
					}
					IF (MinLimit && Maxlimit)
						divider:=" - "
					else
						Divider:= " "
					Final_Requirements:=Prefix " " MinLimit Divider MaxLimit " " Units
					; Unit:=rUnit
					Clipped_Requirement:=Trim(StrReplace(StrReplace(StrReplace(Final_Requirements, "< ","<"),"  "," "),"""",""))
					sleep 300
		; msgbox, prefix: %Prefix% `nlowerlimit: %LowerLimit% `nupperlimit: %UpperLimit% `nunit: %Unit%
		msgbox, prefix: %Prefix% `nMinlimit: %MinLimit% `nMaxlimit: %MaxLimit% `nunits: %Units% `n %rprefix% `t %rUpperLimit% `n %Clipped_Requirement% `t %SubPat% `n %Percision%
}
	; }
	; else
	; msgbox, %Parse%
}


TestCode3:
copyLabelCopyDoc(1,1)
Return
TestCode2:
GetRequirements()
Return
TestCode1:
copyLabelCopyDocRegex(1)
return

