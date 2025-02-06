#ifwinactive,
return
~RWin::Send {Blind}{vkFF}
~LWin::Send {Blind}{vkFF}
; +F4::ExitApp
^+esc::ExitApp
^esc::
	Critical, on
	Thread, Priority, 1
Reload
return
^+`::
	sleep 800
	Run, "U:\VQ_Helper\RawFiles\COMPILE.exe"
	exitapp
	Return
$enter::Enter

	^+w::gosub, get_window_info
	^+e::gosub, get_mouse_info


LabelCopyDoc:
msgbox, % copyLabelCopyDoc(1)
return
copyLabelCopyDocRegex:
Msgbox % copyLabelCopyDocRegex(1)
Return




GetRequirements:
GetRequirements()
return



F7::
	Winactivate, NuGenesis LMS
	preclk()
	LMS.SearchBar(Batch,"{enter}",0)
	postclk()
	Return
F6::
	Winactivate, NuGenesis LMS
	preclk()
	LMS.SearchBar(Product,"{enter}",0)
	postclk()
	Return
F8::
	Winactivate, NuGenesis LMS
	preclk()
	LMS.SearchBar("",,"False")
	postclk()
	Return



;-------------------------------------------------------------------------------
;;[[*** CLIPBAR KEYBINDINGS                                               ]]
;-------------------------------------------------------------------------------
#If MouseIsOver("Methods List ahk_exe VQ_Helper.exe")
Lbutton::send, {ctrldown}{Lbutton}{ctrlup}

#If MouseIsOver("ClipBar ahk_exe VQ_Helper.exe")
+Mbutton::
		ControlGetFocus,winControl,ClipBar
		if (winControl="Edit2"){
			GetAllWholeBatches()
		}
		if (winControl="Edit3"){
			MenuCodeSelect()
		}
; 		else
; send, {ctrldown}{Lbutton}{ctrlup}
; return

RBUTTON::
ControlGetFocus,winControl,ClipBar
		if (winControl="Edit6"){
			SimpleClip:=1
			Clipboard:= GetAllProducts(" ")
			Clipbar.Flash()
			SimpleClip:=
			return
		}
		if (winControl="Edit7"){
			SimpleClip:=1
			Clipboard:=GetAllBatches(" ")
			SimpleClip:=
			Clipbar.Flash()
			 return
		}
		else
			LMS.Menu()
		return
+RBUTTON::
ControlGetFocus,winControl,ClipBar
		if (winControl="Edit6"){
			SimpleClip:=1
			Clipboard:= GetAllProducts(";")
			SimpleClip:=
			Clipbar.Flash()
			return
		}
		if (winControl="Edit7"){
			SimpleClip:=1
			Clipboard:=GetAllBatches(";")
			SimpleClip:=
			Clipbar.Flash()
			 return
		}
		else
			LMS.Menu()
		return
!RBUTTON::
ControlGetFocus,winControl,ClipBar
		if (winControl="Edit6"){
			SimpleClip:=1
			Clipboard:= GetAllProducts("`r`n")
			SimpleClip:=
			Clipbar.Flash()
			return
		}
		if (winControl="Edit7"){
			SimpleClip:=1
			Clipboard:=GetAllBatches("`r`n")
			SimpleClip:=
			Clipbar.Flash()
			return
		}
		else
			LMS.Menu()
		return

	; F7::copyLabelCopyDoc()
; 	; wheelup::
; 	; If NAdd
; 	;   {
; 	; 	; sleep 500
; 	;   return
; 	;   }
; 	; 	NAdd:=1
; 	; 	blockinput, On
; 	; ControlGetFocus,winControl,ClipBar
; 	; settimer, Block_Input,-500
; if (wincontrol="Edit5"){
; 	Clipbar.AddIteration(100)
; 	return
; 	}

; 	if (winControl="Edit1"){
; 		sleep 400
; 	  GetAllProducts(" ", 1)
; 		clipboard:=AllProducts
; 	TT(AllProducts,2000,ClipBar_x2,35,2,250)
; 		return
; 	  }
; if (winControl="Edit2"){
; 	GetAllBatches(" ", 1)
; 		sleep 400
; 		clipboard:=AllBatches
; 	TT(AllBatches,2000,ClipBar_x2,35,2,250)
; 		return
; 	}
; 	if (winControl="Edit3")
; 	  PriorCodesMenu(1)
; 	if (winControl="Edit4")
; 	  PriorCodesMenu(1)
; 	if (winControl="Edit6"){
; 	  WholeBatchMenu()
; 	    ; ControlsetText, Edit6,%AllBatches%,ClipBar
; 	    ; TT(AllBatches,2000,ClipBar_x2,35,2,250)
; 	  }
; else ;(winControl="Edit3")
; 	wincontrol:=
; 	  sleep 500
; 	  ; NAdd:=
; 	return




; 	wheeldown::
; 	If Nsub
; 	  {
; 		; sleep 500
; 	  Return
; 	  }
; 	ControlGetFocus,winControl,ClipBar
; Nsub:=1
; 		settimer, Block_Input,-500
; if (wincontrol="Edit5"){
; 	Clipbar.SubIteration(100)
; 	return
; }
; 	if (winControl="Edit1"){
; 		sleep 400
; 	  GetAllProducts("`r`n")
; 		clipboard:=AllProducts
; 	; TT(AllProducts,2000,ClipBar_x2,35,2,250)
; 		return
; 	  }
; if (winControl="Edit2"){
; 	GetAllBatches("`r`n")
; 		sleep 400
; 		clipboard:=AllBatches
; 	; TT(AllBatches,2000,ClipBar_x2,35,2,250)
; 	return
; }
; 	if (winControl="Edit3")
; 	  PriorCodesMenu(1)
; ; 	if (winControl="Edit4")
; ; 	  PriorCodesMenu(1)
; ; 	 if (winControl="Edit6"){
; ; 	  MsgBox, 4, Delete file, Do you want to delete WholeBatches.txt?
; ; 	  IfMsgBox, OK
; ; 	      FileDelete, WholeBatches.txt
; ; 	  }
; ; else ;(winControl="Edit3")
; ; 	wincontrol:=
; ; 	  sleep 500
; return


Mbutton::
	ControlGetFocus,winControl,ClipBar
		if (wincontrol="Edit5")
			worktab.CustomerMenu()
		; else if (winControl="Edit1")
			; LMS.SearchBar(Product,"{enter}",0)
		; else if (winControl="Edit2")
			; LMS.SearchBar(Batch,"{enter}",0)
		else if (winControl="Edit3")
			ControlsetText, Edit3,,ClipBar
		else if (winControl="Edit4")
			ControlgetText, clipboard,Edit6,ClipBar
		else if (winControl="Edit7"){
			ControlgetText, clipboard, Edit7,ClipBar
		}
			; clipboard:=Sampleguid
		else
			MenuCodeSelect()
	return


#if

#Ifwinactive, ClipBar ahk_exe VQ_Helper.exe
	enter::
		GUI, ClipBar:default
	Send, ^{a}^{c}
		sleep 100
		LMS.Searchbar(clipboard,"{enter}")
	return
	+enter::
		GUI, ClipBar:default
		sleep 100
		Send, ^a^c
		LMS.Searchbar(clipboard,"{enter}","False")
	return
; F7::copyLabelCopyDoc()
	; 	Gui, ClipBar:submit, nohide
	; }
	; 	else if (winControl="Edit4"){
	; 		Coated:=
	; 		GUI, ClipBar:default
	;
	; 		Gui, ClipBar:submit, nohide
	; 		iniwrite, Coated, Settings.ini, SavedVariables, Coated
	; 	}
	; Mbutton::
	; 		; ControlGetFocus,winControl,ClipBar ahk_exe VQ_Helper
	; 		; MouseGetPos, , , winid, wincontrol
	; 		; if (winControl="Edit1") || (winControl="Edit2") || (winControl="Edit3"){
	; 			ClipBar.Menu()
	; 			return
	; }
	; else if (winControl="Edit4"){
	; Coated:=
	; GUI, ClipBar:default
	; ControlsetText, Edit4,%Coated%,ClipBar ahk_exe VQ_Helper
	; Gui, ClipBar:submit, nohide
	; iniwrite, Coated, Settings.ini, SavedVariables, Coated
	; }

#ifwinactive, Edit test `(Field Configuration: I`, Analytical`)
	mbutton::MsgBox, % "yo"
	F10::Send,{Click, 402, 284}{end}{down 2}{shiftdown}{9}{shiftup}on sample log{shiftdown}{0}{shiftup}{click, 334, 618}

#ifWinExist, Select Product ahk_exe EXCEL.EXE
	Mbutton::
		WinActivate, Select Product ahk_exe EXCEL.EXE
		sendinput, {Click 112, 63}%Product%{enter}{enter}
		; SelectProductToggle:=
	WinWaitActive, ahk_class XLMAIN ahk_exe EXCEL.EXE, ,3
		; if !Errorlevel
		WinRestore, ahk_class XLMAIN ahk_exe EXCEL.EXE
			WinMove, Book ahk_class XLMAIN ahk_exe EXCEL.EXE,, 4, 1, 1150, 1200
		; WinMove, Book ahk_class XLMAIN ahk_exe EXCEL.EXE,, 4, 1, 1150, 1200
	return
#ifwinactive, ahk_class XLMAIN ahk_exe EXCEL.EXE
F6::
	XL.Range("G28").Select
	send, ^{v}
		return
F7::WinMove, ahk_class XLMAIN ahk_exe EXCEL.EXE,, %NuX%, %NuY%, 1250, 1200
		; MouseGetPos, mx, mY

#ifWinExist, LMS Actions ahk_exe EXCEL.EXE
	+enter::

	Mbutton::
		WinActivate, LMS Actions ahk_exe EXCEL.EXE
		sendinput, {click 45, 43}
		winwaitactive, Select Product ahk_exe EXCEL.EXE,, 5
		if !ErrorLevel
			sendinput, {Click 112, 63}%Product%{enter}{enter}
	return


	;-----------------------------------------------------------------------------
	;;[[*** WINDOWS KEYBINDINGS                                               ]]
	;-----------------------------------------------------------------------------

;;\\  									  Edit sample template																	_
		#ifwinactive, Edit sample template
			Mbutton::Sendinput,{Click 256,85}{Home}%Product%
			F6::Sendinput,{Click 256,85}%Product%, `In Process, Analytical{tab 3}{right 6}{tab}{right}
			F7::Sendinput, %Product%
			:*:fm`;::
				Sendinput, `Finished, Micro{tab 3}{right 5}{tab}{right 2}
			Return
			:*:ia`;::
				sendinput, `In Process, Analytical{tab 3}{right 6}{tab}{right}
			Return
			:*:iaa`;::
				Sendinput, `In Process, Analytical (Annual){tab 3}{right 6}{tab}{right}
			Return
			:*:ip`;::
				Sendinput, `In Process, Physical{tab 3}{right 7}{tab}{right}
			Return




;;\\  									    Edit specification																	_
		#ifwinactive, Edit specification
			mbutton::SpecTab.Edit_Analytical()
			Enter::LMSClick.OK()
			+Enter::Sendinput, {enter}




		#Ifwinactive, Edit Formulation


		#IFwinexist, Release: Rotational Testing Schedule ;
			mbutton::
				winactivate,Release: Rotational Testing Schedule
				clk(128, 140)
			return
		#ifwinexist, Release:
			mbutton::
				winactivate, Release:
				clk(128, 140)
			return
		#ifwinexist, Sign :
			mbutton::LMSClick.password()


;;\\  									    Select methods tests																	_
		#ifwinactive, Select methods tests
			+enter::clk2(854, 658) ;select okay
			#enter::clk2(854, 658) ;select okay
			^enter::clk2(854, 658) ;select okay
			F6::clk2(511, 337) ;move over test
			; F6::clk2(511, 375) ;move test back over
			F9::LMS.Menu()
			; F8::clk2(854, 658) ;select okay
			mbutton::
			F10::Spectab.MethodsDropdown()
			F8::Spectab.MethodsDropdown()
					; sendinput, {click}
					; sleep 20
					; click 511, 336 ; click move over
					; return

;;\\  									    Select samples for test																	_
		#Ifwinactive, Select samples for test:
			Mbutton::sendinput, {click 248, 68}{up} ;click dropdown then
			F10::
			F8::WorkTab.SelectTestSample()
			+enter::Clk(503, 290) ;click okay
			F9::send % Clk(250, 70) ;"{up}" ; click okay.
			F7::Clk(853, 657) ;click arrow
				; Clk(853, 657) ;click okay
				; return
			+F6::SpecTab.Methods()
			F6::SpecTab.MethodsDropdown()

;;\\  									    Select tests for request:																	_
		#ifwinactive, Select tests for request
			+F10:: ;Select all tests and move over
					click, 31, 102 ; click Select All Box
					sleep 700
					click 511, 336 ; click move over
					sleep 300
					Click 668, 151
					sleep 800
					WorkTab.SelectTestSample()
					sleep 1000
					mousemove 837, 655
					return
				F10::Clk2(511,336,1, 1) ; move over
				F8::
					clk(311,68) ;Click arrow
					sleep 200
					clk(249,129) ;Select Methods
					sleep 300
					clk(123,68,"Left",2)
					Spectab.MethodsDropdown()
					return
				; Mbutton::SpecTab.MethodsDropdown()
				+mbutton::SpecTab.Methods()
				F7::WorkTab.SelectTestSample()
				mbutton::
					sendinput, {click}
					sleep 20
					click 511, 336 ; click move over
					return
				F6::mouseclick, Left, 638, 70 ;click assign Sample

;;\\  									    Select Product																	_
		#ifwinactive, Select Product ahk_exe eln.exe
			mbutton::send % clk(107, 66) Product "{enter}{enter}"





;;\\  									    Edit Product																	_
		#ifwinactive, Edit Product
			mbutton::ProductTab.AddNewProduct()
			Enter::clk2(285, 574, 1)
			+Enter::sendinput, {enter}


;;\\  									    Edit Formulation																	_
		#Ifwinactive, Edit Formulation
			mbutton::ProductTab.AddNewFormulation()
			wheeldown::
					if MouseMoved:=1
					{
						sendinput, {wheeldown}
						Return
					}
					else
					clk(452, 473,,1)  ;add scroll block
					MouseMoved:=1
					winwaitclose,
					MouseMoved:=
					return

				:*R:00e`;::#00 elongated capsule / 0.995" x 0.336"
				:*R:00e`;::#00 elongated capsule / 0.995" x 0.336"
				:*R:00`;::#00 capsule / 0.917" x 0.336"
				:*R:3`;::#3 capsule / 0.626" x 0.229"
				:*R:2`;::#2 capsule / 0.709" x 0.250"
				:*R:1`;::#1 capsule / 0.765" x 0.272"
				:*R:0`;::#0 capsule / 0.854" x 0.300"
				:*:USP`;::Meets USP Requirements
				:*R:7/16`;::Round / 0.4375"
				:*R:5.5o`;::Oblong / 0.750" x 0.313"
				:*R:5.5ov`;::Oval / 0.625" x 0.344"
				:*R:5o`;::Oblong / 0.750" x 0.250"
				:*R:1c;::Each (1) capsule contains
				:*R:2c`;::Each two (2) capsules contains
				:*R:3c`;::Each three (3) capsules contains
				:*R:4c`;::Each four (4) capsules contains
				:*R:5c`;::Each five (5) capsules contains
				:*R:6c`;::Each six (6) capsules contains
				:*R:7c`;::Each seven (7) capsules contains
				:*R:1t`;::Each (1) tablet contains
				:*R:2t`;::Each two (2) tablets contains
				:*R:3t`;::Each three (3) tablets contains
				:*R:4t`;::Each four (4) tablets contains
				:*R:5t`;::Each five (5) tablets contains
				:*R:6t`;::Each six (6) tablets contains
				:*R:7t`;::Each seven (7) tablets contains
				:*:1s`;::
					sendinput, `Each (1) scoop (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:2s`;::
					sendinput, `Each two (2) scoops (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:3s`;::
					sendinput, `Each three (3) scoops (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:1p`;::
					sendinput, `Each (1) pouch (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:1j`;::
					sendinput, `Each (1) jar (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:1sp`;::
					sendinput, `Each (1) stick packet (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:1ps`;::
					sendinput, `Each (1) packet or scoop (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:2sp`;::
					sendinput, `Each two (2) stick packet (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return
				:*:3sp`;::
					sendinput, `Each three (3) stick packet (%ShapeAndSize% g) contains{tab}^{a}Blend{tab}^{a}%color%+{tab 2}^{left 3}{left}
					return


;;\\  									    Composition																	_
		#ifwinactive, Composition
			Mbutton::
			^v::
				Clk(56, 66) ;click add
				winwait, Edit Ingredient
				gosub, Paste_Clipped_Ingredient
			return



;;\\  									    Edit Ingredient																	_
		#Ifwinactive,Edit Ingredient
			Mbutton::
			Paste_Clipped_Ingredient:
				mouseclick, left, 244, 133,1,0
				sendinput, {tab 4}^{a}%clipped_Position%{tab}%clipped_LabelName%{tab 2}%clipped_LabelClaim% ;click potencty box
				if (clipped_IngredientGroup)
					Sendinput, {Tab}%clipped_IngredientGroup%	;ingredientgroup
				mousemove, 280, 558, 0
				Tooltip
			return
			F9::ProductTab.IngredientsMenu()
				:*:1s`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of (1) scoop (%ShapeAndSize% g){left 2}
				return
				:*:2s`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of two (2) scoops (%ShapeAndSize% g){left 2}
				return
				:*:3s`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of three (3) scoops (%ShapeAndSize% g){left 2}
				return
				:*:1p`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of (1) pouch (%ShapeAndSize% g){left 2}
				return
				:*:1j`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of (1) jar (%ShapeAndSize% g){left 2}
				return
				:*:1sp`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of (1) stick packet (%ShapeAndSize% g){left 2}
				return
				:*:1ps`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of (1) packet or scoop (%ShapeAndSize% g){left 2}
				return
				:*:2sp`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of two (2) stick packet (%ShapeAndSize% g){left 2}
				return
				:*:3sp`;::
					sendinput, ^{a}* Heavy Metals results are based on a daily dose of three (3) stick packet (%ShapeAndSize% g){left 2}
				return





;;\\  									    Edit request																	_
		#ifwinactive, Edit request
			mbutton::WorkTab.EditRequest()
			F10::
			winactivate, Edit request
			click 239, 617
			return

;;\\  									    Edit sample																	_
		#ifwinactive, Edit sample (
		F9::worktab.CustomerMenu()
		Mbutton::Worktab.Dropdown_CustomerSelect(Iteration)


;;\\  									    Edit test (Field Configuration																	_
		#ifwinactive, Edit test (Field Configuration
		; lctrl::
		printscreen::
			F10::Send,{Click, 402, 284}{end}{down 2}{shiftdown}{9}{shiftup}on sample log{shiftdown}{0}{shiftup}{click, 334, 618}
			; +F10::
			; 	sendinput, {click 384, 222}{tab 2}{end 2}
			; 	sendinput, {shiftdown}{9}{shiftup}on sample log{shiftdown}{0}{shiftup}
			; 	sendinput, {Click 334, 618}
			; 	return
			; wheelright::
			F7::
			winactivate, Edit Test
			mbutton::
					preclk()
					MouseClick, left, 464, 533,1,0 ; Click scrollbar
					; MouseClick, left, 245, 485,1,0  ; click results
					sleep 20
					click 239, 538 ; click results
					postclk()
		return



;;\\  									    Result Entry																	_
		#Ifwinactive, Result Entry
			Enter::LMSClick.OK()
			F7::numbermenu(6,"toggle") ;WorkTab.CorrectTestResults("toggle", "Loop")
			F6::numbermenu(6,0) ;WorkTab.CorrectTestResults("toggle", "Loop")
			; F6::WorkTab.CorrectTestResults(0,5)

			Mbutton::WorkTab.CorrectTestResults("Toggle")
			F9::LMS.Menu()
			F10::
			+Mbutton::WorkTab.CorrectTestResults(0,"","Test")
			F8::numbermenu(6)

;;\\  									    Results Definition																	_
		#Ifwinactive, Results Definition
			+mbutton::SpecTab.Autofill()
			Mbutton::
				if winexist("Spec Table ahk_class AutoHotkeyGUI")
					SpecTab.AutoInputResultEditor()
				else If (!MinLimit && !MaxLimit && !winexist("Spec Table ahk_class AutoHotkeyGUI"))
					SpecTab.PasteClipboardIntoSpec()
				Else
					SpecTab.Autofill()
				Return


		;mbutton::Spectab.PasteClipboardIntoSpec()
			F7::lmsclick.edit()
			F9::lms.menu()
			Backspace::LMSClick.Remove()
		;Backspace::sendinput, {Click %RemoveButton%}

;;\\  									    Results																	_
		#ifwinactive, Results
			F7::lmsclick.edit()
			F9::lms.menu()
			Mbutton::SpecTab.AutoInputResultEditor()
			+mbutton::SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,FullRequirements)
		F6::spectab.toggleUseLimitsFromTheTest()

;;\\  									    Result Editor																	_
		#ifwinactive, Result Editor
		+Enter::send, {enter}
		Enter::LMSclick.OK()

		+Tab::clk2(333,670,1,2) ;click requirement box
		down::clk2(333,670,1,2) ;click requirement box
		\::clk2(362,463,1,2) ;click lower limmit box

		Up::clk2(295,192,1,2) ;click units box
		; +Tab::send, {tab}
			mbutton::
			if winexist("Spec Table ahk_class AutoHotkeyGUI")
				SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,FullRequirements)
			else
				SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,0,Clipped_Requirements)
			return


		wheelDown::
			if Mousemoved:=1
				return
			Else
				clk(503, 574,1,,,0)
				; clk(288, 318,1,2,,0)
				mousemoved:=1
			SetTimer, Block_Input, -5000
		return
			`;::clk(405, 534,,2)
			+`;::sendraw, :
		F6::spectab.toggleUseLimitsFromTheTest()
		return


;;\\  									    Test Definition Editor																	_
		#Ifwinactive,Test Definition Editor
			mbutton::
				sleep 200
				if (Description)
					SpecTab.TestDefinitionEditor(Description) ; the pre window
				Breaking.Point()
				sleep 200
				MouseClick, left, 464, 532,2,0 ;click scrollbar
				click 239, 246 ;results link
				sleep 200
				Breaking.Point()
				winactivate, Results Definition
				sleep 100
		return
		Enter::Clk2(343, 616)
		+Enter::sendinput, {Enter}

		wheeldown::
			if Mousemoved
			{
					sendinput, {wheeldown}
					return
			}
			else
			{
				clk(464, 532,,2,"Test Definition Editor",0)
				mousemove, 245, 246, 0
			mousemoved:=1
			winwaitclose,
			mousemoved:=
			}
		return
		F7::
			clk(464, 532,,2,"Test Definition Editor",0)
			click, 245, 246
			sleep 500 ;add scroll block
			return
			; 		mbutton::mouseclick, left, 333, 615




;;\\  									    Register new samples																	_
		#ifwinactive, Register new samples
		F7::sendinput,{Click 505, 356}{Click 860, 661}
		F6::
				clk(181, 104,2,2)
				sleep 300
				Send, %Product%{enter}
			return
			F10::
		+Mbutton::
			if Lot
				loopingCount:=4
			else
				LoopingCount:=3
				loop, %LoopingCount%
				{
					MouseGetPos, rnmx, rnmy
					Breaking.Point()
					WorkTab.registerNewSamples()
					sleep 300
					Breaking.Point()
					rnmy:=RnMy+26
					MouseMove, %Rnmx%, %rnmy%, 0
					sleep 100
				}
				Breaking.Point()
			return
			mbutton::WorkTab.registerNewSamples()


;;\\  									REASON  Change                                .																	_
		#ifwinactive, Reason For Change ahk_exe eln.exe ahk_pid 27396)
			F7::sendinput, {enter}
			F10::
			F8::Sendinput, {tab}^{a}Update Spec Version{Click 78 366 0}
			Enter::clk2(241, 236)
			+Enter::sendinput, {enter}

		#ifwinactive, Reason For Change
			F7::Sendinput, {Click 160 130}^{a}Correcting Notebook Page#{Click 240 237 0}
			F10::
			F8::Sendinput, {Click 160 130}^{a}Fixing Rotation{Click 240 237 0}
			Enter::clk2(241, 236)
			+Enter::sendinput, {enter}
			F9::sendinput, {tab 2}{space}{down}{enter}{tab 2}{Click 240 237 0} ; use revious reason





;;\\____ NuGenesis MAIN                                     .
#Ifwinactive, NuGenesis LMS
Enter::
	LMS.SaveCode()
	sleep 200
	send, {Enter}
	return
+Enter::LMS.SaveCode()
+^F11::LMS.OrientSearchbar()
!F4::Return
^F9::
	send, ^c
	sleep 500
		If (SampleguID){
				; FileRead, PreviousSampleguIDs, % PreviousSampleguIDsFile
				  ; {
				; NewPreviousSampleguIDs:=Trim(RemoveDuplicates(PreviousSampleguIDs)"`n"SampleguID)
					FileDelete, %PreviousSampleguIDsFile%
					sleep 200
					FileAppend, %NewPreviousSampleguIDs%, %PreviousSampleguIDsFile%
						; return
			; }
		}
		clipboard:=sampleguid
		TT(Sampleguid)
	return
	!F10::LMS.AddsampleLog(5)
	!F9::
		If (LMS.DetectTab() != "Requests"){
			send, {click 40 40}
			sleep 50
			send,{click 50 75 0}
			sleep 50
			send,{click 280 75 0}
			sleep 100
			send,{click 280 220}
		}
		else
			return
	return
^F7::FileRead, Clipboard, U:\VQ_Helper\ClippedExcelData.txt
	; SpecTab.Table()
	;^F10::LMS.AddSampleLog(15)
	; F10::
mbutton::
	Tab:=LMS.DetectTab()
	if (Tab = "Specs"){
		if winexist("Spec Table ahk_class AutoHotkeyGUI") {
			ContinueToRun:=1
			MouseGetPos, mbX, mbY
		If winactive("Test Definition Editor") || winactive("NuGenesis LMS")
				SpecTab.AutoInputTestDefinitionEditor()

		else if winactive("Result Editor") || winactive("Results Definition")
			SpecTab.AutoInputResultEditor()
			; ; else if winactive("Edit request")

			; else winactive("NuGenesis LMS")
			ContinueToRun:=
			winwaitactive, NuGenesis LMS,,5
			mby:=Mby+26
			sleep 200
			clk2(mbx,mby, 0)
			sleep 200
			; send, {down}
					; SpecTab.AutoInputSpecResults()
				return
			}
			else
				SpecTab.CopySpecTemplate()
			; Sendinput, {ctrl down}{click}{ctrl up}
	}
	else if (Tab = "Samples"){    ; assign all non-retain tests to new requFsamests
		clk2(75,298) ; click assign to new reuest
		WinWaitActive, Edit request,,3
		; winactivate, Edit request
		click 239, 617 ; click pick sample
		WinWaitActive, Select tests for request,,3
		click, 473,66 ; click filter toggle
		sleep 200
		click 78, 127 ; click filter box
		sendinput, ^{a}cal{enter} ;type "cal" to filter out retains
		click, 41, 108 ;click select all
		click, 511,336 ; click move arrow
	}
	else
		Sendinput, {ctrl down}{click}{ctrl up}
	; Else
		; Send, ^{mbutton}
return

	F9::lms.Menu()
	+F9::lms.Menu("shift")
	F7::
		if (AllBatches)
			Batch:=AllBatches
		LMS.SearchBar(Batch,"{enter}",0)
		; Clipboard:=Batch
		return
	F6::
		if (AllProducts)
			Product:=AllProducts
		LMS.SearchBar(Product,"{enter}",0)
		; Clipboard:=Product
		return

	+F7::
		Batch:=GetAllBatches("`n")
		sleep 200
		Clipboard:=Batch
		; LMS.SearchBar(Batch,"{enter}",0)
		return
	+F6::
		Product:=GetAllProducts("`n")
		sleep 200
		Clipboard:=Product
		; LMS.SearchBar(Product,"{enter}",0)
		return
^+F6::LMSClick.Edit_Composition()
	; +F6::sendinput % GetAllProducts(" ")
	; +F7::sendinput % GetAllBatches(" ")

		F8::LMS.SearchBar("",,"False")
	; +F6::LMSClick.EnterResults()


		; clk(54,734,"",1,"NuGenesis LMS",2)  ;:((Delete Test))

	+F8::LMS.SearchBar("",,"False")

	+#v::LMS.Searchbarpaste(";")
	+^v::LMS.Searchbarpaste(";")
	^v::LMS.Searchbarpaste(A_space)
	+!v::LMS.Searchbarpaste(A_space)
	; ~Enter::LMS.SaveCode()


;;\\____eln.EXE ____________                                                       .
#Ifwinactive, ahk_exe eln.exe
	F1::Sendinput, %Product%
	+F1::sendinput % GetAllProducts(" ")
	+^F1::sendinput % GetAllProducts("`;")
	F2::Sendinput, %Batch%
	+F2::sendinput % GetAllBatches(" ")
	+^F2::sendinput % GetAllBatches("`;")
	F3::Sendinput, %lot%
	F4::Sendinput, %Coated%
	!1::Sendinput, %Product%
	!2::Sendinput, %Batch%
	!3::Sendinput, %lot%
	!4::Sendinput, %Coated%
	+^1::sendinput % GetAllProducts("`;")
	+!1::sendinput % GetAllProducts("`n")
	+^2::sendinput % GetAllBatches("`;")





	^+F10::
			tt("Chagne Percision: 3")
			lms.ChangePercision(3)
			return
	+F10::
			tt("Chagne Percision: 2")
			lms.ChangePercision(2)
			Return
	^F10::
		loop 40
		{
		breaking.point()
		lms.ChangePercision(2)
		WinWaitActive, NuGenesis LMS,,10
		if errorlevel
			exit
		sleep 1300
		breaking.point()
		}
		return

	enter::Enter
	+enter::LMSclick.OK()
	esc::LMSclick.esc()
	; F6::Send, %Product%
+F5::Sendinput, %SampleGUID%
^F5::Clipboard:=SampleGUID
	; +mbutton::lms.Menu()
	+F9::lms.Menu("Shift")
	F9::lms.Menu()

	<!left::GetAllProducts()
	<!right::GetAllBatches()
	F13::breaking.point(1)
	+!F5::LMS.Menu()
	^Space::LMS.SearchBar("",,"False")
	!^Space::LMS.SearchBar("","{delete}","False")
	+!F10::LMS.AddDataFromClipboard()

	pause::						Suspend, Toggle
	#h::return
	#p::return
	#k::return
	>+Backspace::			Delete
	;F8::esc




	:*:osl`;::`(On Sample Log)
	:*:fm`;::`Finished, Micro
	:*:ia`;::`In Process, Analytical
	;:*:pia`;::%product%`, In Process, Analytical
	:*:pm`;::
		Sendinput, %product%`, Finished, Micro{tab 3}{right 5}{tab}{right 2}
	Return
	Return
	:*:paa`;::
		Sendinput, %product%`, In Process, Analytical (Annual){tab 3}{right 6}{tab}{right}
	Return
	:*:pp`;::
		Sendinput, `In Process, Physical{tab 3}{right 7}{tab}{right}
	return
	:*:pr`;::
		Sendinput, `In Process, Retain{tab 3}{right 7}{tab}{right}
	return
	:*:iaa`;::`In Process, Analytical (Annual)
	;:*:piaa`;::%Product%`, In Process, Analytical (Annual)
	:*:ip`;::`In Process, Physical
	:*:ir`;::`In Process, Retain
	:*:st`;::`Stability
	:*:cr`;::`Coated, Retain
	:*:ca`;::`Coated, Analytical
	:*:cp`;::`Coated, Physical
	:*:in`;::`ingredient
	; :*:Pa`;::`P. aeruginosa
	:*:uc`;::`Update Total Coliforms Method

;;\\____explorer.exe ______
#ifwinactive, ahk_class CabinetWClass ahk_exe explorer.exe
	+F9::send, ^e
	^w::
				winactivate, ahk_exe explorer.exe
        Send, !f
				sleep 150
        Send, !c
				Send, !{f4}
    return
		winactivate, *%Product%* ahk_exe explorer.exe
		sleep 200
		send ^{e}
		sleep 340
		sendinput, {*}%Product%{*}
		sleep 700
		send, ^{e}{enter}{down 2}{up}
		; sendinput, {tab 2}{right}{pgup 2}
	return
	F10::
	F8::send, ^{F4}
Mbutton::
	send, ^{e}{*}%Product%{*}
	sleep 800
	send, ^{e}{enter}{down 2}{up}^a
	return
	; +Mbutton::SelectPreviewPane()
	+F6::
		winactivate, ahk_class CabinetWClass ahk_exe explorer.exe,,VQ_Helper
		send, ^{e}
		sleep 700
	send, {enter}{down 2}{up}
	sleep 700
		send ^{a}^{c}
	return
	F6::WinMove, ahk_exe explorer.exe,, 5, 10, 1250, 1200
#ifwinexist, Search Results ahk_exe explorer.exe
+F6::
	winactivate, ahk_exe explorer.exe
	sleep 400
	send, ^{e}{*}%Product%{*}
	sleep 900
	send, ^{e}{enter}{down 2}{up}
return


#ifwinactive



Test_msgbox(msg){
	global
	; msgbox % " Name: " Name[] "`n Claim: " Claim[] "`n Position: " Position[] "`n IngredientID: " IngredientID[]
	; TT(Listarray(Name))
	TT(msg,2000)
	; listarray(Position),1000,300,300,2,,"R")
	return
}
Ctest_1:
return










AddToList(){
	; ListItem:=[]
	TT("AddToList")
	; ListItem:=StrSplit(ParseText,"`n")
	CoordMode, mouse, window
	Loop, parse, clipboard, "`n"
	{
		TT(A_index ": `n" A_LoopField,2000)
		Breaking.Point()
		click, 407, 73 ;click New Value
		sleep 200
		click, 608, 201 ;click First EditBox
		sleep 200
		sendinput, %A_loopField%
		sleep 400
		Breaking.Point()
	}
	return
	; CoordMode, mouse, Screen
}



;---------------------------------------------------------------------------------
;;[[*** MENU HANDLES                                               ]]
;---------------------------------------------------------------------------------
Reloadsub(){
	Critical, on
	Thread, Priority, 1
	reload
}
StopTimer(){
	SetTimer, activeCheck, Off
}
Exitsub(){
	global
	ifwinnotexist, ahk_exe explorer.exe
		run, explorer "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper"
	exitApp
}
CopyGUID(){
	global
	Clipboard:=SampleGUID
}

windowSpy(){
	Run, WS.exe
}
get_window_info:
	if !getTitle
		WingetTitle getTitle, A
	simpleclip:=1
	Clipboard:=getTitle
	Clipwait,1
	sleep 500
	simpleclip:=
	tt(getTitle,1000,100,100)
return

get_mouse_info:
	MouseGetPos, getX, getY, getWin
	simpleclip:=1
	Clipboard:=getX ", " GetY
	clipwait,1
	sleep 300
	simpleclip:=
	tt(getX ", " GetY,2000,100,50,,,"M")
	return

return
AllBatchesMsgbox:
	AllBatchesMsg:=Trim(GetAllBatches(" ",1))
	tt(AllBatchesMsg)
	; clip.EditBox(AllBatchesMsg)
return
AllBatchesNoMsgbox:
	AllBatchesMsg:=Trim(GetAllBatches(" "))
	tt(AllBatchesMsg)
	; clip.EditBox(AllBatchesMsg)
return
AllProductsMsgbox:
	AllProductsMsg:=Trim(GetAllProducts(" ",1))
	sleep 600
	tt(AllProductsMsg)
	; clip.EditBox(AllProductsMsg)
return
AllProductsNoMsgbox:
	AllProductsMsg:=Trim(GetAllProducts(" "))
	sleep 600
	tt(AllProductsMsg)
	; clip.EditBox(AllProductsMsg)
return
ShowWholeBatches:
	run, edit WholeBatches.txt
return
Add15SampleLog:
	LMS.AddsampleLog(5)
Return
ShowVariables:
	listvars
return
ListLines:
	ListLines
return
ShowHotkeys:
msgbox, %HotkeysTip%

;---------------------------File Path Labels

mmigninFolder(){
	global
	if !winexist("VQ_Helper ahk_exe explorer.exe")
		run, explorer "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper"
	else
		winactivate, VQ_Helper ahk_exe explorer.exe
}
SettingsFile(){
		Run , Edit "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\Settings.ini"
}
EditMethodList(){
		Run , Edit "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\Methods.ini"
}
ShowFinalLabelCopy:
	; run, find "\\10.1.2.118\Label Copy Final"
	runwait, find "\\netapp\Label Copy Final"
	sleep 550
	; winmaximize, Search Results
	winactivate, ahk_class CabinetWClass ahk_exe explorer.exe
	send, {*}%Product%{*}{enter}
	sleep 400
	send, ^{e}{tab 2}{right}
	; SelectPreviewPane(Product)
return
ShowScanLabelCopy:
	runwait, find "\\10.1.2.118\share\QC LAB\Label Copy Scans"
	; runwait, find "\\netapp\share\QC LAB\Label Copy Scans"
	sleep 150
	; winmaximize, Search Results
	WinMove, *%Product%* ahk_exe explorer.exe,, 5, 10, 1250, 1200
	winactivate, *%Product%* ahk_exe explorer.exe
	send, {*}%Product%{*}{enter}
	sleep 700
	send, ^{e}{tab 2}{Right}
return
ShowManualCOA:
	run, explorer "\\10.1.2.118\coa-lot#"
return
DeletePriorCodes:
	FileDelete, PriorCodes.txt
	tt("Deleted")
Return
DeleteWholeBatches:
	FileDelete, WholeBatches.txt
	tt("Deleted")
Return
^F10::
TestCode:
	sleep 500
	autohotkeyPath := "U:\VQ_Helper\RawFiles\AHK\AutoHotkeyU64.exe"
	scriptPath := "U:\VQ_Helper\RawFiles\TestingCode.ahk"
	RunWait, %autohotkeyPath% "%scriptPath%"
return
Showmfg:
	run, find "\\10.1.2.118\lms\Information\ECOPY\mfg"
	sleep 200
	sendinput, {*}%Product%{*}{enter}
	sleep 300
	send, ^{e}{tab 2}{right}
return
ShowGlobalVision:
	runwait, find "\\10.1.2.118\Globalvision Master Copy Files"
	sleep 200
	sendinput, {*}%Product%{*}{enter}
	; winmaximize, Search Results
	sleep 300
	send, ^{e}{tab 2}{right}
return
ShowFINAL_C_O_A:
	run, explorer "\\10.1.2.118\final_c_o_a\2023 CoAs"
return
ShowC_O_A:
	run, explorer "\\10.1.2.118\c_o_a"
return
ShowEditBox:
	clip.editbox()
return

FtenMenuHandler:
	Ften := A_ThisMenuItem
	IniWrite, %Ften%, Settings.ini, Config, Ften
return














;;-------------------------------------------------------------
					;;[[***UTILITIES                                               ]]


ReadIniFiles(){
	global
	iniRead, Product, Settings.ini, SavedVariables, Product
	iniRead, Batch, Settings.ini, SavedVariables, Batch
	iniRead, Lot, Settings.ini, SavedVariables, Lot
	iniRead, Coated, Settings.ini, SavedVariables, Coated
	; iniRead, GeneralBox, Settings.ini, SavedVariables, GeneralBox
	iniRead, SampleGUID, Settings.ini, SavedVariables, SampleGUID
	iniRead, Iteration, Settings.ini, SavedVariables, Iteration
	iniRead, ClipbarLocation, Settings.ini, Config, ClipbarLocation
	iniRead, CompileTime, Settings.ini, Config, CompileTime
	; iniRead, Version, Settings.ini, Config, Version
	; iniRead, GeneralBox, Settings.ini, SavedVariables, GeneralBox
	; iniRead, SampleGUIDMode, Settings.ini, SavedVariables, SampleGUIDMode
	; iniread, PriorCodeString, Settings.ini, SavedVariables, PriorCodeString
	iniread, CodeString, Settings.ini, SavedVariables, CodeString
	iniRead, Ingredient_List_Adjustment, Settings.ini, Config, Ingredient_List_Adjustment
	; iniread, Ften, Settings.ini, Config, Ften

	; iniread, NormalWinDelay, Settings.ini, Config, NormalWinDelay
	iniread, ActiveTimerCheck, Settings.ini, Config, ActiveTimerCheck
	iniRead, IngredientNoteDropDownCount, Settings.ini, Config, IngredientNoteDropDownCount

	iniRead, NewSpecVersionDescriptionText, Settings.ini, SavedVariables, NewSpecVersionDescriptionText
	;iniRead, Iteration, Settings.ini, SavedVariables, Iteration
	; iniRead, Iteration, Settings.ini, SavedVariables, CustomerPosition

	iniRead, HM_Units, Settings.ini, HeavyMetal_Variables, HM_Units
	iniRead, HM_Lower_Limit, Settings.ini, HeavyMetal_Variables, HM_Lower_Limit
	iniRead, Arsenic_Limit, Settings.ini, HeavyMetal_Variables, Arsenic_Limit
	iniRead, Arsenic_Requirement, Settings.ini, HeavyMetal_Variables, Arsenic_Requirement
	iniRead, Lead_Limit, Settings.ini, HeavyMetal_Variables, Lead_Limit
	iniRead, Lead_Requirement, Settings.ini, HeavyMetal_Variables, Lead_Requirement
	iniRead, Cadmium_Limit, Settings.ini, HeavyMetal_Variables, Cadmium_Limit
	iniRead, Cadmium_Requirement, Settings.ini, HeavyMetal_Variables, Cadmium_Requirement
	iniRead, Mercury_Limit, Settings.ini, HeavyMetal_Variables, Mercury_Limit
	iniRead, Mercury_Requirement, Settings.ini, HeavyMetal_Variables, Mercury_Requirement
	iniread, WindowSpyPath, Settings.ini, FilePaths, WindowSpyPath
	iniread, AppIconPath, Settings.ini, FilePaths, AppIconPath
	iniread, CustomerListPath, Settings.ini, FilePaths, CustomerListPath
	iniread, CodeFile, Settings.ini, FilePaths, CodeFile
	; wholebatches = []
	; 	FileRead, WholeBatches, WholeBatches.txt
	; 	sleep 300
	; 	loop, parse, WholeBatches, "`n"
	;   	WholeBatches.insert(a_LoopField)

}

GetIniValue(IniFile,IniKey){
	Loop, Read, %IniFile%
	{
		If A_Index := 1
			Continue
		if instr(A_LoopReadLine, IniKey,false,1,1){
			MatchedLine := StrSplit(A_LoopReadLine, "=")
			; tt(MatchedLine[2],2000,400,400,3)
			return MatchedLine[2]
		}
		else
			continue
	}
	return
}











TT(msg:="yo", time=1500, X:="",Y:="",N:="", Transparent:=240,Position:="S") {
	global simpleclip
	ttMx:=100
	ttmy:=1
	if Position= M
	{
	; 	; CoordMode, tooltip, screen
		CoordMode, tooltip, mouse
		CoordMode, Mouse, screen
	}
	winGetPos, ttwinx,ttwiny, ttwinw, ttwinh, A
	MouseGetPos, mX, mY
	; if Position="W"
	; else
	; {
		CoordMode, tooltip, window
		CoordMode, Mouse, window
	; }
	if Simpleclip
		return
	ttwinx:=Floor(ttwinx+X)
	ttwiny:=Floor(ttwiny+y)
	ttx:=Floor(mx+X)
	tty:=Floor(my+y)
	; CoordMode, ToolTip, Window
	sleep 20
	; CoordMode, mouse, window
	if Position = M
		tooltip, %msg%, %ttx%, %tty%,%N%
	else if Position = W
		tooltip, %msg%, %ttwinx%, %ttwiny%,%N%
	else
		tooltip, %msg%, %ttMX%, %ttmY%,%N%
	; X+=100

	; else
	hwnd := winExist("ahk_class tooltips_class32")
	; if Transparent
		WinSet, Exstyle, 0x20, % "ahk_id" hwnd
	WinSet, AlwaysOnTop, On, % "ahk_id" hwnd
	winSet, Trans, %Transparent%, % "ahk_id" hwnd
	if !(time=0)
		SetTimer, RemoveToolTip%N%, -%time%
		return
		RemoveToolTip:
		ToolTip
		return
		RemoveToolTip1:
		ToolTip,,,,1
		return
		RemoveToolTip2:
		ToolTip,,,,2
		return
		RemoveToolTip3:
		ToolTip,,,,3
		return
		RemoveToolTip4:
		ToolTip,,,,4
		return
		RemoveToolTip5:
		ToolTip,,,,5
		return
		RemoveToolTip6:
		ToolTip,,,,6
		return
		RemoveToolTip7:
		ToolTip,,,,7
		return
		RemoveToolTip8:
		ToolTip,,,,8
		return
		RemoveToolTip9:
		ToolTip,,,,9
		return
		RemoveToolTip10:
		ToolTip,,,,10
		return
		RemoveToolTip11:
		ToolTip,,,,11
		return
		RemoveToolTip12:
		ToolTip,,,,12
		return
		RemoveToolTip13:
		ToolTip,,,,13
		return
		RemoveToolTip14:
		ToolTip,,,,14
		return
		RemoveToolTip15:
		ToolTip,,,,15
		return
		RemoveToolTip16:
		ToolTip,,,,16
		return
		RemoveToolTip17:
		ToolTip,,,,17
		return
		RemoveToolTip18:
		ToolTip,,,,18
		return
		RemoveToolTip19:
		ToolTip,,,,19
		return
		RemoveToolTip20:
		ToolTip,,,,20
		return
	}


MouseIsOver(winTitle){
	Global
	MouseGetPos,,, win, WinControl
	Return winExist(winTitle . " ahk_id " . win)
}

RemoveDuplicates(InputString,Delimiter:="`n"){
	Loop, Parse, InputString, Delimiter
	{
		If Output not contains %A_LoopReadLine%%Delimiter%
			Output .= A_LoopReadLine . Delimiter
	}
	Return Output
}

ListArray(The_Array,The_Delimeter:="n"){
	; global
	if (The_Delimeter<>"n"){
		for Each, Element in The_Array
			ArrayList .=Element The_Delimeter
		return ArrayList
	}
	else {
		For Each, Element In The_Array {
			ArrayList .= "`n" A_index ": "
			ArrayList .= Element
		}
		return ArrayList
	}
}

HasValue(haystack, needle) {
	for index, value in haystack
		if (value = needle)
			return index
	if !(IsObject(haystack))
		throw Exception("Bad haystack!", -1, haystack)
	return 0
}



BlockRepeat(Time:=300, ToolTipMessage:=""){
	Global N
	; critical
	sleep 25
	If N
		exit
	If ToolTipMessage
		TT(TooltipMessage)
	N:=1
	SetTimer, BlockTheInput, -%time%
	sleep 50
	return

	BlockTheInput:
	N:=
	return
}

Clk(x,y,Button:="Left",n=1,window:="",returnMouse:=1){
	global
	mouseGetPos, mx, my, mw,
	MouseReturn:="{click " Mx ", " My ",0}"
	if window
		if !winactive(window)
			sleep 200 ; winactivate, %window%
	click, %x% %y% %n%
	; mouseclick, %Button%, %x%,%y%,%n%,0
	sleep 25
	if (window!="")
		winactivate, %mw%
	If (ReturnMouse=x)
		Return MouseReturn
	If (ReturnMouse=0)
		Return
	else
		mousemove,%mx%,%my%,0,
		return
}

preClk(){
	Global Premx, premy
		CoordMode, Mouse, Screen
		mouseGetPos, PREmx, PREmy, PREmw
		CoordMode, Mouse, window
		return
}
PostClk(){
	Global Premx, premy
	sleep 100
		CoordMode, Mouse, Screen
		mousemove,%PreMx%,%PreMy%,0
		CoordMode, Mouse, window
		PreMX:=
		PreMy:=
		return
}




Clk2(x,y,returnMouse:=1, n:=1){
		Global ClkMx, ClkMy, clkMw
	If (ReturnMouse=1){
		CoordMode, Mouse, Screen
		mouseGetPos, clkmx, clkmy, clkmw
		CoordMode, Mouse, window
	}
	If (ReturnMouse=2){
		mouseGetPos, clkmx, clkmy, clkmw
	}
	click, %x% %y% %n%
	sleep 2
	if (ReturnMouse=1){
	winactivate, %clkmw%
	CoordMode, Mouse, Screen
	mousemove,%clkmx%,%clkmy%,0
	CoordMode, Mouse, window
		}
	if (ReturnMouse=2){
	winactivate, %clkmw%
	mousemove,%clkmx%,%clkmy%,0
		}
	else
		return
}

class Breaking {
	Point(break:=""){
		Global
		critical
		SimpleClip:=
		; If SpecMenuClosed
		; 	{
		; 		SpecMenuClosed:=
		; 		Exit
		; 	}
		If GetKeyState("Lbutton", "P") || break ||GetKeyState("Rwin", "P") || GetKeyState("Rcontrol", "P") ||GetKeyState("Shift", "P") || GetKeyState("RShift", "P") || GetKeyState("LShift", "P") || (A_PriorKey="RShift")|| (A_PriorKey="Rcontrol") || (A_PriorKey="rbutton") {
			blockinput, off
				TT("-`n`n        Broken        `n`n-",3000,100,100,,255,"M")
			Clipbar.Flash()
			sleep 50
			Clipbar.Flash()
			sleep 50
			Clipbar.Flash()
			sleep 50
			Clipbar.Flash()
				; TT("-`n   Broken   `n-",2000)
			exit
		}
		if keep_running = n ;another signal to stop
		{
			blockinput, off
			; Clipbar.Flash()
				TT("-`n`n        Broken        `n`n-",3000,,,,255,"M")
			Exit
		}
	}
	Preamble(){
		Global
		if keep_running = y
		{
			keep_running = n ;signal other thread to stop
			exit
		}
		keep_running = y
	}
}

