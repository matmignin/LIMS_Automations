

#IfWinActive, VarBar ahk_exe AutoHotkey.exe 
	F19 & F20::Send, {tab}{shiftdown}{tab}{shiftup} ;varbar.focus("Batch")
	F19::varbar.focus("Product")
	; F20::varbar.focus("Batch")
	^enter::                           
							winactivate, %the_WinTitle%
							click, %caret_X%, %caret_y%
							return
	; mbutton::
				; Click
				; if Toggle := !Toggle
				;  Notes.SHow()
				; else
					; notes.Save()
				; return
	+numlock::
			MouseGetPos,,,,WinControl
			ControlGetFocus,WinControl,VarBar ahk_exe AutoHotkey.exe
			msgbox, %Wincontrol%
			Send, ^{a}{backspace}{enter}
			ControlsetText, Wincontrol,,VarBar
		Return
	Mbutton::
		click
	enter::
		ControlGetFocus,WinControl,VarBar ahk_exe AutoHotkey.exe	
		if (WinControl="Edit1") || (WinControl="Edit2") || (WinControl="Edit3") ||(WinControl="Edit4"){
					Send, ^a^c
					winactivate, NuGenesis LMS - \\Remote			
			LMS.Searchbar(clipboard,"{enter}")
		}
		if (WinControl="Edit6") || (WinControl="Edit7") || (WinControl="Edit8") || (WinControl="Edit9")
			varbar.show()
		return
	; Lbutton::click
#ifwinactive
	; ^left::Lms.SelectTab("Left")
	; ^right::LMs.SelectTab("right")

#If MouseIsOver("VarBar ahk_exe AutoHotkey.exe")
	; wheelleft::    Excel.PrevSheet()
	; wheelRight::   excel.Nextsheet()
	+Mbutton::
			MouseGetPos,,,,WinControl
			; ControlGetFocus,WinControl,VarBar ahk_exe AutoHotkey.exe
			if (WinControl="Edit1"){
				Product1:=Product
				if Toggle := !Toggle
					GuiControl, Varbar:Text, Product,%Product1%
				else
					GuiControl, Varbar:Text, Product,%Product0%
				}
			if (WinControl="Edit2"){
				Batch:=Batch0
				}
			if (WinControl="Edit3"){
				GuiControl, Varbar:Text, Lot,%Lot%
				}
			if (winControl="Edit4") {
				GuiControl, Varbar:Text, Coated,%Coated%
			}
			return
	Mbutton::
				click
				Send, ^a
				clip()
				winactivate, NuGenesis LMS - \\Remote			
				LMS.Searchbar(clipboard,"{enter}")
				return


	WheelUp::      send % Blockrepeat(400) Varbar.AddIteration()
	Wheeldown::    send % Blockrepeat(400) Varbar.SubIteration()
	; wheelright::	Varbar.AddIteration(0)
	; Wheelleft::   	Varbar.SubIteration(0)
	up::				Varbar.AddIteration(0)
	down::   		Varbar.SubIteration(0)
	F9::           Excel.connect()
	F7::           Excel.NextSheet()
	F6::           Excel.PrevSheet()
	F8::				Varbar.launchTable()
	Numlock::				send, {click}^a
		return
	Rbutton::		menu.Varbar()	
	return
	numpaddot:: 	 Openapp.Workbook()
#if


Class VarBar{	
	Show(X:=1, Y:=1, Destroy:="Reset"){
			Global
			try Gui,VarBar:Destroy
			if !WinExist("LMS Workbook.xlsb"){
				Iniread, Batch, data.ini, SavedVariables, Batch
				Iniread, Product, data.ini, Products, Product
				Iniread, Batch0, data.ini, SavedVariables, Batch0
				Iniread, Batch1, data.ini, SavedVariables, Batch1
				Iniread, SampleID, data.ini, SavedVariables, SampleID
				Iniread, Lot, data.ini, SavedVariables, Lot
				Iniread, Coated, data.ini, SavedVariables, Coated
			}
				Iniread, Iteration, data.ini, SavedVariables, Iteration
				Iniread, ShowSampleID, data.ini, Options, ShowSampleID
				Iniread, ShowCoated, data.ini, Options, ShowSampleID
				Iniread, ShowNote3, data.ini, Options, ShowNote3
				Iniread, SwitchWorkSheets, data.ini, Options, SwitchWorkSheets
				Gui Varbar:Default
				Gui VarBar: +AlwaysOnTop -Caption +ToolWindow +owner +HwndGUIID
				WinSet, Transparent, 100, %GUIID%
				Gui, VarBar:color,DC734F, 97BA7F   
				; Gui, VarBar:color,016D07, 97BA7F   
				this.AddBoxes()
				iniread, note1, data.ini, Notes, note1
				Iniread, note2, data.ini, Notes, note2
				Iniread, note3, data.ini, Notes, note3
				Iniread, note4, data.ini, Notes, note4
			CoordMode, mouse, screen
			IfWinexist, NuGenesis LMS - \\Remote
					LMS.Orient()
					; WinGetPos, VarBar_X, VarBar_Y,Varbar_W,Varbar_H, NuGenesis LMS - \\Remote,
			; if ShowSampleID || !Coated
				; try 
				MidScreen:=A_ScreenWidth//2
  				TopScreen:=1 ;A_ScreenHeight-35
		If Winactive("NuGenesiy LMS - \\Remote")
			try Gui, VarBar:Show, h32 x%NuX% y%Nuy%  NoActivate, VarBar
		else
			Try Gui, VarBar:Show, h32 x%Varbar_X% y%TopScreen%  NoActivate, VarBar
		Catch 
			Gui, VarBar:Show, h32 x%MidScreen% y%TopScreen%  NoActivate, VarBar
		CoordMode, mouse, window
		ControlsetText, Static1, %Iteration%,VarBar
		OnMessage(0x0201, "WM_LBUTTONDOWN")
		OnMessage(0x203, "VariableBar_Relocate")
		WinSet, Transparent, 235, AHK_id %GUIID%
		return



			ProductVarBar:
			BatchVarBar:
			LotVarBar:
			SampleIDVarBar:
			NoteVarBar1:
			NoteVarBar2:
			NoteVarBar3:
			NoteVarBar4:
			CoatedVarBar:
			sleep 100
			Gui, VarBar:submit,NoHide
			return

			VarBarGuiClose:
				coordmode, mouse, Screen
				WinGetPos,VarBar_X,Varbar_Y,w,h
				sleep 100
				IniWrite, %note1%, data.ini, Notes, note1
				IniWrite, %note2%, data.ini, Notes, note2
				IniWrite, %note3%, data.ini, Notes, note3
				; IniWrite, %Iteration%, data.ini, SavedVariables, Iteration
				IniWrite, %VarBar_X%, data.ini, Locations, VarBar_X
				; IniWrite, %Iteration%, data.ini, SavedVariables, yteration
				IniWrite, %VarBar_y%, data.ini, Locations, VarBar_Y
				; IniWrite, %Follow%, data.ini, Locations, Follow
				this.exit()
				coordmode, mouse, Window
				sleep 500
			return
	}


	AddBoxes(){
			global
		GUI,VarBar:Font,				s16 cBlack Bold, Consolas
		Gui,VarBar:Add,edit,		vProduct 	gproductVarBar 		 left h29 x20 y0  w65 ,		    %Product%  ; edit1
		GUI,VarBar:Font,				s11 cBlack,Consolas
		Gui,VarBar:add,Edit,		vBatch 		gbatchVarbar 		    left H29 x+1 y1 w75, 			    %Batch% 	 ; edit2
		GUI,VarBar:Font,				s11 cBlack , Consolas
		Gui,VarBar:add,Edit,		vlot 			gLotVarbar 		        x+1 left H29 y1 w65, 		    %Lot% 		 ; edit3
		GUI,VarBar:Font,				s9 cBlack,arial Narrow
			if !Coated
		Gui,VarBar:add,Edit,		vCoated 		gCoatedVarbar 		    left x+1 H29 y1 w25, 			  %Coated%   ; edit4
			else
		Gui,VarBar:add,Edit,		vCoated 		gCoatedVarbar 		    center x+1 H29 y1 w65, 			%Coated%   ; edit4
			if ShowSampleID
		Gui,VarBar:add,Edit,		vSampleID 	gSampleIDVarbar 		  x+1 H29 y1 w85, 		       %SampleID%  ; edit5
			else
		Gui,VarBar:add,Edit,		vSampleID 	gSampleIDVarbar 		  x+1 H29 y1 w0, 		        %SampleID%  ; edit5
		GUI,VarBar:Font,				s20 107C41, Consolas
		Gui,VarBar:Add,text,		vIteration								  x+5 65 center y-3 w23,		%Iteration%	; Text1
		GUI,VarBar:Font,				s9 cBlack,arial Narrow
		Gui,VarBar:add,Edit,		vNote1 		gNotevarbar1 		     x+3 H29 y1 left, 		 %Note1%     ; edit6
		Gui,VarBar:add,Edit,		vNote2 		gNotevarbar2 		     X+2 H29 y1 left,			  %Note2%  	; edit7
		Gui,VarBar:add,Edit,		vNote3 		gNotevarbar3 		     X+2 H29 y1 left,			  %Note3%  	; edit8
		Gui,VarBar:add,Edit,		vNote4 		gNotevarbar4 		     X+2 H29 y1 left,			  %Note4%  	; edit9
		}


			
		


		
		
		
		
		
	Focus(Control){
		global
		WinGetTitle, the_WinTitle, A
		caret_x:=A_CaretX
		caret_y:=A_Carety
		WinActivate, VarBar ahk_exe AutoHotkey.exe
		GuiControl Varbar:Focus, %Control%
		; sleep 100
		SendInput, ^{a}{ctrlup}{altup}
		return
		}	


	Reset(){
		Global
		; GUI, VarBar:destroy
			; if !xpos 
			; IniWrite, %xpos%, data.ini, Locations, VarBar_X
			; if !ypos 
			tt("Place bar")
		if !winactive("NuGenesis = \\Remote"){
			coordmode, mouse, Screen
			; keywait, F13, U T2
			MouseGetPos,Varbar_X,Varbar_Y
			Send, {laltup}
		} 
		else {
			WinGetPos, LMS_X, LMS_Y, LMS_W,LMS_H, NuGenesis LMS - \\Remote
			xpos := LMS_X+1000
			ypos := LMS_Y
		}
		tooltip,
		IniWrite, %ypos%, data.ini, Locations, VarBar_Y
		IniWrite, %xpos%, data.ini, Locations, VarBar_X
		IniWrite, %Xpos%, %A_ScriptDir%\ClipData\ClipChain\ClipChain.ini, Settings, ClipChainX
		IniWrite, %Ypos%, %A_ScriptDir%\ClipData\ClipChain\ClipChain.ini, Settings, ClipChainY
		; Gui, VarBar:Show, h30 x%xpos% y%ypos% w390 NoActivate
		; IniWrite, 0, data.ini, Locations, Follow
		VarBar.show(0)
		coordmode, mouse, window
		; Varbar.Relocate()
		return
		}



	Follow(){
		global
		if !winexist("ahk_exe WFICA32.EXE")
			return
		if winactive("VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe")
			return
		WinGet, NewWindow, ProcessName, A
		if (NewWindow=CurrentWindow) 
			return
			CurrentWindow:=NewWindow
		if winactive("ahk_exe EXCEL.EXE") { 
			WinGetPos, VarBar_oX, VarBar_oY, Varbar_oW,Varbar_oH, A
					WinMove, VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe, ,varbar_ox+100, Varbar_oy +2 ;+ Varbar_oh-33
			return
		}
		if winactive("ahk_exe Code.exe") || !winexist("ahk_exe WFICA32.EXE"){ 
			WinGetPos, VarBar_oX, VarBar_oY, Varbar_oW,Varbar_oH, A
					WinMove, VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe, ,varbar_ox+400, Varbar_oy + Varbar_oh-35
			return
		}
		if winactive("Inbox ahk_exe OUTLOOK.EXE") { 
			WinGetPos, VarBar_oX, VarBar_oY, Varbar_oW,Varbar_oH, A
					WinMove, VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe, ,varbar_ox+400, Varbar_oy + Varbar_oh-30
			return
		}
		if winactive("ahk_exe explorer.exe") { 
			; WinGetPos, VarBar_X, VarBar_Y, Varbar_W,Varbar_x, 
			varbar_NuX:=A_ScreenWidth/2
			varbar_wy:=A_ScreenHeight
			WinMove, VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe, ,varbar_NuX, varbar_wy
			return
		}
		if  winactive("ahk_exe WFICA32.EXE") && !winactive("ahk_exe EXCEL.EXE") && !winactive("ahk_exe Code.exe") && !winactive("ahk_exe OUTLOOK.EXE") { ;&& (ActiveWin="LMS"){
			lms.Orient()
			WinMove, VarBar ahk_class AutoHotkeyGUI ahk_exe AutoHotkey.exe, ,varbar_NuX, Varbar_wy
			return
		}
	}

	Relocate(){
		global
				PostMessage, 0xA1, 2
				keywait, Lbutton, U T2
					Send, ^a
		return
	}
		

	LaunchTable(){
		global
		try GUI ingredient_table:destroy
		try GUI Spec_table:destroy
		excel.connect(1)
		if winexist("Result Editor - \\Remote") || Winactive("Test Definition Editor - \\Remote") || winactive("Results Definition - \\Remote")
			SpecTab.Table()
		if winexist("Composition - \\Remote") || WinActive("Edit Ingredient - \\Remote")
			ProductTab.Table()
		else {
			ifwinnotactive, ahk_exe WFICA32.EXE 
			winactivate, ahk_exe WFICA32.EXE
		lms.DetectTab()
			If (Tab="Products")
				ProductTab.Table()
			else If (Tab="Specs")
				SpecTab.Table()
			else	
				Menu.Tables()
		}
		return
	}




		Search(input){
		global
		Clip()
		LMS.SearchBar(input)
		return
		}

		AddIteration(speed:=350){
		global Iteration
		; GuiControl, -redraw, varbar
		sleep 10
		Iteration+=1
		ControlsetText, Static1,%Iteration%,VarBar
		Pop(Iteration)
		; tt(Iteration,300,Varbar_x,Varbar_y,2,200)
		sleep %Speed%
		IniWrite, %Iteration%, data.ini, SavedVariables, Iteration
		; GuiControl, +redraw, varbar
		return
		}
		SubIteration(speed:=350){
		global Iteration
		; GuiControl, -redraw, varbar
		sleep 10
		Iteration-=1
		ControlsetText, Static1,%Iteration%,VarBar
		; tt(Iteration,300,Varbar_x,Varbar_y,2,200)
		Pop(Iteration)
		sleep %speed%
		IniWrite, %Iteration%, data.ini, SavedVariables, 
		; GuiControl, +redraw, varbar
		return
		}

	Load(){
		Global
		; Iniread, Product, data.ini, SavedVariables, Product
		Iniread, Batch, data.ini, Batches, Batch
		Iniread, Product, data.ini, Products, Product
		Iniread, Product0, data.ini, Products, Product0
		Iniread, Product1, data.ini, Products, Product1
		Iniread, Product2, data.ini, Products, Product2
		Iniread, Product3, data.ini, Products, Product3
		Iniread, Batch4, data.ini, Batches, Batch4
		Iniread, Batch1, data.ini, Batches, Batch1
		Iniread, Batch2, data.ini, Batches, Batch2
		Iniread, Batch3, data.ini, Batches, Batch3
		Iniread, Lot, data.ini, SavedVariables, Lot
		Iniread, Coated, data.ini, SavedVariables, Coated
		Iniread, SampleID, data.ini, SavedVariables, SampleID
		Iniread, Iteration, data.ini, SavedVariables, Iteration
		; iniread, note1, data.ini, Notes, note1
		; Iniread, note2, data.ini, Notes, note2
		; Iniread, note3, data.ini, Notes, note3
		; Iniread, note4, data.ini, Notes, note4
		Iniread, VarBar_Y, data.ini, Locations, VarBar_Y
		Iniread, VarBar_X, data.ini, Locations, VarBar_x
		}

	exit(){
		global
		wingetpos, Varbar_X, Varbar_Y,,, VarBar ahk_class AutoHotkeyGUI


		IniWrite, %Varbar_X%, data.ini, Locations, VarBar_X
		IniWrite, %Varbar_Y%, data.ini, Locations, VarBar_Y
		; wingetpos, Varbar_X, Varbar_Y,W,H, VarBar ahk_class AutoHotkeyGUI
		iniwrite, %Product%, data.ini, Products, Product
		iniwrite, %Batch%, data.ini, Batches, Batch
		; iniwrite, %Batch1%, data.ini, SavedVariables, Batch1
		; loop 3 {
			; n:=A_index
			; Item:=Product%n%
			; iniwrite, %Item%, data.ini, Products, Product%n%
			; Item:=Batch%n%
			; iniwrite, %Item%, data.ini, Batches, Batch%n%
			; }	
		; iniwrite, %Product0%, data.ini, Products, Product0
		; iniwrite, %Product1%, data.ini, Products, Product1
		; iniwrite, %Product2%, data.ini, Products, Product2
		; iniwrite, %Product3%, data.ini, Products, Product3
		; iniwrite, %Batch0%, data.ini, SavedVariables, Batch0
		iniwrite, %Lot%, data.ini, SavedVariables, Lot
		iniwrite, %Coated%, data.ini, SavedVariables, Coated
		iniwrite, %SampleID%, data.ini, SavedVariables, SampleID
		IniWrite, %Iteration%, data.ini, SavedVariables, Iteration
		IniWrite, %note1%, data.ini, Notes, note1
		IniWrite, %note2%, data.ini, Notes, note2
		IniWrite, %note3%, data.ini, Notes, note3
		IniWrite, %note4%, data.ini, Notes, note4
		; iniwrite, %VarBar_Y%, data.ini, Locations, VarBar_Y
		; iniwrite, %VarBar_X%, data.ini, Locations, VarBar_x
		; FileRead, OutputVar, Products.txt  
		; Sort, OutputVar, u
		; FileDelete, Products.txt
		; sleep, 300
		; FileAppend, %OutputVar%, Products.txt
		; Notes.Save()
		Gui,VarBar:Destroy
		}

HistoryMenuItem(){
	global
	sleep 200
	RegExMatch(A_ThisMenuItem, "i)[abdefghijkl]\d{3}\b", rProduct)
	ControlsetText, Edit1,%rProduct%, VarBar
	RegExMatch(A_ThisMenuItem, "i)(?<!Ct#)\b\d{3}-\d{4}\b", rBatch)
	ControlsetText, Edit2,%rBatch%, VarBar
	RegExMatch(A_ThisMenuItem, "i)(\b\d{4}\w\d\w?|\bBulk\b)", rlot)
	ControlsetText, Edit3,%rLot%, VarBar
	RegExMatch(A_ThisMenuItem, "i)(coated: |/?ct#/s|Ct#|ct/s|coated/s)\d{3}-\d{4}\b", rCoated)
	RegExMatch(rCoated,   "\d{3}-\d{4}", rCoated)
	ControlsetText, Edit4,%rCoated%, VarBar 
	Product:=rProduct
	Batch:=rBatch
	Lot:=rLot
	Coated:=rRoated 
	return
	}
	}
			
			
/* 			
	Class Notes{
		Show(){  ;array - remove duplicates (case insensitive)
			global
			try, GUI, Notes:destroy
			Iniread, Notes_X, data.ini, Locations, Notes_X
			Iniread, Notes_Y, data.ini, Locations, Notes_Y
			my_screenwidth:=Notes_x
			my_screenheight:=Notes_y
			MyArray:=[]
			FileRead, LoadedNotes, lib/Notes.txt
			MyArray := StrSplit(LoadedNotes,"`r`n")
			; loop % myarray.maxindex() {
				; LoadedNotes1:=MyArray[A_index]
			; }
			LoadedNotes1:=MyArray[1]
			LoadedNotes2:=MyArray[2]
			LoadedNotes3:=MyArray[3]
			; LoadedNotes4:=MyArray[4]
			; LoadedNotes5:=MyArray[5]
			; LoadedNotes8:=MyArray[8]
			; LoadedNotes6:=MyArray[6]
			; LoadedNotes7:=MyArray[7]
			gui Notes:+LastFound +AlwaysOnTop -Caption -ToolWindow +owner
			gui, Notes:add, button, Hidden default gNotesButtonOK, OK 
			gui, Notes:add, edit, y2 x2 w140 -Choose -VScroll +resize vMyEdit1, %LoadedNotes1%
			; gui, Notes:add, edit, w140 -Choose -VScroll +resize vMyedit2, %LoadedNotes2%
			; gui, Notes:add, edit, w140 -Choose -VScroll +resize vMyedit3, %LoadedNotes3%
			loop 2 {
				n:=A_index + 1
				Myedit=myedit%n%
				Note:=myArray[n]
				gui, Notes:add, edit, w140 -Choose -VScroll +resize v%Myedit%, % myarray[n]	
			}
			; OnMessage(0x84, "WM_NCHITTEST")
			; OnMessage(0x83, "WM_NCCALCSIZE")
			gui, Notes:color, 21a366
			; OnMessage(0x203, "Notes.Relocate")
			Notes_x:=Varbar_x+136
			Notes_y:=Varbar_Y+30
			gui, Notes:show, w145 x%Notes_x% y%Notes_y% ,Notes
			WinSet, Transparent, 195
			return
			
		}
		Add(){
			global
			; notes.close()
			; notes.show
			; n:=4
			Myedit:="vmyedit4"
			loadedNotes:="loadednotes4"
				gui, Notes:add, edit, w140 -Choose -VScroll +resize %Myedit%, %LoadedNotes%
			return
			
			
		}
		close(){
			global
				gui, Notes:submit, nohide
		Filedelete, lib/Notes.txt
			sleep 200
			; loop 4
			IniWrite, %Notes_X%, data.ini, Notes, Notes_X
			IniWrite, %Notes_y%, data.ini, Notes, Notes_Y
			Fileappend, %MyEdit1%`n, lib/Notes.txt
			Fileappend, %MyEdit2%`n, lib/Notes.txt
			Fileappend, %Myedit3%`n, lib/Notes.txt
			; Fileappend, %Myedit4%`n, lib/Notes.txt
			; Fileappend, %Myedit5%`n, lib/Notes.txt
			gui, Notes:destroy
		return
		} 
			Save(){
				global
			gui, Notes:submit, nohide
			Filedelete, lib/Notes.txt
			sleep 200
			Fileappend, %MyEdit1%`n, lib/Notes.txt
			Fileappend, %MyEdit2%`n, lib/Notes.txt
			Fileappend, %Myedit3%`n, lib/Notes.txt
			; Fileappend, %Myedit4%`n, lib/Notes.txt
			; Fileappend, %Myedit5%`n, lib/Notes.txt
			; Fileappend, %Myedit6%`n, lib/Notes.txt
			; Fileappend, %Myedit7%`n, lib/Notes.txt
			; Fileappend, %Myedit8%`n, lib/Notes.txt
			; Fileappend, %Myedit9%`n, lib/Notes.txt
			; Fileappend, %Myedit10%`n, lib/Notes.txt
			gui, Notes:destroy
			return
		}
		Relocate(){
				global
				; PostMessage, 0xA1, 2
				; keywait, Lbutton, U
				wingetpos, Notes_x, Notes_y,W,H, Notes ahk_class AutoHotkeyGUI
				; Excel.Connect()
				IniWrite, %Notes_x%, data.ini, Locations, Notes_x
				IniWrite, %Notes_y%, data.ini, Locations, Notes_y
				sleep 300
				return
			}
		WM_NCCALCSIZE()
		{
			if A_Gui
				return 0    ; Sizes the client area to fill the entire window.
		}
		}
		; ButtonAdd:
			; gui, Notes:submit
			; sleep 100
			; Note:=RegExReplace(Note "`n", "m`a)(?=^\s*;).*\R") ; remove commented lines
			; Note:=RegExReplace(Note, "\R+\R", "`r`n")     ; remove empty lines
		NotesButtonOK:
			notes.Save()
			notes.Close()
			return
			
			NotesGuiClose:
			NotesGuiEscape:
			notes.Save()
			notes.Close()
			; gui, Notes:submit, nohide
			; gui, Notes:destroy
			return \
			*/