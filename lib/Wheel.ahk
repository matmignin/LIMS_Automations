﻿Wheel_scroll(direction){
	ifwinactive, Result Editor - \\Remote
		mousemove, 505, 200,0	

	mouseclickDrag, L,0, 0, 0, %Direction%,0,R	; mouseclick, %upOrDown%, ,,1,0	0
	return
	}













Wheel_2(commands,Sleeptime:=600)
{
	global
	If (A_PriorKey!=A_ThisHotkey) || (A_TickCount-LastActivation > 200)
	{
	LastActivation:=A_TickCount
	send, %Commands%
	sleep %sleeptime%
	}
	return
}

Wheel(commands,Sleeptime:=200)
{
	BlockInput, On
	send, %Commands%
	BlockInput, Off
	sleep %sleeptime%
	return
}



Wheel_Right(){
	global
	If winactive("ahk_exe AHK-Studio.exe"){
		sendinput, !{right}	
	} Else If winactive("ahk_exe explorer.exe"){
		sendinput, !{right}	
} Else If winactive("ahk_exe EXCEL.EXE"){
		Wheel("{wheelleft}",0)	
	} Else If winactive("outlook"){
		Send, {wheelright} 
	} Else If winactive("Result Entry - \\Remote"){
		WorkTab_ChangeTestResults("toggle")	
	} else if Winactive("Select tests for request"){
		Mouse_Click("Select_TestsForRequests")
		;LMS
	} Else If winactive("NuGenesis LMS - \\Remote"){
		WorkTab_Main_EditResults()
		;ProductTab
	} else if winactive("Edit Formulation - \\Remote"){
		return
	} else if winactive("Composition - \\Remote"){
		Mouse_Click("Edit")
		
	} else if winactive("Edit Ingredient - \\Remote"){
		ProductTab_DropDown_Ingredient()	
	} Else If winactive("Test Definition Editor - \\Remote "){
		SpecTab_TestDefinitionEditor(Description)
	} Else If winactive("Results Definition - \\Remote "){
		Mouse_Click("Add")	
	} Else If winactive("Paster - Snipaste ahk_exe Snipaste.exe"){
		sendinput, ^=^=
		;=============================================================================
	} Else if winactive("Edit sample (Field Configuration: I`, Physical) - \\Remote"){
		Excel.Connect()
		sendinput, {tab 2}{right}{click 277, 139}{tab 6}%Batch%
		WorkTab_ShipToSelect()
	} Else if winactive("Edit sample (Field Configuration: CT`, Physical) - \\Remote"){
		WorkTab_ShipToSelect()
	} Else if winactive("Edit sample (Field Configuration: I`, Analytical) - \\Remote"){
		WorkTab_ShipToSelect()
	} Else if winactive("Edit sample (Field Configuration: F`, Micro) - \\Remote"){
		WorkTab_ShipToSelect()
	} Else if winactive("Edit sample (Field Configuration: CT`, Retain) - \\Remote, "){
		WorkTab_ShipToSelect()
	} Else if winactive("Edit sample (Field Configuration: I`, Retain) - \\Remote"){
		WorkTab_ShipToSelect()
	} else
		Send, {wheelright} 
		return
	
}


Wheel_left(){
global
	If winactive("ahk_exe AHK-Studio.exe"){
		sendinput, !{left}	
	} Else If winactive("ahk_exe explorer.exe"){
		sendinput, !{left}	
	} Else If winactive("Result Entry - \\Remote"){
		WorkTab_ChangeTestResults()	
		;ProductTab
} Else If winactive("ahk_exe EXCEL.EXE"){
		Wheel("{wheelright}",0)	
	} Else If winactive("outlook"){
		Send, {wheelleft} 
	} else if Winactive("Select tests for request"){
		Mouse_Click("Select_TestsForRequests")
		;LMS
	} Else If winactive("NuGenesis LMS - \\Remote"){
		Mouse_click("Main_search")
	} else if winactive("Edit Formulation - \\Remote"){
		return
	} else if winactive("Composition - \\Remote"){
		Mouse_Click("Add")
	} else 
		Send, {wheelleft}
		return 	 

}



Wheel_Paste() 
{
	Global
	sleep 100
	if winactive("ahk_exe WFICA32.EXE")
	{
		Clipboard := Trim((Clipboard, "`r`n"))
		clipwait, 1
		BlockInput, on
		send, %Clipboard%
		tooltip("Paste")
		blockinput off
	}
	else 
	sendinput, ^v
	ToolTip("Paste")
	sleep 1000
	return
}
Wheel_Cut() 
{
	global
	Send, ^x
	clipwait, 0.25
	ToolTip(clipboard)
	sleep 800
	
	return
}
Wheel_Copy() 
{
	global
	Send, ^c
	clipwait, 0.25
	ToolTip(clipboard)
	sleep 800
	return
}