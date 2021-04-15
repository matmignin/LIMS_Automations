﻿gosub, vquest_start


return




Test(n:=0){
  Global
  ToggleFilter_Test()
  FilterSearch_Test("Vitamin C","221")
  CopyResults_Test()
  ParseResultSpecs()
  FilterSearch_Test()
  return
}

CopyResults_Test(){
  WinActivate, NuGenesis LMS - \\Remote
    click 57, 1079 ; edit results
    WinWaitActive, Results Definition - \\Remote
    click 282, 121 ; click row
      sleep 100
  SendLevel, 1
  send, ^c
  clipwait, 1
  if ErrorLevel
    msgbox, yo
  sendlevel,0
  sleep 200
  send, {esc}
}

ParseResultSpecs(){
  global
  ParsedSpecs:=[]
  Loop, parse, Clipboard, `t 
    ParsedSpecs.insert(A_LoopField)
  LowerLimit:=Parsedspecs[17]
  UpperLimit:=Parsedspecs[18]
  Precision:=Parsedspecs[19]
  Requirement:=Parsedspecs[20]
  Unit:=Parsedspecs[21]
}

ToggleFilter_Test(){
  WinActivate, NuGenesis LMS - \\Remote
  click 489, 836, R
  sendinput, {down 2}{enter}
}

FilterSearch_Test(TestName:="", MethodName:=""){
  WinActivate, NuGenesis LMS - \\Remote
  click 837, 810 ;click name filterbox
  sendinput, %TestName%{enter}
  click 687, 809
  sendinput, %MethodName%{enter}
}

return








#IfWinActive,
Sendlevel 1
  F14 & F13::sendinput, {CtrlDown}{F21}{CtrlUp}
  F13 & F14::
IfWinnotExist, ahk_exe firefox.exe,
 send, #2 ; open firefox
WinActivate, ahk_exe firefox.exe 
return
  Rbutton & Mbutton::send, {F21}
  Rbutton & F17::F21
  F19::send, {F21}
  Sendlevel 0


KEY_DEFAULT:
Return & K::Enter_Product("K")
  Mbutton & WheelDown::Wheel("^{WheelDown}") 
  Mbutton & Wheelup::Wheel("^{WheelUp}") 
  Mbutton & F17::Wheel_Right()
  Mbutton & F16::Wheel_left()
  Mbutton::Mbutton
  Rbutton & Wheelup::Wheel_cut() 
  Rbutton & Wheeldown::Wheel_paste()
  Rbutton & F14::Get_WindowInfo()
  Rbutton & F16::Backspace
  Rbutton & Lbutton::Enter
  Rbutton up::Mouse_RbuttonUP()
  F13 & LButton::Sendinput, +^4 ;screenshot"
  F13 & RButton::Sendinput, +^3 ;screenshot"
  F13 & Mbutton:: varbar.set("OCR") 
  F13 & WheelUp::Varbar.Sendinput("Product")
  F13 & WheelDown::Varbar.Sendinput("Batch")
  ; F14 & Rbutton::Get_WindowInfo()
  F13::Clip_set()
  F13 & F18::Varbar.reset()
  F14 & wheelDown::Mouse_CloseWindow()
  F18 & Wheelup::!^tab
  F14 & Lbutton::Sendinput, #{down}
  
  F14::Menu()
  ; F18 & Wheelup::sendinput, {F15}
  F18 & Lbutton::sendinput, ^{Click}
  F18 & Rbutton::sendinput, +{Click}
  F18 & Mbutton::^a
  F18 & F17::sendinput, !^{tab}
  F18 & F16::sendinput, !^{+tab}
enter::enter

#If (A_PriorHotKey = "F18" AND A_TimeSincePriorHotkey < 2000) ;____F18____
  f18::ReloadScript()
  Wheelup::test()
  Wheeldown::send, {esc}
  F16::sendinput, {Click 354, 44}
  F17::sendinput, {click, 743, 41}
  Lbutton::sendinput, ^{click}
  Rbutton::sendinput, +{click}
  mbutton::VarBar.set()
  #If
  F18::Autofill() ;Tooltip("☩",2000) 

KEY_Varbar:
  #If Mouse_IsOver("VarBar ahk_class AutoHotkeyGUI") 
  WheelUp::Varbar.AddIteration()
  wheelleft::Excel.NextSheet()
  wheeldown::Varbar.SubIteration()
  F17::Excel.NextSheet()
  wheelRight::excel.previoussheet()
  Rbutton::Excel.connect()
  F16::excel.previoussheet()
  mbutton::
    if WinExist("Result Editor - \\Remote") || WinExist("Test Definition Editor - \\Remote") || winexist("Results Definition - \\Remote")
      SpecTab_Table()
    else
      ProductTab_Table()
  return
  #if
  
  
  
  
  
  
  
;___________________________________________________________________________
KEY_LMS:
  #Ifwinactive, NuGenesis LMS - \\Remote 
  F20::Mouse_Click("SearchBar_Batch") 
  F14 & WheelRight:: sendinput, {click, 743, 41}
  F14 & WheelLeft::sendinput, {Click 354, 44}
  F17::click 78, 750
  F14 & wheeldown::mouse_click("searchBar_Batch")
  F14 & wheelup::mouse_click("searchBar_Product")
    F13 & WheelUp::Varbar.Sendinput("Product")
  F13 & WheelDown::Varbar.Sendinput("Batch")
  #IfWinActive
  ;select methods
  #Ifwinactive, Select methods tests - \\Remote
  F19 & space::AutoFill()
#ifwinactive, Edit test (Field Configuration:
  F16::Autofill()
#Ifwinactive, Result Entry - \\Remote
  F13::WorkTab_ChangeTestResults("toggle")	
#ifwinactive, Register new samples - \\Remote
  F13 & wheelup::sendinput, {click 180, 103, 2}%Product%{enter}
  F18::Autofill(1)

#IfWinActive, ahk_exe WFICA32.EXE, ;GENERIC LMS
  F14 & WheelUp::Varbar.AddIteration()
  F14 & wheeldown::Varbar.SubIteration()
  browser_Forward::Excehl.NextSheet()
  browser_Back::Excel.PreviousSheet()
  F13 & F17::Excel.NextSheet()
  F13 & F16::Excel.PreviousSheet()
  Mbutton::AutoFill()
  F18::autofill()






KEY_Excel:
  #IfWinActive, LMS Workbook.xlsm - Excel
  F18::Excel.Connect(1)
  MButton::Excel.Connect(1)
  #ifwinactive, ahk_exe EXCEL.EXE
  +Enter::sendinput, !{enter}
  $Enter::sendinput, {enter}
  Mbutton::Excel.Search()
  F16::wheel("{wheelleft}",80)
  F17::wheel("{wheelRight}",80)
  #ifwinactive, Find and Replace,
  F13 & WheelUp::
  Wheel("{alt down}n{alt up}")
  Varbar.Sendinput("Product")
  return
  F13 & WheelDown::
  Wheel("{alt down}n{alt up}")
  Varbar.Sendinput("Batch")
  return
  return::sendinput, !i
  rbutton & Lbutton::sendinput, !i

#IfWinActive, ahk_exe explorer.EXE
Mbutton::Open_in_Notepad()

KEY_OUTLOOK:
  #IfWinActive, ahk_exe OUTLOOK.EXE 
  Mbutton::SENDINPUT % Varbar.Sendinput("Batch") " is updated"
  Rbutton & Wheelup::Wheel("{ctrl down}x{ctrl up}")
  Rbutton & wheeldown::Wheel("{ctrl down}v{ctrl up}")
  Rbutton & F17::Varbar.Sendinput("Batch"," is updated.")
  F14 & WheelDown::varbar.search("batch")
  F14 & Wheelup::varbar.search("Product")
  F13::varbar.set()
  F19 & Space::Varbar.Sendinput("Product")
  F20 & Space::Varbar.Sendinput("batch")
  F18::LMS_Search()

KEY_Browser:
  #IfWinActive, ahk_exe firefox.exe
  F14 & WheelDown::Wheel("^w")
  F16::sendinput, !{left}
  F13 & WheelDOWN::Browser_Back
  F17::!right
  F14 & wheelUP::Browser_Forward
  Mbutton::sendinput, ^{click}

KEY_Snipper: 
  #IfWinActive, Snipper - Snipaste ; the capture screen
  rbutton::send, ^c
  WheelDown::wheel("^1")
  Wheelup::Wheel("^5")
  #IfWinActive, Paster - Snipaste ahk_exe Snipaste.exe ; the floating window
  Mbutton & wheelUp::Wheel("{click right}z1{click right}e{ctrl down}5{ctrl up}")
  Mbutton::send, {esc}
  F13::sendinput, ^c
  F17::sendinput, ^+{+}
  F16::sendinput, ^+{-}
  #If mouse_isover("Paster - Snipaste ahk_class Qt5QWindowToolSaveBits")
  sendlevel 1
  Mbutton::send, !{click}
  #if

Scroll_Fix:
  #ifwinactive, Result Editor - \\Remote
  wheeldown::Wheel_scroll("100")
  wheelup::Wheel_scroll("-100")
  
ClipTool:
#IfWinactive, C:\Users\mmignin\Documents\VQuest\lib\CL3\cl3.ahk - AutoHotkey v1.1.33.06,
  F19::send, {esc}
  F19 & j::send, {down}
  F19 & k::send, {up}
  F19 & l::send, {enter}
  F19 & h::send, {esc}
  Capslock & j::send, {down}
  Capslock & k::send, {up}
  capslock & l::send, {enter}
  capslock & h::send, {esc}
  capslock::esc
  #IfWinActive

#If (A_PriorHotKey = "Browser_forward" AND A_TimeSincePriorHotkey < 1000)
  browser_forward::Send, !{tab}{space}
  browser_back::send, +!{tab}{space}
  media_Next::Mouse_CloseWindow()
  #if

KEY_Otherstuff:
  Media_Play_Pause::F18 ;4finger tap
  Media_Prev::F7 ;3fringer down
  ; Media_Next:: ;3finger up
  Browser_Back::+!tab ;3finger left
  Browser_Forward::^!tab

  Volume_Down::F11
  Volume_Up::F12
  #right::tooltip(A_ThisHotkey)
  ^+!down::tooltip(A_ThisHotkey)
  CapsLock::return
  F20 & Right::send, #{right}
  F20 & Left::send, #{Left}
  F20 & UP::send, #{UP}
  F20 & Down::send, #{Down}
  F20::AutoFill()
  
#If (A_PriorHotKey = "Media_next" AND A_TimeSincePriorHotkey < 1000)
    Browser_forward::#right
    Browser_back::#left
    Media_Next::sendinput, #{up}
    Wheelup::tooltip(A_ThisHotkey)
    Wheeldown::tooltip(A_ThisHotkey)
    wheelleft::tooltip(A_ThisHotkey)
    wheelright::tooltip(A_ThisHotkey)
    Lbutton::tooltip(A_ThisHotkey)
    Rbutton::tooltip(A_ThisHotkey)
    mbutton::tooltip(A_ThisHotkey)
    #If
  Media_next::Tooltip("Swipe",2000) 
  
#If WinActive("ahk_exe Code.exe") && Getkeystate("Capslock","p") ;editor
    ,::sendinput, +!/
    m::+!down
    n::+!n
    u::+!up
    '::+^!n
    .::+F1
    up::^+up
    down::^+down
    /::sendinput, !^w
    5::^+/

  Open_in_Notepad(){
    click
    WinGetClass class, % " ahk_id " WinExist("A")
    for Window in ComObjCreate("Shell.Application").Windows 
      Selection := Window.Document.SelectedItems
    for Items in Selection
      Path_to_Selection := Items.path		
    Run C:\Windows\system32\Notepad.exe %Path_to_Selection%
  }

  Snip_groupChange(){
    Wheel("{ctrl down}22{Ctrl up}")
    tooltip("Snip Group Changed",2000)
  return
}
Enter_Batch(key){
  global
  ;CoordMode, mouse, screen
  inputbox,Batch,, %key%-,,70,130,,,,%Batch%
  if ErrorLevel
  return
else
  Batch:=key . "-" . Batch
CoordMode, mouse, window
;Save_Code("Batches", Batch)
;Envset, Batch, %Batch%
Sleep 100		
VarBar.show()
return
}

Enter_Product(key){
  WinActivate, ahk_class Transparent Windows Client ahk_exe WFICA32.EXE
  inputbox, Code,, %key% `t ,,,,%varbar_x%,%Varbar_y%
  Sleep 100	
  click 755, 48
  sleep 200
  ;click 1131, 80
  click 577, 131, 2 ;click search bar
  sleep 100
  varbar.set(Key . Code)
  Sendinput, %Key%%Code%{enter}
  ;sendinput, {ctrl down}a{ctrl up}%Product%{enter}
  ; VarBar.show()
  return
} 

LMS_Search(){
  send, ^c
  ClipWait, 1,
  Clipboard := Trim((Clipboard, "`r`n"))
  WinActivate, NuGenesis LMS - \\Remote
  click 783, 45
  sleep 400
  click 487, 125, 2 ;click search bar
  ; click, 500,127, 2 ;click search bar
  sleep 200
  Send, %clipboard%{enter}
  return
}

StopSub: 
exitapp
Return
VarBar_ResetSub:
  VarBar.Reset()
return
Run_Display:
  run, Display.url, C:\Users\mmignin\Desktop\
Run_CL3:
  Run, cl3.Ahk, lib\CL3
return
run_VIM:
  Run, ViM.Ahk
return
#IfWinActive,
WindowSpySub: 
Run, WindowSpy.ahk,C:\Program Files\AutoHotkey\

#Include <AutoFill>
#include <varBar>
#include <ProductTab>
#include <WorkTab>
#include <SpecTab>
#include <Rotation>
#include <Excel>
#include <vis2>
#include <wheel>
#include <VScode>
#include <mouse>
VQuest_Start:
  #Persistent 
  #NoEnv
  #SingleInstance,Force
  #KeyHistory 100
  SetWorkingDir, %A_ScriptDir%
  Menu, Tray, Add,CL3, Run_cl3
  menu, tray, add, Display settings, Run_display
  Menu, Tray, Add, Vim, Run_Vim
  Menu, Tray, Add, Stop, StopSub
  Menu, Tray, Add, windowSpy, WindowSpySub 
  Menu, Tray, Add, ResetVarbar, Varbar_ResetSub
  Menu, Tray, Default, ResetVarbar
  Setnumlockstate Alwayson
  setCapslockstate alwaysoff
  SetscrolllockState, alwaysOff
  SetDefaultMouseSpeed, 0
  detecthiddenwindows, on
  SetTitleMatchMode, 2
  settitlematchmode, slow
  #MaxHotkeysPerInterval 200
  #HotkeyModifierTimeout 100
  #maxthreadsperhotkey, 1
  SetKeyDelay, -1, 0
  setwindelay, 450
  AutoTrim, On
  Menu, Tray, Icon, Robot.ico
  #HotkeyModifierTimeout 
  Results_Definition_edit:="78,63"
Result_Editor_Ok:="370,660"
Results_Definition_ok:="1336,592"
Requests_tab:="865,83"
MyWork_Tab:="761,44"
Mywork_Searchbar:="500,127,2"
  Iniread, Iteration, data.ini, SavedVariables, Iteration
  Iniread, VarBar_X, data.ini, Locations, VarBar_x
  Iniread, VarBar_Y, data.ini, Locations, VarBar_Y
  Excel.Connect(1)