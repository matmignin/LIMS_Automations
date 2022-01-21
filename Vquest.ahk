﻿
VQuest_Start:
    #SingleInstance,Force
    #Persistent
  ;  #ErrorStdOut
    Process, Priority, , High
    #NoEnv
    Iniread, Iteration, Settings.ini, SavedVariables, Iteration
    ; Iniread, Mode, Settings.ini, Options, Mode

    iniread, DebuggingScript, Settings.ini, Options, DebuggingScript

    ; Iniread, HideVarbar, Settings.ini, Options, HideVarbar
    Iniread, ExcelConnect, Settings.ini, Options, ExcelConnect
    #KeyHistory 500
    #InstallKeybdHook
    #InstallMouseHook
    CheckTime:=600
    #ClipboardTimeout 1500
    ; #HotkeyInterval 50
    ; #MaxHotkeysPerInterval 500
    ; #MaxThreadsBuffer, On
    #InstallKeybdHook
    #InstallMouseHook
    #HotkeyModifierTimeout
    #maxthreadsperhotkey, 2
    SetTitleMatchMode, 2
    FormatTime, DayString,, MM/d/yy
    FormatTime, TimeString, R
    FormatTime, CurrentDateTime,, MM/dd/yy
    FormatTime, The_Day,, MMMM d
    FormatTime, The_Hour,, htt
    FormatTime, The_Time,, hh:mm
    SetNumLockState, on
    SetscrolllockState, off
    CrLf=`r`n
    EnvGet, VimOpen, VimOpen
    SetNumlockState Alwayson
    setcapslockstate alwaysoff
    SetscrolllockState, alwaysoff
    CoordMode, mouse, window
    SetWorkingDir, %A_ScriptDir%
    #winactivateForce
    AutoTrim, On
    ; else
    OnExit("Varbar.SaveVariables")
    ;FileName:="lib/winPos.txt"
    if !VarBar_x
      VarBar_x=1
    if !VarBar_y
      VarBar_y=1
    Menu, Tray, Add, CL3, CL3
    Menu, Tray, Add, DebugVars, DebugVars
    Menu, tray, NoStandard
    ; Menu, tray, Click,
    Menu, Tray, Add, KeyHistory, KeyHistory
    Menu, Tray, Add, listlines, listlines
    Menu, Tray, Add, windowSpy, windowSpy
    Menu, Tray, Add, Edit_Batches, Edit_Batches
    Menu, Tray, Add, Entering_Rotations, Entering_Rotations
    Menu, Tray, Add, TempCode, TempCode
    PasteTime:=A_TickCount
    if Mode=Edit_Batches
      Menu, Tray, Check, Edit_Batches
    else if Mode=Entering_Rotations
      Menu, Tray, Check, Entering_Rotations
    else if Mode=TempCode
      Menu, Tray, Check, TempCode
     ShowVarBar:=CreateMenu("showVarbar")
    ;  HideShowSampleID:=CreateMenu("ShowSampleID")
    ; ExcelConnect:=CreateMenu("")
    DebuggingScript:=CreateMenu("DebuggingScript")
    ; HideVarbar:=CreateMenu("HideVarbar")
    Menu, Tray, Add, E&xit, ExitSub
    Menu, Tray, Default, E&xit
    ; GuiControl, -redraw, varbar
    try Menu, Tray, Icon, bin\Robot.ico
    Currentwindow:=A
    varbar.Show()

    ifwinexist, Mats LMS Workbook.xlsb - Excel
      Excel.Connect(0)
    ; if !VimOpen
    IfwinExist, ahk_exe WFICA32.EXE
      LMS.Orient()
    copypasteToggle:=0
      ; run, lib\Vim.ahk
  if !ClipOpen
    try Run, cl3.Ahk, lib\CL3
    ; GuiControl, +redraw, varbar
    RegexProduct:="i)(?<=\w{3})?(?P<Product>[abcdefghijkl]\d{3}\b)"
    RegexBatch:=  "i)(?<!Ct#)(?P<Batch>\d{3}-\d{4}\b)"
    RegexLot:=    "i)(?P<Lot>\b\d{4}\w\d\w?|\bBulk\b|G\d{7}\w?\b|VC\d{6}[ABCDEFGH]?|V[A-Z]\d{5}[A-Z]\d?|\d{5}\[A-Z]{3}\d)"
    RegexCoated:= "i)(coated: |ct#?|ct\s?|coated\s?)(?P<Coated>\d{3}-\d{4})"


    ; if A_DebuggerName


  ;CurrentList:=[]
  ; Loop, Read, C:\Users\mmignin\Documents\VQuest\data\CurrentCodes.txt
    ; {
      ; CurrentList.Insert(A_loopreadline)
      ; maxindex := A_Index
    ; }

  OnClipboardChange("clipChange")
      #Include <SuppressErrorDialog>
  #include <Toggles>
  #Include <Temp>
  if !VimOpen
    #include <VIM>
  #Include <Test>
  #include <HotStrings>
  #include <KEYS>
  #include <PAD>
  #Include <LMS>
  #Include <clip>
  #Include <OpenApp>
  #include <Excel>
  #include <varBar>
  #include <menu>
  #include  C:\Users\mmignin\Documents\VQuest\lib\Vis\Vis2.ahk
  #include  <Functions>
  #include  <Xml>

  #Ifwinactive,
return

;__________________continuously runing sub_________________________________________________
activeCheck:
  If winactive("Result Entry - \\Remote") || winactive("Register new samples - \\Remote")
    varbar.FloatAtopwindow()
  else if winactive("Error - \\Remote") {
    ControlSend,, {enter}, Error - \\Remote
    sleep 200
    if winExist("Register new samples - \\Remote"){
      winactivate,
			Send, {click 180, 103,2}%Product%{enter}
		}
  }
  else if winactive("Information - \\Remote")
    send, {enter}
  else
    return
return


Entering_Rotations:
  Mode:="Entering_Rotations"
  Menu, Tray, Check, Entering_Rotations
  Menu, Tray, unCheck, Edit_Batches
  ; varbar.saveVariables()
  return
Edit_Batches:
  Mode:="Edit_Batches"
  Menu, Tray, unCheck, Entering_Rotations
  Menu, Tray, Check, Edit_Batches
  ; varbar.saveVariables()
  return
DebuggingScript:
  Mode:="DebuggingScript"
  Menu, Tray, unCheck, DebuggingScript
  Menu, Tray, Check, DebuggingScript
  ; varbar.saveVariables()
  return
TempCode:
  Mode:="TempCode"
  Menu, Tray, Check, TempCode
  ; varbar.saveVariables()
  return


