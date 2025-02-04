	#SingleInstance,Force
	#NoEnv
  #NoTrayIcon

CoordMode, mouse, screen
Try GUI, EnterEachLineGUI:destroy


; Define the path to the text file
textFilePath := "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\TEXTtoENTER.txt"
ScreenEdge_X:=A_ScreenWidth-450
ScreenEdge_Y:=A_Screenheight-580
; Function to read the file and populate the list
FileReadLines(filePath)
{
    global lines
    lines := []
    Loop, Read, %filePath%
    {
        lines.Push(StrSplit(A_LoopReadLine, "`t"))
    }
}

; Call the function to read the file
FileReadLines(textFilePath)

		Try GUI, EnterEachLineGUI:destroy

; Create the GUI
    GUI, EnterEachLineGUI:Default
    Gui EnterEachLineGUI:+LastFound +Toolwindow +Owner +AlwaysOnTop -SysMenu +MinimizeBox
    GUI, EnterEachLineGUI:Font, s12  ;cBlack, Arial Narrow
    Gui, EnterEachLineGUI:Add, ListView, x10 y10 w370 h400 -ReadOnly Checked Grid vMyListView gEnterText, Details|Text
    Gui, EnterEachLineGUI:Add, Button, x10 y420 w180 gSubmitChecked, Type Checked Items
    Gui, EnterEachLineGUI:Add, Button, x320 y420 w50 gEditTextFile, Edit

For index, lineArray in lines
{
    if (Trim(LineArray[3]))
    LV_Add("Col0 Check",Trim(lineArray[1]), Trim(lineArray[2]))
    else
    LV_Add("Col0",Trim(lineArray[1]), Trim(lineArray[2]))
      ; LV_Modify(%A_Index%, "Check")
}
LV_ModifyCol()
LV_ModifyCol(1,140)
LV_ModifyCol(2,240)
OnMessage(0x0201, "WM_Lbuttondown")
Gui, EnterEachLineGUI:Show, x%ScreenEdge_X% y%ScreenEdge_Y% w385 h460, ListBox with Checkboxes
return
; Event for ListView item click



; #IfWinActive
EnterEachLineGUICloseGui:
EnterEachLineGUIGuiEscape:
EnterEachLineGUIGuiClose:
Gui, EnterEachLineGUI:Destroy
ExitApp
return


WM_Lbuttondown:
EnterText:
EnterTextMenuButton:
	if (A_GuiEvent = "NORMAL" ){
    ; LV_GetText(NewDetailsText, A_EventInfo, 1)
    ; LV_GetText(NewInputText, A_EventInfo, 2)
    ;   LV_Modify(A_EventInfo,,NewInputText,NewDetailsText)
; Gui, Submit,Nohide
    ; sendinput, {F2}
    ; sendinput, {space}
  ;   LV_Modify(A_EventInfo, "Check")
  ;   	; If !Row
	; 		; row:=A_EventInfo
		}
  If (A_GuiEvent = "E") ; If an edit event occurred
{
    LV_GetText(NewDetailsText, A_EventInfo, 1)
    LV_GetText(NewInputText, A_EventInfo, 2)
  ; msgbox, %oldDetailsText%  %oldInputText%  %A_EventInfo% `n %NewDetailsText%   %NewInputText%
  ; ; Gui, Submit,Nohide
  ;   ; LV_GetText(oldDetailsText, A_EventInfo, 1)
  ;   ; LV_GetText(oldInputText, A_EventInfo, 2)
  ;   ; If (newDetailsText != "") OR (newInputText != "")
  ;   ; {
      ; LV_Modify(A_EventInfo,,NewInputText,NewDetailsText)
  ;       LV_Modify(A_EventInfo, "", newDetailsText, newInputText)
        UpdateTextFile()
  ;   ; }
}
else If (A_GuiEvent = "DoubleClick") ; If a double-click event occurred
{
      LV_GetText(InputText, A_EventInfo, 1)
      LV_GetText(DetailsText, A_EventInfo,2)
			LV_Modify(A_EventInfo, "Check")
      ; LV_Modify(A_EventInfo,,InputText,DetailsText)
      ; LV_GetText(DetailsText, A_EventInfo, 1)
      ; LV_GetText(InputText, A_EventInfo, 2)
      if Winexist("Barcode Scanner")
            WinActivate, Barcode Scanner ahk_exe ELN.exe
      else if Winexist("ahk_exe EXCEL.EXE")
            WinActivate, ahk_class XLMAIN ahk_exe EXCEL.EXE
      else
        winactivate, ahk_exe eln.exe
        sleep 200
      ; msgbox % Trim(DetailsText)
      SendInput %InputText%
      sleep 400
      SendInput {Enter}
}
		else if A_ThisMenuItemPos
		{
      LV_GetText(InputText, A_EventInfo, 1)
      LV_GetText(DetailsText, A_EventInfo, 2)
			; LV_Modify(A_ThisMenuItemPos, "Check")
      msgbox % Trim(InputText)
		}
		return
EditTextFile:
; UpdateTextFile()
		Run , Edit "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\TEXTtoENTER.txt"
    return


SubmitChecked:
; Gui, Submit
LV_ModifyCol()
; tooltipText := ""
if Winexist("ahk_exe EXCEL.EXE")
      WinActivate, ahk_class XLMAIN ahk_exe EXCEL.EXE
else
  winactivate, ahk_exe eln.exe
Loop
{
    RowNumber := LV_GetNext(RowNumber, "Checked")
    If !RowNumber
        Break
    LV_GetText(InputText, RowNumber, 1)
    LV_GetText(DetailsText, RowNumber, 2)
    tooltipText .= InputText "`t " DetailsText "`n"
    ; msgbox % Trim(DetailsText)
    SendInput %InputText%
    Tooltip, %tooltipText%
    sleep 200
    SendInput {Enter}
    sleep 400
}
SetTimer, RemoveTooltip, 3000
Return

; Function to handle individual item click


Gui, EnterEachLineGUI:ListView, Text|Details
LV_ModifyCol()
ListView:
{
    If !A_EventInfo
        Return
    LV_GetText(InputText, A_EventInfo, 1)
    LV_GetText(DetailsText, A_EventInfo, 2)
    if Winexist("ahk_exe EXCEL.EXE")
      WinActivate, ahk_class XLMAIN ahk_exe EXCEL.EXE
    else
      winactivate, ahk_exe eln.exe
    sleep 200
    SendInput % Trim(InputText)
    sleep 200
    SendInput {Enter}
}
Return



UpdateTextFile()
{
    global
    lines := []
    LV_GetCount()
    ; msgbox % lv_getcount()
    Loop % LV_GetCount()
    {
        LV_GetText(DetailsText, A_Index, 2)
        LV_GetText(InputText, A_Index, 1)
        lines.Push([DetailsText, InputText])
    }
    File := FileOpen(textFilePath, "w")
    for index, lineArray in lines
    {
        File.WriteLine(Trim(trim(lineArray[2]) "`t" Trim(lineArray[1]) "`r"))
    }
    File.Close()
}

RemoveTooltip:
Tooltip
SetTimer, RemoveTooltip, Off
Return