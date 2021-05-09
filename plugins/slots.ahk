/*

Plugin            : Slots
Purpose           : Load & Save 10 quick paste texts
Version           : 1.4

10 Slots
Hotkeys: RCTRL-[1-0] 

History:
- 1.4 Attempt to prevent XMLRoot error - https://github.com/hi5/CL3/issues/15
- 1.3 Restore current clipboard from History
- 1.2 Added hotkey for QEDL() Ctrl+E (not public)
- 1.1 Bug fix for not correctly updatin control (Edit0 vs Slot0) and moved XML to ClipData, improved firsttime init

*/

SlotsInit:

If !IsObject(Slots)
	{
	 IfExist, %A_ScriptDir%\ClipData\Slots\Slots.xml
		{
		 If (XA_Load(A_ScriptDir "\ClipData\Slots\Slots.xml") = 1) ; the name of the variable containing the array is returned OR the value 1 in case of error
			{
			 MsgBox, 16, Slots, Slots.xml seems to be corrupt, starting a new Slots.xml
			 FileDelete, %A_ScriptDir%\ClipData\Slots\Slots.xml
			 Slots:=[]
			}
		}
	 else
		{
		 Slots:=[]
		 Loop, 10
			Slots[A_Index-1]:="Slot" A_Index-1 "a"
		}
	}

x:=10
y:=10
Index:=0

Gui, Slots:font,% dpi("s8")
Loop, 10
	{
	 Index++
	 If (Index = 10)
		Index:=0
	 Gui, Slots:Add, Text, % dpi("x" x " y" y),Slot #%Index% [RCtrl + %Index%]
	 Gui, Slots:Add, Edit, % dpi("w290 h60 vSlot" Index), % Slots[Index]
	 y+=80
	 if (A_Index = 5)
		y:=10
	 if (A_Index = 5)
		x:=310
	}
Gui, Slots:Add, Button, % dpi("x10 gSlotsSave"), &Save Slots (slots.xml)
Gui, Slots:Add, Button, % dpi("xp130 gSlotsSaveAs"), Save &As (name.xml)
Gui, Slots:Add, Button, % dpi("xp130 gLoadSlots"), &Load (name.xml)
Gui, Slots:Add, Button, % dpi("xp253 gSlotsClose"), &Close window
Return

;^#F12::
hk_slots:
If !WinExist("CL3Slots ahk_class AutoHotkeyGUI")
	Gui, Slots:Show, ,CL3Slots
else
	Gui, Slots:Hide
Return

;>^1::
;>^2::
;>^3::
;>^4::
;>^5::
;>^6::
;>^7::
;>^8::
;>^9::
;>^0::
hk_slotpaste:
OnClipboardChange("FuncOnClipboardChange", 0)
Clipboard:=Slots[SubStr(A_thisHotkey,0)]
PasteIt()
Sleep 100
Clipboard:=History[1].text
OnClipboardChange("FuncOnClipboardChange", 1)
stats.slots++
Return

~Esc::
SlotsGuiClose:
SlotsClose:
Gui, Slots:Cancel
Return

SlotsSave:
Gui, Slots:Submit, Hide
XMLSave("Slots","-" A_Now)
Index:=0
Loop, 10
	{
	 Slots[Index]:=Slot%Index%
	 Index++
	}
XMLSave("Slots")
Return

SlotsSaveAs:
SaveAsName:=""
Gui, Slots:Submit, Hide
InputBox, SaveAsName, Name for XML, Save slots as
If (SaveAsName = "")
	{
	 MsgBox, Enter filename!`nSlots not saved.
	 Gui, Slots:Show
	 Return
	}
XMLSave("Slots","-" A_Now)
Index:=0
Loop, 10
	{
	 Slots[Index]:=Slot%Index%
	 Index++
	}
StringReplace, SaveAsName, SaveAsName, .xml,,All
XA_Save("Slots", A_ScriptDir "\ClipData\Slots\" SaveAsName ".xml") ; put variable name in quotes
Return

LoadSlots:
Menu, SlotsMenu, Add
Menu, SlotsMenu, Delete
Menu, SlotsMenu, Add, Slots.xml, MenuHandlerSlots
Menu, SlotsMenu, Add
Loop, %A_ScriptDir%\ClipData\Slots\*.xml
	{
	 If (A_LoopFileName = "slots.xml")
		Continue
	 Menu, SlotsMenu, Add, %A_LoopFileName%, MenuHandlerSlots
	}
Menu, SlotsMenu, Show
Return

MenuHandlerSlots:
XMLSave("Slots","-" A_Now)
Slots:=[]
If (XA_Load(A_ScriptDir "\ClipData\Slots\" A_ThisMenuItem) = 1) ; the name of the variable containing the array is returned OR the value 1 in case of error
	{
	 MsgBox, 16, Slots, %A_ThisMenuItem% seems to be corrupt, starting a new Slots file
	 FileDelete, %A_ScriptDir%\ClipData\Slots\%A_ThisMenuItem%
	 Slots:=[]
	 Loop, 10
		Slots[Index-1]:="Slot" A_Index-1 "a"
	}
Index:=0	
Loop, 10
	{
	 GuiControl,Slots:, Slot%Index%, % Slots[Index]
	 Index++
	}
Return

; not public
#include *i %A_ScriptDir%\plugins\MyQEDLG-Slots.ahk
