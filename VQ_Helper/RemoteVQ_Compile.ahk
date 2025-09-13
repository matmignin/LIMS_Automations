DetectHiddenWindows On
SetTitleMatchMode,2
;SetBatchLines, -1
Process, Priority,, High
WinClose VQ_Helper.ahk - AutoHotkey
Sleep 600
SettingsFile:=A_ScriptDir "\Settings.ini"
; Read the version number from the text file
; FileReadLine, versionNumber, %VersionFile%, 1
FormatTime, The_Time,, h:mm:ss


; Update the version number in the text file
iniwrite, %The_Time%, %SettingsFile%, Config, CompileTime


; Display the new version number
; MsgBox, New Version: %newVersionNumber%

compileagain:
Try
{
RunWait  "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\RawFiles\AHK\Compiler\Ahk2Exe.exe"
 /in  "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\RawFiles\VQ_Helper.ahk"
 /out  "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\VQ_Helper.exe"
 /icon  "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\icon.ico"
 /base  "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\RawFiles\AHK\Compiler\Unicode 64-bit.bin"
 /compress 0
}
Catch e
{
	msgbox,, Compile didnt work `n %e%
	WinClose VQ_Helper.ahk
clipboard:="<<QuIT>>"
sleep 500
	goto, compileagain
}


sleep 3000

try
{
   run A_ScriptDir "\VQ_Helper.exe"
}
catch e
{
	clipboard:="<<QuIT>>"
	;msgbox, , didnt work `n %e%
	sleep 4000
	Run A_ScriptDir "\VQ_Helper.exe"
}
return



