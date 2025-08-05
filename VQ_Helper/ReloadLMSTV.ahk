
	#SingleInstance,Force
	#Persistent
	#NoEnv
	SetTitleMatchMode, 2
	CoordMode, mouse, Window
	#winactivateForce
    #InstallKeybdHook
	#InstallMouseHook
	IdleThreshold := 24 * 60 * 60 * 1000 ;24 hours


; Set timer to trigger every 5 minutes (300,000 milliseconds)
SetTimer, ClickReloadOrPosition, 300000



	Menu,Tray,NoStandard
	Menu, Tray, Add, Pause, Pausesub
	Menu, Tray, Add, Reload, ReloadSub
	Menu, Tray, Add, Exit, Exitsub
    Menu, Tray, tip, %A_TimeIdlePhysical%
	Menu, Tray, Default, Pause

ClickReloadOrPosition:

if (A_TimeIdlePhysical > IdleThreshold)
	ExitApp

    If WinExist("NuGenesis LMS")
    {
        WinActivate,
            WinGetPos, xWin, yWin, width, height, NuGenesis LMS
            xClick2 := width * 0.10
            ; xClick := width * 0.30
            ; xClick2 := xClick - 200
            yClick := 120
            Click, 271, %yClick%
            Click, 260, %yClick%
            Click, %xClick2%, %yClick%
        sleep 500
        MouseMove,10, 10,0
    }



return


#Ifwinactive, Login
Mbutton::Sendinput, ldisplay{tab}labfc1{enter}



#ifwinactive,


Reloadsub(){
	Critical, on
	Thread, Priority, 1
	reload
}

Exitsub(){
	global
	exitApp
}
Pausesub(){
	global
	Suspend, Toggle


}