
class Vim{
	
	find(){
  send, ~ 
  input, letter, L1
		    send, f
		send, %letter%{esc}
		send, {esc}
		return
  }
	
		
	Leader(LeaderKey){
  send, ~ 
  Input, Numbers, L3,,{enter}
			send, %leader%%Numbers%{enter}
	send, {esc}
  return
  }
}

    
















VIM:
  #If Getkeystate("Capslock","p") 
    #If (A_PriorHotKey = "d" AND A_TimeSincePriorHotkey < 4000)
    d::Send, {home 2}+{end}{Delete}
    
  #If (A_PriorHotKey = "d" AND A_TimeSincePriorHotkey < 4000)
    d::Send, {home 2}+{end}{Delete}
    w::sendinput, {right}^{left}+^{right}{backspace}
    4::sendinput, +{end}^x
    5::sendinput, ^m^x
    0::sendinput, +{home}^x
  #If (A_PriorHotKey = "y" AND A_TimeSincePriorHotkey < 400)
    d::Send, {home 2}+{end}^c
    w::sendinput, {right}^{left}+^{right}+{left}^c 
    5::sendinput, ^m^c
    
    4::sendinput, +{end}^c{right}
    y::
    Send {home}+s{end}^c
    sleep 150
    
    ; ClipWait, 1
    
    send, {right}
  
    return
 

    0::sendinput, +{home}^c
  #If (A_PriorHotKey = "g" AND A_TimeSincePriorHotkey < 500)
    g::Send, ^{home}
  #If (A_PriorHotKey = "p" AND A_TimeSincePriorHotkey < 500)
    p::Sendinput, {~}p{esc}
  #If Getkeystate("Capslock","p") 
	s::send, {altup}{CtrlUp}{ShiftUp}{LWinUp}
    p::send, {F21}
    1::F1
    2::F2
    3::F3
    `::esc
    j::down
    k::Up
    h::left
    l::right
    x::Delete
    g::send, ^{end}
    4::end
    0::home
    y::return
    d::return ; Send, {home 2}+{end}^x{delete}
    w::Send {right}^{right 2}{left}
    e::Send ^{right}
  ;  u::Send ^z
    up::+up
    down::+down
    right::+right
    left::+left
    b::^left
    Shift & o::sendinput, {home}{enter}{up}

    Capslock & o::Send, {end}{enter}
    o::Send, {end}{enter}
    +o::Send, {Home}{enter}
    Shift & ,::sendinput, !{,}
    left::sendinput, {home}+{Tab}
    right::sendinput, {home}{Tab}
    ; f::sendinput, {CtrlDown}{f}{Ctrlup}
    ; v up::
    ;   while GetKeyState("Capslock","p")
    ;     sendinput, {Shift down}
    ;   sleep 200
    ;   sendinput, {shift up}
    ;   retur

,::sendinput, {Altdown}{Ctrldown}{/}{CtrlUp}{AltUp}
  $a::sendinput, {ShiftDown}{altDown}{a}{shiftUp}{altup}
  Tab::sendinput, {CtrlDown}{i}{Ctrlup}
  f::Vim.find("f")
  [::
    sendraw,{
    send, {CtrlDown}{right}{Ctrlup}
    sendraw,}
    return
  ]::
    sendraw,}
    send, {CtrlDown}{left}{Ctrlup}
    sendraw,{
    return
;   3::+F3{
  $m::sendinput, {CtrlDown}{m}{Ctrlup}
  $u::sendinput, {CtrlDown}{u}{Ctrlup}
  ; u::+!up ;pselect revious instance of line
  $n::+!n ;select next instance of line
  up::sendinput, {Altdown}{shiftdown}{Ctrldown}{up}{shiftup}{CtrlUp}{AltUp}
  down::sendinput, {Altdown}{shiftdown}{Ctrldown}{down}{shiftup}{CtrlUp}{AltUp}
  `;::vim.Leader(":")
  i::sendinput, {CtrlDown}{i}{Ctrlup}
 5::sendinput, {shiftDown}{Ctrldown}/{CtrlUp}{{shiftup}
   r::^+!r
       Enter::sendinput, {shiftDown}{enter}{shiftup}
CapsLock::send, {esc}
    $<^k::sendinput, {Ctrldown}{up}{CtrlUp}
    $<^l::sendinput, {Ctrldown}]{CtrlUp}
    $<^h::sendinput, {Ctrldown}[{CtrlUp}
    $<^j::sendinput, {Ctrldown}{down}{CtrlUp}


   #if
 ; capslock::esc


#If WinActive("ahk_exe Code.exe") && Getkeystate("Capslock","p") ;editor



    







PsudoNumlock:
  #If Getkeystate("F19","p")
  0::sendinput, {CtrlDown}{0}{Ctrlup}
    ;sendlevel 1
    m::numpad1
    ,::numpad2
    .::numpad3
    j::numpad4
    k::numpad5
    l::numpad6
    u::numpad7
    h::numpaddiv
    i::numpad8
    o::numpad9
    `;::sendinput, {numpad0}
    n::Numpadsub
    /::,
    p::numpadmult
    '::numpadDot
    RShift::sendinput, {Tab 2}^a
    ~ENTER::sendinput, {enter}
    space::SendInput, 0
    ;sendlevel 0
      #if