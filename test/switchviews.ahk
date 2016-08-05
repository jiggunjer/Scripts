#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#IfWinActive, ahk_class CabinetWClass

; View ICONS
F1::
ControlGet, View, Hwnd,, SysListView321, A
parent := DllCall("GetParent","UInt", View) 
PostMessage, 0x111, 0x7029, 0, , ahk_id %parent% 
return