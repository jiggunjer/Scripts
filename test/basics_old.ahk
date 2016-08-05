#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
notepad_location := "C:\Users\Jurgen\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\notepad++"

#n::
   Run, %notepad_location%
Return

#q::
	FileRecycleEmpty
return

#g:: 
    Run % "http://www.google.com/search?q=" . GetSelectedText()
return

;--- Functions
GetSelectedText() {
	tmp = %ClipboardAll% ; save clipboard
	Clipboard := "" ; clear clipboard
	Send, ^c ; simulate Ctrl+C (=selection in clipboard)
	ClipWait, 1 ; wait until clipboard contains data
	selection = %Clipboard% ; save the content of the clipboard
	Clipboard = %tmp% ; restore old content of the clipboard
	return (selection = "" ? Clipboard : selection)
}