#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

run, basics.ahk
run, defaultFTP.au3
run, changeview.au3
run, capsbackspace.au3
ExitApp