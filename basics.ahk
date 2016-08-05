#NoEnv                       ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input              ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode RegEx
;---Shortcuts outside the scripts folder:
notepad_location := "C:\Users\Jurgen\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\notepad++"
return

;--- Map WINDOWS+C to open admin CMD from Windows Explorer {SECTION 1}
#IfWinActive ahk_class ExploreWClass|CabinetWClass  
    #c::OpenCmdInCurrent()
#IfWinActive

;--- Map PRINTSCR to snipping tool {SECTION 2}
PrintScreen::LaunchSnippingTool()

;--- Map SCROLL LOCK to calculator {SECTION 3}
Scrolllock::
    Run, scripts\calc.lnk
    
    IfWinExist, emulator  
        WinActivate, emulator
    else ;try calc
    {
        Sleep, 100
        IfWinExist, calc | emulator
            WinActivate, calc
        else
            MsgBox, Can not find a window with appropriate name (calc,emulator)
    }    
return

;--- Map Win + Z to zip/unzip {SECTION 4}
; unzips all selected archives to their own folder OR zips all selected files a single .7z archive
; Buggy if used with symlinks/shortcuts.
#z::
    zipunzip_selection()
return

;--- Map Win + X to ConEmu64 {SECTION 5}
#x::
    Run, scripts\ConEmu64.lnk
return

;--- Map Win + G to quick google text {SECTION 6}
#g:: 
    Run % "http://www.google.com/search?q=" . GetSelectedText()
return

;--- Map Win + N to open NPP {SECTION 7}
#n::
   Run, %notepad_location%
Return

;--- Map Win + Q to empty recycle bin {SECTION 8}
#q::
	FileRecycleEmpty
return

;====--- FUNCTIONS ---====

;--- Multipurpose {SECTION 0}
GetSelectedText() {
	tmp = %ClipboardAll% ; save clipboard
	Clipboard := "" ; clear clipboard
	Send, ^c ; simulate Ctrl+C (=selection in clipboard)
	ClipWait, 1 ; wait until clipboard contains data
	selection = %Clipboard% ; save the content of the clipboard
	Clipboard = %tmp% ; restore old content of the clipboard
	return (selection = "" ? Clipboard : selection)
}


;---  {SECTION 1}
; Opens the command shell 'cmd' in the directory browsed in Explorer.
; Note: expecting to be run when the active window is Explorer.
OpenCmdInCurrent()
{
    ; This is required to get the full path of the file from the address bar
    WinGetText, full_path, A

    ; Split on newline (`n)
    StringSplit, word_array, full_path, `n

    ; Find and take the element from the array that contains address
    Loop, %word_array0%
    {
        IfInString, word_array%A_Index%, Address
        {
            full_path := word_array%A_Index%
            break
        }
    }  

    ; strip to bare address
    full_path := RegExReplace(full_path, "^Address: ", "")

    ; Just in case - remove all carriage returns (`r)
    StringReplace, full_path, full_path, `r, , all


    IfInString full_path, \
    {
        try
		{
		Run,  scripts\elecmd.lnk "/K cd /D %full_path%"		
		}
		catch e
		{
        Run, cmd "/K cd /D C:\"
		exit
		}
    }
    else
    {
        MsgBox, Path not found
    }
}

;---  {SECTION 2}
; Determines if we are running a 32 bit program (autohotkey) on 64 bit Windows
IsWow64Process()
{
   hProcess := DllCall("kernel32\GetCurrentProcess")
   ret := DllCall("kernel32\IsWow64Process", "UInt", hProcess, "UInt *", bIsWOW64)
   return ret & bIsWOW64
}

; Launch snipping tool using correct path based on 64 bit or 32 bit Windows
LaunchSnippingTool()
{
    if(IsWow64Process())
    {
        Run, %A_WinDir%\Sysnative\SnippingTool.exe
    }
    else
    {
        Run, %A_WinDir%\System32\SnippingTool.exe
    }
}


;---  {SECTION 4}
zipunzip_selection()
{
    ;ListLines
    send, {# Up} ;had problem with windows start menu showing up during hotkey press :S
    Suspend On

    ; --- GLOBAL boolean switches
    global zip := true ;defines zip mode
    global multi := false ;ignored in zip mode
    global hasfile := false ;ignored in unzip mode

    ; --- check what should be done ()
    temp = %clipboard%
    clipboard =
    Send {Ctrl Down}c{Ctrl Up}
    ClipWait
    checkselection() ;sets zip,multi and hasfile
    sleep 100
    clipboard = %temp%

    ; --- use appropriate command to extract
    if (zip = true)
    {
        if (hasfile = true)
            filezip()
        else
            folderzip()
    }
    else ;unzip
    {
        if (multi = true)
            multizip()
        else
            singlezip()
    }

    Suspend Off
}



checkselection()
{
global zip
global hasfile
global multi
	loop, parse, clipboard, `n,`r
	{	
		Attributes := FileExist(A_LoopField) ;FileGetAttrib acted weird, returning empty strings sometimes	
		If !(InStr(Attributes,"D"))
		{	
			SplitPath,A_LoopField,,,ext
			if ext in zip,rar,7z,iso,tar,gz,lz ;unzip mode if extension is found
			{
				zip := false 					
				if (A_Index > 1)
				{
					multi := true 
					break ; exit loop, all relevant parameters are known
				}
			}
			else ;enter zip mode if normal file found
			{
				zip := true 
				hasfile := true
				break ; exit loop, all relevant parameters are known
			}
		}
	else
		zip := true ;continue and leave hasfile false
	}
}
 
 ;the unzipping context menu had some problems being too slow on my system, so added delays
singlezip()
{
	send {Shift Down}{F10}{Shift Up}
	sleep 50
	send 7eee
	sleep 50
	send {enter}
	return
}
 
multizip()
{
	send {Shift Down}{F10}{Shift Up}
	sleep 50
	send 7ee
	sleep 50
	send {enter}
	return
}

filezip()
{
	send {Shift Down}{F10}{Shift Up}7aa{enter}
	return
}

folderzip()
{
	send {Shift Down}{F10}{Shift Up}7a{enter}
	return
}