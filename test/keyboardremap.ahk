#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
 
; bind to Shift+Space, change this according to your taste, refer to http://ahkscript.org/docs/Tutorial.htm for more info
+Space::

;--- Layouts
;you can also check HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Keyboard Layouts\ to see the codes for different layouts

;US_CUSTOM := 0xF0C00409
KeyboardLayout_QWERTY := 0x04090409
MsgBox % "The active window's ID is " . WinExist("A")
;--- Main
handl := WinExist("A")
currentKeyboardLayout := dec2hex(getKeyboardLayout(handl))

msgBox, % "test " . currentKeyboardLayout
MsgBox % "The active window's ID is " . WinExist("A")
if (currentKeyboardLayout = KeyboardLayout_QWERTY)
	Send {R}
else
    Send {F}
return 

;--- Functions 
getKeyboardLayout(h)
{
    idThread := GetThreadOfWindow(h) ;gets active window--WinExist("A")--if hwind not given
    keyboardLayout := DllCall("GetKeyboardLayout", "uint", idThread, "uint")
    MsgBox % "The active window's ID is " . WinExist("A")
    return keyboardLayout
}

GetProcessThreadOrList( processID, byRef list="" )
{
	;THREADENTRY32 {
	THREADENTRY32_dwSize=0 ; DWORD
	THREADENTRY32_cntUsage = 4	;DWORD
	THREADENTRY32_th32ThreadID = 8	;DWORD
	THREADENTRY32_th32OwnerProcessID = 12	;DWORD
	THREADENTRY32_tpBasePri = 16	;LONG
	THREADENTRY32_tpDeltaPri = 20	;LONG
	THREADENTRY32_dwFlags = 24	;DWORD
	THREADENTRY32_SIZEOF = 28

	TH32CS_SNAPTHREAD=4

	hProcessSnap := DllCall("CreateToolhelp32Snapshot","uint",TH32CS_SNAPTHREAD, "uint",0)
	ifEqual,hProcessSnap,-1, return

	VarSetCapacity( thE, THREADENTRY32_SIZEOF, 0 )
	NumPut( THREADENTRY32_SIZEOF, thE )

	ret=-1

	if( DllCall("Thread32First","uint",hProcessSnap, "uint",&thE ))
		loop
		{
			if( NumGet( thE ) >= THREADENTRY32_th32OwnerProcessID + 4)
				if( NumGet( thE, THREADENTRY32_th32OwnerProcessID ) = processID )
				{	th := NumGet( thE, THREADENTRY32_th32ThreadID )
					IfEqual,ret,-1
						ret:=th
					list .=  th "|"
				}
			NumPut( THREADENTRY32_SIZEOF, thE )
			if( DllCall("Thread32Next","uint",hProcessSnap, "uint",&thE )=0)
				break
		}

	DllCall("CloseHandle","uint",hProcessSnap)
	StringTrimRight,list,list,1
	return ret
}
GetThreadOfWindow( hWnd=0 )
{
	IfEqual,hWnd,0
		hWnd:=WinExist("A")
	DllCall("GetWindowThreadProcessId", "uint",hWnd, "uintp",id)
	GetProcessThreadOrList(  id, threads )
	IfEqual,threads,
		return 0
	CB:=RegisterCallback("GetThreadOfWindowCallBack","Fast")
	lRet=0
	lParam:=hWnd
	loop,parse,threads,|
	{	NumPut( hWnd, lParam )
		DllCall("EnumThreadWindows", "uint",A_LoopField, "uint",CB, "uint",&lParam)
		if( NumGet( lParam )=true )
		{	lRet:=A_LoopField
			break
		}
	}
	DllCall("GlobalFree", "uint", CB)
	return lRet
}

GetThreadOfWindowCallBack( hWnd, lParam )
{
	IfNotEqual,hWnd,% NumGet( 0+lParam )
		return true
	NumPut( true, 0+lParam )
	return 0
}

dec2hex(decValue)
{
    if (decValue < 0)
    {
        ErrorLevel := 1
        return 0
    }

    hexValue := ""
    base := 16

    Loop
    {
        remainder := mod(decValue, base)
        decValue //= base

        if (remainder >= 0 && remainder <= 9)
        {
            hexValue := remainder . hexValue
        }
        else
        {
            hexValue := Chr(remainder - 10 + Asc("A")) . hexValue
        }

        if (decValue = 0)
        {
            ErrorLevol := 0
            return "0x" . hexValue
        }
    }
}