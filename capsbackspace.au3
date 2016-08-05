#include <WinAPISys.au3> ;get keyboard
$KBL_CUSTOM  = "0xF0C00409" ;My Custom Keyboard Layout code
$KBL_CUSTOM2 = "0xF0C10409" ;My Custom Keyboard Layout code
;$KBL_US 	= "0x04090409" ;US Keyboard Layout code
;$KBL_USINT	= "0xF0010409" ;US-international Keyboard Layout code
;$KBL_GREEK	= "0x04080408" ;Greek Keyboard Layout code

Opt("SendCapslockMode", 0)
Send("{CapsLock off}")
HotKeySet("{CapsLock}","capsfun") ;Seems to capture capslock, in contrast to what documentation says...

while 1
	Sleep(1000)
WEnd

;--- Functions
Func capsfun()   
   Local $ID = _WinAPI_GetKeyboardLayout ( WinGetHandle("") ) ;gets 32bit hex value of keyboard layout 
   If ($ID = $KBL_CUSTOM) Or ($ID = $KBL_CUSTOM2) Then
     Send("{BACKSPACE}")     
   Else
     HotKeySet("{CAPSLOCK}")
     Send("{CAPSLOCK}")
     HotKeySet("{CAPSLOCK}","capsfun")
   EndIf
EndFunc

Func _GetCapsLockState() ;returns bool

    Local $Ret = DllCall('user32.dll', 'int', 'GetKeyState', 'int', 0x14)

    If @error Then
        Return SetError(1, 0, 0)
    EndIf
    Return BitAND($Ret[0], 1)
EndFunc   ;==>_GetCapsLockState

