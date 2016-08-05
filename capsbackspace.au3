#include <WinAPISys.au3> ;get keyboard
$KBL_CUSTOM = "0xF0C00409" ;My Custom Keyboard Layout code
;$KBL_US 	= "0x04090409" ;US Keyboard Layout code
;$KBL_USINT	= "0xF0010409" ;US-international Keyboard Layout code
;$KBL_GREEK	= "0x04080408" ;Greek Keyboard Layout code

;Opt("SendAttachMode", 1)
Opt("SendCapslockMode", 0) 	;1 restores Caps to pre-send state (default!), 0 doesn't store
Send("{CapsLock off}")
HotKeySet("{CapsLock}","capsfun")

while 1
	Sleep(1000)
WEnd

;--- Functions
Func capsfun()
BlockInput($BI_DISABLE)
   $handl = WinGetHandle("")
   $ID = _WinAPI_GetKeyboardLayout ( $handl ) ;gets 32bit hex value of keyboard layout

   If $ID = $KBL_CUSTOM Then
	  HotKeySet("{CapsLock}") ;disable hotkey
	  Send("{BACKSPACE}{CapsLock off}") ;undo caps since it isn't captured by HotKeySet
	  HotKeySet("{CapsLock}","capsfun") ;reset hotkey
	  _DisableCaps() ;bugpatch
   Else ;no hotkey stuff
	  Send("{CapsLock toggle}") ;Since the capslock isn't captured by HotKeySet this shouldn't be necessary, but it is?!
   EndIf
BlockInput($BI_ENABLE)
EndFunc

Func _DisableCaps()
   $state = _GetCapsLockState()
   If $state Then
	  HotKeySet("{CapsLock}") ;disable hotkey
	  Send("{CapsLock off}")
	  HotKeySet("{CapsLock}","capsfun") ;reset hotkey
   EndIf
EndFunc

Func _GetCapsLockState()

    Local $Ret = DllCall('user32.dll', 'int', 'GetKeyState', 'int', 0x14)

    If @error Then
        Return SetError(1, 0, 0)
    EndIf
    Return BitAND($Ret[0], 1)
EndFunc   ;==>_GetCapsLockState

