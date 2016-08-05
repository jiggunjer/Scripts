#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer.au3" ;get explorer view


;Opt("SendAttachMode", 1)
Opt("SendCapslockMode", 0) 	;1 restores Caps to pre-send state (default!), 0 doesn't store
Send("{CapsLock off}")
HotKeySet("{CapsLock}","capsfun")

while 1
	Sleep(10000)
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


;---Bugs
;Sometimes the backspace-version actives caps lock with repeated presses. Using that mode should always result in caps lock off, and a matching keyboard light state.
;Related to above: keyboard light also toggles unpredicably. Light state doesn't correlate with CapsLock state

