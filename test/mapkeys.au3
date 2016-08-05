#include <WinAPISys.au3> ;get keyboard
#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer.au3" ;get explorer view


Opt("SendAttachMode", 1)
Opt("SendCapslockMode", 0) 	;1 restores Caps to pre-send state (default!), 0 doesn't store
Send("{CapsLock off}")
HotKeySet("{CapsLock}","capsfun")
HotKeySet("{f5}","changeview")
HotKeySet("+{f5}","changeviewreverse")
$KBL_CUSTOM = "0xF0C00409" ;My Custom Keyboard Layout code
;$KBL_US 	= "0x04090409" ;US Keyboard Layout code
;$KBL_USINT	= "0xF0010409" ;US-international Keyboard Layout code
;$KBL_GREEK	= "0x04080408" ;Greek Keyboard Layout code


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


Func changeview()
   ; Always send normal f5
   HotKeySet("{f5}")
   Send("{f5}")
   HotKeySet("{f5}","changeview")

   If WinActive("[REGEXPCLASS:^(Cabinet|Explore)WClass$]") Then

  ; Windows Explorer on XP, Vista, 7, 8
  Local $hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )
  If Not $hExplorer Then
    MsgBox( 0, "Automating Windows Explorer", "Could not find Windows Explorer. Terminating." )
    Return
  EndIf

  ; Get an IShellBrowser interface
  GetIShellBrowser( $hExplorer )
  If Not IsObj( $oIShellBrowser ) Then
    MsgBox( 0, "Automating Windows Explorer", "Could not get an IShellBrowser interface. Terminating." )
    Return
  EndIf

  ; Get other interfaces
  GetShellInterfaces()

  ; Get current icon view
  Local $view = GetIconView() ;returns array [view,size]

  ; Determine the new view
  Local $iView, $iSize, $iNewView, $iNewSize
  $iView = $view[0] ; view
  $iSize = $view[1] ; size
  If $iView = 8 Then
	   $iNewView = 1
	   $iNewSize = 48
  Else
	  $iNewView = $iView + 1
	  If ($iNewView = 5) Or ($iNewView = 7) Then
		$iNewView += 1 ;skip from 5 to 6, or from 7 to 8
	  EndIf
  EndIf
  Switch $iNewView
  Case 2 To 4
	$iNewSize = 16
  Case 6
	$iNewSize = 48
  Case 8
	$iNewSize = 32
  EndSwitch

  ;MsgBox( 0, "iView", "Old: " & $iView )
  ;MsgBox( 0, "NewView", "New: " & $iNewView )
  SetIconView( $iNewView, $iNewSize ) ; Set view
  EndIf ;Winactive

EndFunc

Func changeviewreverse()
   If WinActive("[REGEXPCLASS:^(Cabinet|Explore)WClass$]") Then

  ; Windows Explorer on XP, Vista, 7, 8
  Local $hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )
  If Not $hExplorer Then
    MsgBox( 0, "Automating Windows Explorer", "Could not find Windows Explorer. Terminating." )
    Return
  EndIf

  ; Get an IShellBrowser interface
  GetIShellBrowser( $hExplorer )
  If Not IsObj( $oIShellBrowser ) Then
    MsgBox( 0, "Automating Windows Explorer", "Could not get an IShellBrowser interface. Terminating." )
    Return
  EndIf

  ; Get other interfaces
  GetShellInterfaces()

  ; Get current icon view
  Local $view = GetIconView() ;returns array [view,size]

  ; Determine the new view
  Local $iView, $iSize, $iNewView, $iNewSize
  $iView = $view[0] ; Icon view
  $iSize = $view[1] ; Icon size
  If $iView = 1 Then
	   $iNewView = 8
	   $iNewSize = 32
  Else
	  $iNewView = $iView - 1
	  If ($iNewView = 5) Or ($iNewView = 7) Then
		$iNewView -= 1 ;skip from 5 to 4, or from 7 to 6
	  EndIf
  EndIf
  Switch $iNewView
  Case 2 To 4
	$iNewSize = 16
  Case 6
	$iNewSize = 48
  Case 1
	$iNewSize = 48
  EndSwitch

  ;MsgBox( 0, "iView", "Old: " & $iView )
  ;MsgBox( 0, "NewView", "New: " & $iNewView )
  SetIconView( $iNewView, $iNewSize ) ; Set view
  EndIf ;Winactive

EndFunc

;---Bugs
;Sometimes the backspace-version actives caps lock with repeated presses. Using that mode should always result in caps lock off, and a matching keyboard light state.
;Related to above: keyboard light also toggles unpredicably. Light state doesn't correlate with CapsLock state

