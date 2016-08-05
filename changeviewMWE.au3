#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer.au3" ;UDF get/set icon view
HotKeySet("{f5}","ChangeViewForward")

While 1
	Sleep(1000)
WEnd

;--- Functions
Func ChangeViewForward()  
   ChangeView("forward")
EndFunc

Func ChangeView($direction)
    Local $forward = False   
    if $direction == "forward" then $forward = True    

    ; Always send normal f5
   HotKeySet("{f5}")
   Send("{f5}")
   HotKeySet("{f5}","ChangeViewForward")

   ; If WinActive("[REGEXPCLASS:^(Cabinet|Explore)WClass$]") Then

  ; ; Windows Explorer on XP, Vista, 7, 8
  ; ;Local $hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )  
  ; ; GetIShellBrowser( $hExplorer )
  ; ; GetShellInterfaces()  
  
  ; EndIf ;Winactive
EndFunc


;---Bugs
;Sometimes the backspace-version actives caps lock with repeated presses. Using that mode should always result in caps lock off, and a matching keyboard light state.
;Related to above: keyboard light also toggles unpredicably. Light state doesn't correlate with CapsLock state

