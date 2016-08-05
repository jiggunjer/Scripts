#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer.au3" ;UDF get/set icon view

HotKeySet("{f5}","ChangeViewForward")
HotKeySet("+{f5}","ChangeViewReverse")

while 1
	Sleep(1000)
WEnd

;--- Functions
Func ChangeViewForward()
   changeview("forward")
EndFunc

Func ChangeViewReverse()
    changeview("backward")
EndFunc

Func changeview($direction)
    if $direction == "forward" then
        Local $forward = True
    Else
        Local $forward = False
    EndIf

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


;---Bugs
;Sometimes the backspace-version actives caps lock with repeated presses. Using that mode should always result in caps lock off, and a matching keyboard light state.
;Related to above: keyboard light also toggles unpredicably. Light state doesn't correlate with CapsLock state

