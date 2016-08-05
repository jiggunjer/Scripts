#NoTrayIcon
#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer2.au3" ;UDF get/set icon view

HotKeySet("{f5}","ChangeViewForward")
HotKeySet("+{f5}","ChangeViewReverse")

while 1
	Sleep(1000)
WEnd

;--- Functions
Func ChangeViewForward()
   HotKeySet("{f5}")
   Send("{f5}")
   ChangeView("forward")
   HotKeySet("{f5}","ChangeViewForward")
EndFunc

Func ChangeViewReverse()
   HotKeySet("+{f5}")
   Send("+{f5}")
   ChangeView("backward")
   HotKeySet("+{f5}","ChangeViewReverse")
EndFunc

Func ChangeView($direction)
    Local $forward = False
    if $direction == "forward" then $forward = True
    If WinActive("[REGEXPCLASS:^(Cabinet|Explore)WClass$]") Then
    ;---

    Local $hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )
    GetIShellBrowser( $hExplorer )
    GetShellInterfaces()

    Local $view = GetIconView() ;returns array [view,size]
    Local $iNewView, $iNewSize
    ;ConsoleWrite("STARTVIEW: " & $view[0] & " ")

    If $forward Then
    If $view[0] = 8 Then
      $iNewView = 1
    Else
      $iNewView = $view[0] + 1
      If ($iNewView = 5) Or ($iNewView = 7) Then $iNewView += 1 ;skip from 5 to 6, or from 7 to 8
    EndIf
    Else ;backward
      If $view[0] = 1 Then
        $iNewView = 8
      Else
        $iNewView = $view[0] - 1
        If ($iNewView = 5) Or ($iNewView = 7) Then $iNewView -= 1 ;skip from 5 to 4, or from 7 to 6
      EndIf
    EndIf

    Switch $iNewView
    Case 1
      $iNewSize = 48
    Case 2 To 4
      $iNewSize = 16
    Case 6
      $iNewSize = 48
    Case 8
      $iNewSize = 32
    EndSwitch
    ;ConsoleWrite("ENDVIEW: " & $iNewView & " ")
    SetIconView( $iNewView, $iNewSize ) ; Set view

    ;---
    EndIf ;Winactive
EndFunc


;Icon sizes are standard: 16, 48, 96, 196
;Details & list: 16
;Tiles: 48
;Content: 32 (!)
;~ FVM_ICON        = 1,  (48, 96, 196)
;~ FVM_SMALLICON   = 2,  (16)
;~ FVM_LIST        = 3,
;~ FVM_DETAILS     = 4,
;~ FVM_THUMBNAIL   = 5,  (seems to be same as ICON in win7)
;~ FVM_TILE        = 6,
;~ FVM_THUMBSTRIP  = 7,  (seems to be same as ICON in win7)
;~ FVM_CONTENT     = 8,

