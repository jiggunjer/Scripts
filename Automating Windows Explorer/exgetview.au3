#include "Includes\AutomatingWindowsExplorer.au3"

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

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
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

  ;MsgBox( 0, "NewView", $iNewView )
  SetIconView( $iNewView, $iNewSize ) ; Set details view
  Sleep( 1000 )                     ; Wait
  SetIconView( $iView, $iSize )     ; Restore old view
EndFunc