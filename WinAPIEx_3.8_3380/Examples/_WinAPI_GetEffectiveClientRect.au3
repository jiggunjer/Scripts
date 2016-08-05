#Include <APIConstants.au3>
#Include <GUIConstantsEx.au3>
#Include <GUIStatusBar.au3>
#Include <StaticConstants.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global $hForm, $Msg, $Menu, $Exit, $ID[2], $tRECT, $Pos

; Create GUI
$hForm = GUICreate('MyGUI', 600, 400)
$Menu = GUICtrlCreateMenu("&File")
$Exit = GUICtrlCreateMenuItem("E&xit", $Menu)
$ID[0] = _GUICtrlStatusBar_Create($hForm)
$ID[1] = GUICtrlCreateListView('', 0, 0, 600, 200, -1, 0)

; Calculate effective client area (excluding Menu, ListView, and StatusBar controls)
$tRECT = _WinAPI_GetEffectiveClientRect($hForm, $ID)
$Pos = _WinAPI_GetPosFromRect($tRECT)
GUICtrlCreateLabel($Pos[2] & 'x' & $Pos[3], $Pos[0], $Pos[1], $Pos[2], $Pos[3], BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1, 25, 400, 0, 'Tahoma')
GUICtrlSetBkColor(-1, 0xFFD0D0)

; Show GUI
GUISetState()

While 1
	$Msg = GUIGetMsg()
	Switch $Msg
		Case $GUI_EVENT_CLOSE, $Exit
			ExitLoop
	EndSwitch
WEnd
