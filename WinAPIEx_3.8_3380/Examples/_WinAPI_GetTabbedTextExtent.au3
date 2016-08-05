#Include <APIConstants.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global Const $Text = 'String' & @TAB & 'with' & @TAB & 'tab' & @TAB & 'characters'

Global $hForm, $Pos, $Size

; Create GUI
$hForm = GUICreate('MyGUI', 0, 0, -1, -1, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
GUICtrlCreateLabel($Text, 0, 0)
GUICtrlSetFont(-1, 26, 400, 0, 'Comic Sans MS')

; Compute dimensions of a string and resize the window
$Size = _GetTabbedStringSize(-1, $Text)
If IsArray($Size) Then
	GUICtrlSetPos(-1, 0, 0, $Size[0], $Size[1])
	$Pos = WinGetPos($hForm)
	If Not @error Then
		WinMove($hForm, '', (@DesktopWidth - ($Pos[2] + $Size[0])) / 2, (@DesktopHeight - ($Pos[3] + $Size[1])) / 2, $Pos[2] + $Size[0], $Pos[3] + $Size[1])
	EndIf
EndIf

; Show GUI
GUISetState()

Do
Until GUIGetMsg() = -3

Func _GetTabbedStringSize($hWnd, $sText)

;	Local $tTM = DllStructCreate($tagTEXTMETRIC)
	Local $tTM = DllStructCreate('long tmHeight;long tmAscent;long tmDescent;long tmInternalLeading;long tmExternalLeading;long tmAveCharWidth;long tmMaxCharWidth;long tmWeight;long tmOverhang;long tmDigitizedAspectX;long tmDigitizedAspectY;wchar tmFirstChar;wchar tmLastChar;wchar tmDefaultChar;wchar tmBreakChar;byte tmItalic;byte tmUnderlined;byte tmStruckOut;byte tmPitchAndFamily;byte tmCharSet')
	Local $tSIZE, $hDC, $hSv, $hFont, $Ret, $Size = 0, $Tab = 0

	If Not IsHWnd($hWnd) Then
		$hWnd = GUICtrlGetHandle($hWnd)
		If Not $hWnd Then
			Return 0
		EndIf
	EndIf

	$hDC = _WinAPI_GetDC($hWnd)
	$hFont = _SendMessage($hWnd, $WM_GETFONT)
	$hSv = _WinAPI_SelectObject($hDC, $hFont)
	$Ret = DllCall('gdi32.dll', 'int', 'GetTextMetricsW', 'hwnd', $hDC, 'ptr', DllStructGetPtr($tTM))
	If (Not @error) And ($Ret[0]) Then
		$Tab = 8 * DllStructGetData($tTM, 'tmAveCharWidth')
	EndIf
	$tSIZE = _WinAPI_GetTabbedTextExtent($hDC, $sText, $Tab)
	If Not @error Then
		Dim $Size[2]
		For $i = 0 To 1
			$Size[$i] = DllStructGetData($tSIZE, $i + 1)
		Next
	EndIf
	_WinAPI_SelectObject($hDC, $hSv)
	_WinAPI_ReleaseDC($hWnd, $hDC)
	Return $Size
EndFunc   ;==>_GetTabbedStringSize
