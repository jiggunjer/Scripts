#Include <APIConstants.au3>
#Include <GUIImageList.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global Const $STM_SETIMAGE = 0x0172
Global Const $STM_GETIMAGE = 0x0173

Global $hForm, $Pic, $hPic, $hImageList, $CX, $CY, $W, $H, $Count, $Index, $Size, $hObj, $hBitmap, $hDC, $hMemDC, $hMemSv

; Initialize system image list
_WinAPI_FileIconInit()

; Retrieve system image list and calculate bitmap size
$hImageList = _WinAPI_ShellGetImageList()
If @error Then
	Exit
EndIf
$Count = _GUIImageList_GetImageCount($hImageList)
$Size = _GUIImageList_GetIconSize($hImageList)
$CX = Sqrt($Count)
If $CX Then
	$CX = Ceiling($CX)
	$CY = Ceiling($Count / $CX)
Else
	$CX = 1
	$CY = 1
EndIf
$W = $CX * ($Size[0] + 14)
$H = $CY * ($Size[1] + 14)

; Create GUI
$hForm = GUICreate('MyGUI', $W, $H)
$Pic = GUICtrlCreatePic('', 0, 0, $W, $H)
$hPic = GUICtrlGetHandle($Pic)

; Create bitmap
$hDC = _WinAPI_GetDC($hPic)
$hMemDC = _WinAPI_CreateCompatibleDC($hDC)
$hBitmap = _WinAPI_CreateCompatibleBitmap($hDC, $W, $H)
$hMemSv = _WinAPI_SelectObject($hMemDC, $hBitmap)

; Draw all icons from the system image list into bitmap
$Index = 0
For $y = 1 To $CY
	For $x = 1 To $CX
		_GUIImageList_Draw($hImageList, $Index, $hMemDC, ($x - 1) * ($Size[0] + 14) + 7, ($y - 1) * ($Size[0] + 14) + 7)
		$Index += 1
		If $Index >= $Count Then
			ExitLoop
		EndIf
	Next
Next

; Release objects
_WinAPI_ReleaseDC($hPic, $hDC)
_WinAPI_SelectObject($hMemDC, $hMemSv)
_WinAPI_DeleteDC($hMemDC)

; Set bitmap to control
_SendMessage($hPic, $STM_SETIMAGE, 0, $hBitmap)
$hObj = _SendMessage($hPic, $STM_GETIMAGE)
If $hObj <> $hBitmap Then
	_WinAPI_DeleteObject($hBitmap)
EndIf

; Show GUI
GUISetState()

Do
Until GUIGetMsg() = -3
