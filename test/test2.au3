; #FUNCTION# ====================================================================================================================
; Name...........: GetActiveKeyboardLayout() Function
; Description ...: Get Active keyboard layout

; Author ........: Fredj A. Jad (DCCD)
; MSDN  .........: GetWindowThreadProcessId Function  ,http://msdn.microsoft.com/en-us/library/ms633522(VS.85).aspx
; MSDN  .........: GetKeyboardLayout Function         ,http://msdn.microsoft.com/en-us/library/ms646296(VS.85).aspx
; ===============================================================================================================================
While 1
    If Sleep(250)  Then TrayTip('Active!', GetActiveKeyboardLayout(WinGetHandle('')), 5, 1)
WEnd

Func GetActiveKeyboardLayout($hWnd)
    Local $aRet = DllCall('user32.dll', 'long', 'GetWindowThreadProcessId', 'hwnd', $hWnd, 'ptr', 0)
    $aRet = DllCall('user32.dll', 'long', 'GetKeyboardLayout', 'long', $aRet[0])
    Return '0000' & Hex($aRet[0], 4)
EndFunc   ;==>GetActiveKeyboardLayout