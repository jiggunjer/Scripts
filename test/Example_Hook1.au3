#include <MsgBoxConstants.au3>
#include <StructureConstants.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>

Global $g_hHook, $g_sBuffer = ""

Example()

Func Example()
    OnAutoItExitRegister("Cleanup")   

    $pcallback = DllCallbackGetPtr(DllCallbackRegister("MyCallback", "long", "int;wparam;lparam"));$WH_KEYBOARD_LL types
    $g_hHook = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, $pcallback, _WinAPI_GetModuleHandle(0))

    MsgBox($MB_SYSTEMMODAL, "", "Click OK, then in notepad type..." & _
            @CRLF & @CRLF & "Jon" & @CRLF & "AutoIt" & @CRLF & @CRLF & "Press Esc to exit script")

    Run("notepad.exe")
    WinWait("[CLASS:Notepad]")
    WinActivate("[CLASS:Notepad]")

    While 1
        Sleep(10)
    WEnd
EndFunc   ;==>Example

Func EvaluateKey($iKeycode)
    If (($iKeycode > 64) And ($iKeycode < 91)) _ ; a - z
            Or (($iKeycode > 96) And ($iKeycode < 123)) _ ; A - Z
            Or (($iKeycode > 47) And ($iKeycode < 58)) Then ; 0 - 9
        $g_sBuffer &= Chr($iKeycode)
        Switch $g_sBuffer
            Case "Jon"
                ToolTip("What can you say?")
            Case "AutoIt"
                ToolTip("AutoIt Rocks")
        EndSwitch
    ElseIf ($iKeycode > 159) And ($iKeycode < 164) Then
        Return
    ElseIf ($iKeycode = 27) Then ; esc key
        WinKill("[CLASS:Notepad]")
        Exit
    Else
        $g_sBuffer = ""
    EndIf
EndFunc   ;==>EvaluateKey

; ===========================================================
; callback function
; ===========================================================
Func MyCallback($nCode, $wParam, $lParam)    
    If $nCode < 0 Then
        Return _WinAPI_CallNextHookEx($g_hHook, $nCode, $wParam, $lParam)
    EndIf
    
    Local $tKEYHOOKS = DllStructCreate($tagKBDLLHOOKSTRUCT, $lParam)    
    If $wParam = $WM_KEYDOWN Then
        EvaluateKey(DllStructGetData($tKEYHOOKS, "vkCode"))
    Else
        Local $iFlags = DllStructGetData($tKEYHOOKS, "flags")
        Switch $iFlags
            Case $LLKHF_ALTDOWN
                ConsoleWrite("$LLKHF_ALTDOWN" & @CRLF)
            Case $LLKHF_EXTENDED
                ConsoleWrite("$LLKHF_EXTENDED" & @CRLF)
            Case $LLKHF_INJECTED
                ConsoleWrite("$LLKHF_INJECTED" & @CRLF)
            Case $LLKHF_UP
                ConsoleWrite("$LLKHF_UP: scanCode - " & DllStructGetData($tKEYHOOKS, "scanCode") & @TAB & "vkCode - " & DllStructGetData($tKEYHOOKS, "vkCode") & @CRLF)
        EndSwitch
    EndIf
    Return _WinAPI_CallNextHookEx($g_hHook, $nCode, $wParam, $lParam)
EndFunc   ;==>_KeyProc

Func Cleanup()
    _WinAPI_UnhookWindowsHookEx($g_hHook)
    ;DllCallbackFree($g_hStub_KeyProc)
EndFunc   ;==>Cleanup