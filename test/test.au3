#include <WinAPIProc.au3>
#include <WinAPI.au3>
#include <WinAPISys.au3>
#include <Array.au3>
#include <GUIConstants.au3>
#include <APISysConstants.au3>


Global $hGUI = GUICreate("GUI")

GUISetState(@SW_SHOW, $hGUI)

GUIRegisterMsg(_WinAPI_RegisterWindowMessage('SHELLHOOK'), 'WM_SHELLHOOK')
_WinAPI_RegisterShellHookWindow($hGUI)



;Run("notepad.exe")

Local $nMsg=0
While 1
    $nMsg = GUIGetMsg()
    Select
        Case $nMsg = $GUI_EVENT_CLOSE
            ExitLoop
    EndSelect

WEnd

_WinAPI_DeregisterShellHookWindow($hGUI)
Exit

Func WM_SHELLHOOK($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg
    Switch $wParam
        Case $HSHELL_WINDOWDESTROYED
            ConsoleWrite('Destroyed: ' & @CRLF & _
                    @TAB & 'PID: ' & WinGetProcess($lParam) & @CRLF & _ ; This will be -1.
                    @TAB & 'ClassName: ' & _WinAPI_GetClassName($lParam) & @CRLF & _ ; This will be empty.
                    @TAB & 'hWnd: ' & $lParam & @CRLF) ; This will be the handle of the window closed.
        Case $HSHELL_WINDOWCREATED
            ConsoleWrite('Created: ' & @CRLF & _
                    @TAB & 'PID: ' & WinGetProcess($lParam) & @CRLF & _
                    @TAB & 'ClassName: ' & _WinAPI_GetClassName($lParam) & @CRLF & _
                    @TAB & 'hWnd: ' & $lParam & @CRLF) ; This will be the handle of the window closed.
    EndSwitch
EndFunc   ;==>WM_SHELLHOOK