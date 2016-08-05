;#NoTrayIcon
#include "Automating Windows Explorer\Includes\AutomatingWindowsExplorer.au3" ;UDF
#include <Array.au3>
#include <GUIConstantsEx.au3>
;Opt("GUIOnEventMode", 1)

Local $str = "Address: ftp" ;part of visible text in explorer control, unique to ftp, I think...
Local $CheckedWindows[5] ;Keep track of activated windows because I don't have a shell hook for window.created
Local $hExplorer
$oShell = ObjCreate("shell.application")
$wOb = ObjEvent($oShell,"_CallBack_")


while 1
	Sleep(2000)    
    $hExplorer = WinWaitActive("[CLASS:CabinetWClass]", $str)
    ;GUISetOnEvent($GUI_EVENT_CLOSE, "CloseCallback",$hExplorer)
    
    If not ContainsElement($CheckedWindows,$hExplorer) then ;Only trigger on a *new* window
        setFTPview($hExplorer)
        _ArrayAdd($CheckedWindows,$hExplorer)
    EndIf
    ;delete unused handles to prevent aliases or large array, but I don't know the shell hook for window.closed
    ;alternative is to periodically loop through existing windows and delete non-existing handles (todo)
WEnd

func ContainsElement($arr,$el)
    Local $Bound = UBound($arr)
    For $i=0 to ($Bound -1)
        If $arr[$i] == $el then return True
    Next
    return False
Endfunc

func setFTPview($hExplorer)
    GetIShellBrowser( $hExplorer )
    If Not IsObj( $oIShellBrowser ) Then
        MsgBox( 0, "Automating Windows Explorer", "Could not get an IShellBrowser interface. Terminating." )
    Return
    EndIf
    GetShellInterfaces() ; Get other interfaces, might not be needed
    SetIconView($FVM_LIST)
    Sleep(1000)
endfunc

func CloseCallback()
    msgbox(0,"t","CLOSED")
Endfunc
