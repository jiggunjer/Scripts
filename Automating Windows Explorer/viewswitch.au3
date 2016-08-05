#include <GuiListView.au3>
#include <MsgBoxConstants.au3>
#include <WinAPI.au3>


HotKeySet("{f}","viewswitch")

while 1
	Sleep(400)
 WEnd

 Func viewswitch()
	;MsgBox(1,"test","tedt")
   $Hwnd = ControlGetHandle("", "", 1)
   $Cid = _WinAPI_GetDlgCtrlID ( $Hwnd )
   _GUICtrlListView_SetView($Cid, 2);2=listview
   $res = _GUICtrlListView_GetView ( $hWnd )
   MsgBox(1,"res",$res)
 EndFunc