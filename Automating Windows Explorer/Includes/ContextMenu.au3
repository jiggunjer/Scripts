#include-once

#include <GuiMenu.au3>

Global $idCmdFirst = 0x0001
Global $idCmdLast  = 0x6FFF

; Execute the default menu item
Func InvokeCommand( $hWnd, $oIShellFolder, $aPidlsRel )

	; Create an array of pointers to PIDLs
	Local $iPidls = UBound( $aPidlsRel )
	Local $aPidls = DllStructCreate( "ptr[" & $iPidls & "]" )
	For $i = 0 To $iPidls - 1
		DllStructSetData( $aPidls, 1, $aPidlsRel[$i], $i+1 )
	Next

	Local $pIContextMenu, $oIContextMenu, $tArray = DllStructCreate( "ptr" )

	; Get a pointer to the IContextMenu interface
	If $oIShellFolder.GetUIObjectOf( $NULL, $iPidls, $aPidls, $tRIID_IContextMenu, 0, $pIContextMenu ) = $S_OK Then

		Local $hPopup, $aRet, $idMenu

		; Create IContextMenu interface
		$oIContextMenu = ObjCreateInterface( $pIContextMenu, $sIID_IContextMenu, $dtag_IContextMenu )

		; Create a Popup menu
		$hPopup = _GUICtrlMenu_CreatePopup()

		; Add the default command to the Popup menu
		$aRet = $oIContextMenu.QueryContextMenu( $hPopup, 0, $idCmdFirst, $idCmdLast, $CMF_DEFAULTONLY )

		If $aRet >= 0 Then
			; Get the menu item identifier
			Local $idMenu = _GUICtrlMenu_GetItemID( $hPopup, 0 )
			; Create and fill a $tagCMINVOKECOMMANDINFO structure
			Local $tCMINVOKECOMMANDINFO = DllStructCreate( $tagCMINVOKECOMMANDINFO )
			DllStructSetData( $tCMINVOKECOMMANDINFO, "cbSize", DllStructGetSize( $tCMINVOKECOMMANDINFO ) )
			DllStructSetData( $tCMINVOKECOMMANDINFO, "fMask", 0 )
			DllStructSetData( $tCMINVOKECOMMANDINFO, "hWnd", $hWnd )
			DllStructSetData( $tCMINVOKECOMMANDINFO, "lpVerb", $idMenu - $idCmdFirst )
			DllStructSetData( $tCMINVOKECOMMANDINFO, "nShow", $SW_SHOWNORMAL )
			; Invoke the command
			$oIContextMenu.InvokeCommand( $tCMINVOKECOMMANDINFO )
		EndIf

		; Destroy menu and free memory
		_GUICtrlMenu_DestroyMenu( $hPopup )
	EndIf

EndFunc
