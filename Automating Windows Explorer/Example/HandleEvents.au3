#include-once

; --- Window events ---

; After deactivation, it may be necessary to update controls with new data
Func WM_ACTIVATE( $hWnd, $iMsg, $wParam, $lParam )
	If Not BitAND( $wParam, 0xFFFF ) Then Return $GUI_RUNDEFMSG
	If Not $fActivateEvents Then Return $GUI_RUNDEFMSG

	Switch $hWnd
		Case $hGui, $hTab1, $hTab2, $hTab3
			; Does Explorer still exists, or is it closed?
			If Not WinExists( $hExplorer ) Then Close()

			; New current folder in Explorer?
			; New focused item in Explorer?
			If IsNewShellFolder() Then
				If $oIShellBrowser.QueryActiveShellView( $pIShellView ) <> $S_OK Then Close()
				GetShellInterfaces()
				; Update current folder on main GUI
				IsNewCurrentFolder()
				$fGuiUpdate = 1
				GUICtrlSendToDummy( $idGuiUpdate )
			ElseIf IsNewFocusedItem() Then
				; Update focused item on main GUI
				$fGuiUpdate = 2
				GUICtrlSendToDummy( $idGuiUpdate )
			EndIf

			; Update current Tab item
			Switch $iTab
				Case 0 ; Tab1
					GUICtrlSendToDummy( $idTab1Update )
				Case 1 ; Tab2
					GUICtrlSendToDummy( $idTab2Update )
				Case 2 ; Tab3
					Local $iIndex
					If IsNewIconView( $iIndex ) Then _
						GUICtrlSendToDummy( $idTab3Update, $iIndex )
			EndSwitch
	EndSwitch 

	Return $GUI_RUNDEFMSG
EndFunc

; Prevent child windows (for Tab items) in making main GUI title bar dimmed
Func WM_NCACTIVATE( $hWnd, $iMsg, $wParam )
	If $hWnd = $hGui Then
		If Not $wParam Then Return 1
	EndIf
EndFunc

; Double click in listview on Tab2
Func WM_NOTIFY( $hWnd, $iMsg, $iwParam, $ilParam )
	Local $tNMHDR, $hWndFrom, $iCode
	$tNMHDR = DllStructCreate( $tagNMHDR, $ilParam )
	$hWndFrom = DllStructGetData( $tNMHDR, "hWndFrom" )
	$iCode = DllStructGetData( $tNMHDR, "Code" )
	Switch $hWndFrom
		Case $hLVFolders
			Switch $iCode
				Case $NM_DBLCLK
					GUICtrlSendMsg( $idTab2R5, $BM_CLICK, 0, 0 )
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc


; --- Control events ---

Func SwitchTabItem( $i )
	$iTab = $i
	$fActivateEvents = False
	Switch $iTab
		Case 0 ; Tab1
			GUISetState( @SW_HIDE, $hTab2 )
			GUISetState( @SW_HIDE, $hTab3 )
			GUISetState( @SW_SHOW, $hTab1 )
			GUICtrlSendToDummy( $idTab1Update )
		Case 1 ; Tab2
			GUISetState( @SW_HIDE, $hTab1 )
			GUISetState( @SW_HIDE, $hTab3 )
			GUISetState( @SW_SHOW, $hTab2 )
			GUICtrlSendToDummy( $idTab2Update )
		Case 2 ; Tab3
			Local $iIndex
			GUISetState( @SW_HIDE, $hTab1 )
			GUISetState( @SW_HIDE, $hTab2 )
			GUISetState( @SW_SHOW, $hTab3 )
			If IsNewIconView( $iIndex ) Then _
				GUICtrlSendToDummy( $idTab3Update, $iIndex )
	EndSwitch
	$fActivateEvents = True
EndFunc

; All updates of controls is done here.
; If a control needs to be updated, a
; corresponding update event is generated.
Func HandleUpdateEvents( $iUpdate )
	Switch $iUpdate
		; Update controls on Gui
		Case $idGuiUpdate
			If $fGuiUpdate = 1 Then
				FillCurrentFolder()
				IsNewFocusedItem( True )
				FillFocusedItem()
			ElseIf $fGuiUpdate = 2 Then
				FillFocusedItem()
			EndIf
			$fGuiUpdate = 0

		; Update controls on Tab1
		Case $idTab1Update
			Local $iCount = CountItems(), $fState
			GUICtrlSetData( $idTab1G1, "Get items (" & $iCount & ")" )
			GUICtrlSetData( $idTab1G2, "Get selected (" & CountItems( True ) & ")" )
			If IsNewButtonsState( $iCount, $fState ) Then _
				SetButtonsState( $fState )
			FillFilesFolders()

		; Update controls on Tab2
		Case $idTab2Update
			FillFolders()

		; Update controls on Tab3
		Case $idTab3Update
			SetIconViewTab3( GUICtrlRead( $idTab3Update ) )
	EndSwitch
EndFunc

Func HandleTab1Events( $idRadio )
	Local Static $idRadioLast = 0
	Local $aItems

	If $idRadioLast And $idRadioLast <> $idRadio Then _
		GUICtrlSetState( $idRadioLast, $GUI_UNCHECKED )
	$idRadioLast = $idRadio

	Switch $idRadio
		Case $idTab1R1
			; Get all items
			$aItems = GetItems()
			_ArrayDisplay( $aItems, "All items" )

		Case $idTab1R2
			; Get all file items
			$aItems = GetFiles()
			_ArrayDisplay( $aItems, "All file items" )

		Case $idTab1R3
			; Get all folder items
			$aItems = GetFolders()
			_ArrayDisplay( $aItems, "All folder items" )

		Case $idTab1R4
			; Get selected items
			$aItems = GetItems( True )
			_ArrayDisplay( $aItems, "Selected items" )

		Case $idTab1R5
			; Get selected file items
			$aItems = GetFiles( True )
			_ArrayDisplay( $aItems, "Selected file items" )

		Case $idTab1R6
			; Get selected folder items
			$aItems = GetFolders( True )
			_ArrayDisplay( $aItems, "Selected folder items" )

		Case $idTab1R7
			; Open selected items
			$aItems = GetPidls( True )
			If IsArray( $aItems ) Then
				InvokeCommand( $hExplorer, $oIShellFolder, $aItems )
				For $i = 0 To UBound( $aItems ) - 1
					_WinAPI_CoTaskMemFree( $aItems[$i] )
				Next
			EndIf

		Case $idTab1S1
			; Set all items selected
			Local $iItems = CountItems()
			If $iItems Then
				If CountItems( True ) <> $iItems Then
					For $i = 0 To $iItems - 1
						SetSelectedItem( $i )
					Next
					GUICtrlSetData( $idTab1G2, "Get selected (" & CountItems( True ) & ")" )
				EndIf
			EndIf

		Case $idTab1S2
			; Set all items deselected
			Local $iItems = CountItems()
			If $iItems Then
				If CountItems( True ) Then
					For $i = 0 To $iItems - 1
						SetSelectedItem( $i, False )
					Next
					GUICtrlSetData( $idTab1G2, "Get selected (" & CountItems( True ) & ")" )
				EndIf
			EndIf

		Case $idTab1S3
			; Reverse selected state of all items
			Local $iItems = CountItems()
			If $iItems Then
				Local $iSel = CountItems( True )
				If Not $iSel Then
					; No items selected, select all
					GUICtrlSendMsg( $idTab1S1, $BM_CLICK, 0, 0 )
				ElseIf $iSel = $iItems Then
					; All items selected, deselect all
					GUICtrlSendMsg( $idTab1S2, $BM_CLICK, 0, 0 )
				Else
					Local $aItems = GetPidls(), $aSel = GetPidls( True ), $j = 0
					For $i = 0 To $iItems - 1
						If $j < $iSel And ILIsEqual( $aItems[$i], $aSel[$j] ) Then
							_WinAPI_CoTaskMemFree( $aItems[$i] )
							_WinAPI_CoTaskMemFree( $aSel[$j] )
							SetSelectedItem( $i, False )
							$j += 1
						Else
							_WinAPI_CoTaskMemFree( $aItems[$i] )
							SetSelectedItem( $i )
						EndIf
					Next
					GUICtrlSetData( $idTab1G2, "Get selected (" & CountItems( True ) & ")" )
				EndIf
			EndIf

		Case $idTab1S4
			; Set focused item
			Local $aSel = _GUICtrlListView_GetSelectedIndices( $hLVFilesFolders, True )
			If $aSel[0] Then
				; Select new focused item
				SetFocusedItem( $aFilesFolders[$aSel[1]][1], True )
				If IsNewFocusedItem() Then
					$fGuiUpdate = 2
					GUICtrlSendToDummy( $idGuiUpdate )
				EndIf
			EndIf

		Case $idTab1S5
			; Set selected items
			Local $aSel = _GUICtrlListView_GetSelectedIndices( $hLVFilesFolders, True )
			For $i = 1 To $aSel[0]
				SetSelectedItem( $aFilesFolders[$aSel[$i]][1], True, True )
			Next
			GUICtrlSetData( $idTab1G2, "Get selected (" & CountItems( True ) & ")" )
	EndSwitch
EndFunc

Func HandleTab2Events( $idRadio )
	Local Static $idRadioLast = 0

	If $idRadioLast And $idRadioLast <> $idRadio Then _
		GUICtrlSetState( $idRadioLast, $GUI_UNCHECKED )
	$idRadioLast = $idRadio

	Switch $idRadio
		Case $idTab2R1
			; Browse to Desktop
			If ILIsEqual( $pPidlAbsCurFolder, $pPidlAbsDesktop ) Then Return
			SetCurrentFolder( $pPidlAbsDesktop, $SBSP_ABSOLUTE )

		Case $idTab2R2
			; Browse to Computer
			If ILIsEqual( $pPidlAbsCurFolder, $pPidlAbsComputer ) Then Return
			SetCurrentFolder( $pPidlAbsComputer, $SBSP_ABSOLUTE )

		Case $idTab2R3
			; Browse to home/startup folder
			If ILIsEqual( $pPidlAbsCurFolder, $pPidlAbsHome ) Then Return
			SetCurrentFolder( $pPidlAbsHome, $SBSP_ABSOLUTE )

		Case $idTab2R4
			; Browse to parent folder
			If ILIsEqual( $pPidlAbsCurFolder, $pPidlAbsDesktop ) Then Return
			SetCurrentFolder( 0, $SBSP_PARENT )

		Case $idTab2R5
			; Browse to child folder
			Local $aSel = _GUICtrlListView_GetSelectedIndices( $hLVFolders, True )
			If $aSel[0] Then SetCurrentFolder( $aFolders[$aSel[1]][1], $SBSP_RELATIVE )
	EndSwitch

	If IsNewShellFolder() Then
		GetShellInterfaces()
		IsNewCurrentFolder()
		$fGuiUpdate = 1
		GUICtrlSendToDummy( $idGuiUpdate )
		GUICtrlSendToDummy( $idTab2Update )
	EndIf
EndFunc

Func HandleTab3Events( $idRadio )
	Local Static $idRadioLast = 0
	Local $iViewIndex, $iViewMode, $iIconSize

	If $idRadioLast And $idRadioLast <> $idRadio Then _
		GUICtrlSetState( $idRadioLast, $GUI_UNCHECKED )
	$idRadioLast = $idRadio

	Switch $idRadio
		; Windws 7, 8
		Case $idTab3RW78 To $idTab3RW78 + 7
			$iViewIndex = $aW78ViewModeIndex[$idRadio-$idTab3RW78]
			$iViewMode = $aViewModes[$iViewIndex]
			$iIconSize = $aIconSizes[$iViewIndex]
			SetIconView( $iViewMode, $iIconSize )

		; Windws Vista
		Case $idTab3RVst To $idTab3RVst + 6
			$iViewIndex = $aVstViewModeIndex[$idRadio-$idTab3RVst]
			$iViewMode = $aViewModes[$iViewIndex]
			$iIconSize = $aIconSizes[$iViewIndex]
			SetIconView( $iViewMode, $iIconSize )

		; Windws XP
		Case $idTab3RWXP To $idTab3RWXP + 4
			$iViewIndex = $aWXPViewModeIndex[$idRadio-$idTab3RWXP]
			$iViewMode = $aViewModesXP[$iViewIndex]
			SetIconView( $iViewMode )
	EndSwitch
EndFunc
