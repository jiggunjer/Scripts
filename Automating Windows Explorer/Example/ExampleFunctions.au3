#include-once

#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <TabConstants.au3>
#include <GuiListView.au3>
#include <GuiButton.au3>
#include <Array.au3>

Global $hExplorer, $hGui, $idGuiUpdate, $fGuiUpdate = 1, $fActivateEvents = False, $iTab = 0

Global $hImageList, $idLVCurrentFolder, $hLVCurrentFolder, $idLVFocusedItem, $hLVFocusedItem

Global $pPidlAbsCurFolder = 0, $pPidlRelFocusedItem = 0, $pPidlAbsHome
Global $pPidlAbsComputer = _WinAPI_ShellGetSpecialFolderLocation( $CSIDL_DRIVES )

Global $hTab1, $idTab1Update, $idLVFilesFolders, $hLVFilesFolders, $aFilesFolders = 0, $idTab1G1, $idTab1G2
Global $idTab1R1, $idTab1R2, $idTab1R3, $idTab1R4, $idTab1R5, $idTab1R6, $idTab1R7, $idTab1S1, $idTab1S2, $idTab1S3, $idTab1S4, $idTab1S5

Global $hTab2, $idTab2Update, $idLVFolders, $hLVFolders, $aFolders = 0
Global $idTab2R1, $idTab2R2, $idTab2R3, $idTab2R4, $idTab2R5

Global $hTab3, $idTab3Update, $iGheight
Global $idTab3RW78, $aW78ViewModeIndex = [ 7, 6, 0, 1, 2, 3, 4, 5 ]
Global $idTab3RVst, $aVstViewModeIndex = [ 7, 6, 0, 1, 2, 3, 4 ]
Global $idTab3RWXP, $aWXPViewModeIndex = [ 3, 4, 0, 1, 2 ]
Global $aViewModes = [ $FVM_ICON, $FVM_SMALLICON, $FVM_LIST, $FVM_DETAILS, $FVM_TILE, $FVM_CONTENT, $FVM_ICON, $FVM_ICON ]
Global $aIconViewToButtonIndex = [ -1, 2, 3, 4, 5, -1, 6, -1, 7 ], $aIconSizes = [ 48, 16, 16, 16, 48, 32, 96, 256 ]
Global $aViewModesXP = [ $FVM_ICON, $FVM_LIST, $FVM_DETAILS, $FVM_THUMBNAIL, $FVM_TILE ]
Global $aIconViewToButtonIndexXP = [ -1, 2, -1, 3, 4, 0, 1 ]

#include "CreateTabItems.au3"
#include "HandleEvents.au3"


; --- Functions for controls on main GUI ---

; Current folder on main GUI
Func IsNewCurrentFolder()
	Local Static $pPidlAbsCurFolderPrev = 0
	If $pPidlAbsCurFolder Then _
		_WinAPI_CoTaskMemFree( $pPidlAbsCurFolder )
	$pPidlAbsCurFolder = GetCurrentFolder()
	If $pPidlAbsCurFolderPrev Then
		If ILIsEqual( $pPidlAbsCurFolderPrev, $pPidlAbsCurFolder ) Then Return False
		_WinAPI_CoTaskMemFree( $pPidlAbsCurFolderPrev )
	EndIf
	$pPidlAbsCurFolderPrev = ILClone( $pPidlAbsCurFolder )
	Return True
EndFunc

; Current folder on main GUI
Func FillCurrentFolder()
	Local $sPath
	SHGetPathFromIDList( $pPidlAbsCurFolder, $sPath )
	If $sPath = "" Or StringInStr( $sPath, "::" ) Or _
	   ILIsEqual( $pPidlAbsCurFolder, $pPidlAbsDesktop ) Then _
	     $sPath = GetDisplayNameAbs( $pPidlAbsCurFolder )
	_GUICtrlListView_DeleteAllItems( $hLVCurrentFolder )
	_GUICtrlListView_AddItem( $hLVCurrentFolder, $sPath, GetIconIndex( $pPidlAbsCurFolder ) )
EndFunc

; Focused item on main GUI
Func IsNewFocusedItem( $fNewFolder = False )
	Local Static $pPidlRelFocusedItemPrev = 0
	If $pPidlRelFocusedItem Then _
		_WinAPI_CoTaskMemFree( $pPidlRelFocusedItem )
	$pPidlRelFocusedItem = GetFocusedItem()

	; New folder
	If $fNewFolder Then
		If $pPidlRelFocusedItemPrev Then _
			_WinAPI_CoTaskMemFree( $pPidlRelFocusedItemPrev )
		If $pPidlRelFocusedItem Then
			$pPidlRelFocusedItemPrev = ILClone( $pPidlRelFocusedItem )
		Else
			$pPidlRelFocusedItemPrev = 0
		EndIf
		; Allways return True for a new folder
		; This ensures that the focused item in previous folder is cleared
		; Even if the new folder is empty (which means that there is no focused item)
		Return True
	EndIf

	; No focused item
	; Will only be the case if the folder is empty
	If Not $pPidlRelFocusedItem Then
		If Not $pPidlRelFocusedItemPrev Then Return False
		_WinAPI_CoTaskMemFree( $pPidlRelFocusedItemPrev )
		$pPidlRelFocusedItemPrev = 0
		Return True
	EndIf

	; Focused item
	If $pPidlRelFocusedItemPrev Then
		; Return False if same as previous
		If ILIsEqual( $pPidlRelFocusedItemPrev, $pPidlRelFocusedItem ) Then Return False
		; New focused item, free memory used by old previous
		_WinAPI_CoTaskMemFree( $pPidlRelFocusedItemPrev )
	EndIf
	; Store focused item in new previous
	$pPidlRelFocusedItemPrev = ILClone( $pPidlRelFocusedItem )
	Return True
EndFunc

; Focused item on main GUI
Func FillFocusedItem()
	_GUICtrlListView_DeleteAllItems( $hLVFocusedItem )
	If Not $pPidlRelFocusedItem Then Return
	Local $sName = GetDisplayNameRel( $pPidlRelFocusedItem )
	Local $pPidl = ILCombine( $pPidlAbsCurFolder, $pPidlRelFocusedItem )
	_GUICtrlListView_AddItem( $hLVFocusedItem, $sName, GetIconIndex( $pPidl ) )
	_WinAPI_CoTaskMemFree( $pPidl )
EndFunc


; --- Functions for controls on Tab1 ---

; Listview on Tab1
Func FillFilesFolders()
	Local $pPidl

	; Free memory used by old items
	If IsArray( $aFilesFolders ) Then
		For $i = 0 To UBound( $aFilesFolders ) - 1
			_WinAPI_CoTaskMemFree( $aFilesFolders[$i][1] )
		Next
	EndIf

	; Get new items and fill listview
	$aFilesFolders = GetItems( False, False, True, 100 )
	_GUICtrlListView_BeginUpdate( $hLVFilesFolders )
	_GUICtrlListView_DeleteAllItems( $hLVFilesFolders )
	For $i = 0 To UBound( $aFilesFolders ) - 1
		$pPidl = ILCombine( $pPidlAbsCurFolder, $aFilesFolders[$i][1] )
		_GUICtrlListView_AddItem( $hLVFilesFolders, $aFilesFolders[$i][0], GetIconIndex( $pPidl ) )
		_WinAPI_CoTaskMemFree( $pPidl )
	Next
	_GUICtrlListView_EndUpdate( $hLVFilesFolders )
EndFunc

; Lower left group on Tab1
; Disable buttons if more than 100 items
Func IsNewButtonsState( $iCount, ByRef $fState )
	Local Static $fStatePrev = $GUI_ENABLE
	$fState = $iCount > 100 ? $GUI_DISABLE : $GUI_ENABLE
	If $fStatePrev = $fState Then Return False
	$fStatePrev = $fState
	Return True
EndFunc

Func SetButtonsState( $fState )
	GUICtrlSetState( $idTab1S1, $fState )
	GUICtrlSetState( $idTab1S2, $fState )
	GUICtrlSetState( $idTab1S3, $fState )
EndFunc


; --- Functions for controls on Tab2 ---

; Listview on Tab2
Func FillFolders()
	Local $pPidl

	; Free memory used by old child folders
	If IsArray( $aFolders ) Then
		For $i = 0 To UBound( $aFolders ) - 1
			_WinAPI_CoTaskMemFree( $aFolders[$i][1] )
		Next
	EndIf

	; Get new child folders and fill listview
	$aFolders = GetFolders( False, False, True, 100 )
	_GUICtrlListView_BeginUpdate( $hLVFolders )
	_GUICtrlListView_DeleteAllItems( $hLVFolders )
	For $i = 0 To UBound( $aFolders ) - 1
		$pPidl = ILCombine( $pPidlAbsCurFolder, $aFolders[$i][1] )
		_GUICtrlListView_AddItem( $hLVFolders, $aFolders[$i][0], GetIconIndex( $pPidl ) )
		_WinAPI_CoTaskMemFree( $pPidl )
	Next
	_GUICtrlListView_EndUpdate( $hLVFolders )
EndFunc


; --- Functions for controls on Tab3 ---

; Check icon view on Tab3
Func IsNewIconView( ByRef $iIndex )
	Local Static $iIconIndex = -1
	$iIndex = GetIconViewExplorer()
	If $iIconIndex = $iIndex Then Return False
	$iIconIndex = $iIndex
	Return True
EndFunc

; Get Explorer icon view for Tab3
Func GetIconViewExplorer()
	If @OSVersion = "WIN_XP" Then Return _
		$aIconViewToButtonIndexXP[GetIconView()]
	Local $aView = GetIconView() ; [ view, size ]
	Local $iIndex = $aIconViewToButtonIndex[$aView[0]]
	If $iIndex <> 2 Then Return $iIndex
	Switch $aView[1]
		Case 48
			Return 2
		Case 96
			Return 1
		Case 256
			Return 0
	EndSwitch
EndFunc

; Set icon view on Tab3
Func SetIconViewTab3( $iIndex )
	Switch @OSVersion
		Case "WIN_XP"
			GUICtrlSetState( $idTab3RWXP + $iIndex, $GUI_CHECKED )
		Case "WIN_VISTA"
			GUICtrlSetState( $idTab3RVst + $iIndex, $GUI_CHECKED )
		Case Else ; 7, 8
			GUICtrlSetState( $idTab3RW78 + $iIndex, $GUI_CHECKED )
	EndSwitch
EndFunc


; --- Other functions ---

Func CreateListView( $x, $y, $w, $rows, ByRef $idLV, ByRef $hLV, ByRef $iLVheight, $iStyle = $LVS_NOCOLUMNHEADER )
	$iLVheight = ( @OSVersion = "WIN_XP" ) ? $rows * 17 + 4 : $rows * 19 + 4
	$idLV = GUICtrlCreateListView( "", $x, $y, $w, $iLVheight, $iStyle, $WS_EX_CLIENTEDGE )
	$hLV = GUICtrlGetHandle( $idLV )
	_GUICtrlListView_SetExtendedListViewStyle( $hLV, BitOR( $LVS_EX_DOUBLEBUFFER, $LVS_EX_FULLROWSELECT ) )
	_GUICtrlListView_SetImageList( $hLV, $hImageList, 1 )
	_GUICtrlListView_AddColumn( $hLV, "", $w - 22 )
EndFunc

Func Close()
	GUIRegisterMsg( $WM_ACTIVATE, "" )
	MsgBox( $MB_SYSTEMMODAL + $MB_ICONERROR, "Automating Windows Explorer", "Could not find Windows Explorer. Terminating.", 0, $hGui )
	Cleanup()
	Exit
EndFunc

Func Cleanup()
	; Free memory used by items in LV on Tab1
	If IsArray( $aFilesFolders ) Then
		For $i = 0 To UBound( $aFilesFolders ) - 1
			_WinAPI_CoTaskMemFree( $aFilesFolders[$i][1] )
		Next
	EndIf

	; Free memory used by child folders in LV on Tab2
	If IsArray( $aFolders ) Then
		For $i = 0 To UBound( $aFolders ) - 1
			_WinAPI_CoTaskMemFree( $aFolders[$i][1] )
		Next
	EndIf

	; Free memory used by PIDLs
	If $pPidlRelFocusedItem Then _
		_WinAPI_CoTaskMemFree( $pPidlRelFocusedItem )
	_WinAPI_CoTaskMemFree( $pPidlAbsCurFolder )
	_WinAPI_CoTaskMemFree( $pPidlAbsDesktop )
	_WinAPI_CoTaskMemFree( $pPidlAbsComputer )
	_WinAPI_CoTaskMemFree( $pPidlAbsHome )
	GUIDelete( $hGui )
EndFunc
