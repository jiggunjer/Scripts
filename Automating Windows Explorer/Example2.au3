Opt( "MustDeclareVars", 1 )

#include "Includes\AutomatingWindowsExplorer.au3"
#include "Example\ExampleFunctions.au3"

Example()


Func Example()

	; Windows Explorer on XP, Vista, 7, 8
	$hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )
	If Not $hExplorer Then
		MsgBox( $MB_SYSTEMMODAL + $MB_ICONERROR, "Automating Windows Explorer", "Could not find Windows Explorer. Terminating." )
		Return
	EndIf

	;Local $oErrorHandler = ObjEvent( "AutoIt.Error", "ObjErrFunc" )

	; Get an IShellBrowser interface
	GetIShellBrowser( $hExplorer )
	If Not IsObj( $oIShellBrowser ) Then
		MsgBox( $MB_SYSTEMMODAL + $MB_ICONERROR, "Automating Windows Explorer", "Could not get an IShellBrowser interface. Terminating." )
		Return
	EndIf

	; Get other interfaces
	GetShellInterfaces()

	; Create GUI
	Local $width = 700
	Local $height = ( @OSVersion = "WIN_XP" ) ? 12 * 17 + 4 : 11 * 19 + 4
	Local $iGuiHeight = 109 + ( @OSVersion <> "WIN_XP" ) * 4 + $height + 70
	Local $iStyle = BitOR( $GUI_SS_DEFAULT_GUI, $WS_CLIPCHILDREN )
	$hGui = GUICreate( "Automating Windows Explorer", $width, $iGuiHeight, 200, 100, $iStyle ) ; $WS_EX_TOPMOST
	$idGuiUpdate = GUICtrlCreateDummy()
	Local $xPos = 10, $yPos = 10, $iLVheight
	$width -= 2 * $xPos

	; Windows messages
	GUIRegisterMsg( $WM_ACTIVATE,   "WM_ACTIVATE" )   ; After deactivation, it may be necessary to update controls with new data
	GuiRegisterMsg( $WM_NCACTIVATE, "WM_NCACTIVATE" ) ; Prevent child windows (for Tab items) in making main GUI title bar dimmed
	GUIRegisterMsg( $WM_NOTIFY,     "WM_NOTIFY" )     ; Double click events in listview on the middle Tab item

	; System image list for listviews
	$hImageList = SystemImageList()

	; Windows Explorer handle
	GUICtrlCreateLabel( "Windows Explorer handle: " & $hExplorer, $xPos, $yPos, $width, 17 )
	$yPos += 17 + 10

	; Current folder
	CreateListView( $xPos+80, $yPos, $width-80, 1, $idLVCurrentFolder, $hLVCurrentFolder, $iLVheight )
	GUICtrlCreateLabel( "Current folder", $xPos, $yPos+2, 70, $iLVheight )
	$yPos += $iLVheight + 10

	; Focused item
	CreateListView( $xPos+80, $yPos, $width-80, 1, $idLVFocusedItem, $hLVFocusedItem, $iLVheight )
	GUICtrlCreateLabel( "Focused item", $xPos, $yPos+2, 70, $iLVheight )
	$yPos += $iLVheight + 20

	; Tab control
	Local $idTab = GUICtrlCreateTab( $xPos, $yPos, $width, $height+60, $TCS_MULTILINE )
	Local $idTab1 = GUICtrlCreateTabItem( "Handle items" )
	Local $idTab2 = GUICtrlCreateTabItem( "Browse folders" )
	Local $idTab3 = GUICtrlCreateTabItem( "Set icon view" )
	GUICtrlCreateTabItem( "" )
	GUISetState()

	; Tab items
	CreateTab1( $hGui, $xPos, $yPos, $width, $height+60 )
	CreateTab2( $hGui, $xPos, $yPos, $width, $height+60 )
	CreateTab3( $hGui, $xPos, $yPos, $width, $height+60 )

	; Main GUI
	GUISwitch( $hGui )

	; Update GUI and Tab1
	IsNewCurrentFolder()
	$fGuiUpdate = 1
	GUICtrlSendToDummy( $idGuiUpdate )
	GUICtrlSendToDummy( $idTab1Update )
	$pPidlAbsHome = ILClone( $pPidlAbsCurFolder )
	$fActivateEvents = True

	Local $iMsg
	While 1

		$iMsg = GUIGetMsg()
		Switch $iMsg

			; Skip events
			Case 0, $GUI_EVENT_MOUSEMOVE
				ContinueLoop ; Jump to "While 1"

			; Switch Tab item
			Case $idTab
				Local $i = GUICtrlRead( $idTab )
				If $i <> $iTab Then SwitchTabItem( $i )

			; Update GUI and Tab item controls
			Case $idGuiUpdate, $idTab1Update, $idTab2Update, $idTab3Update
				HandleUpdateEvents( $iMsg )

			; Control events on Tab1
			Case $idTab1R1, $idTab1R2, $idTab1R3, $idTab1R4, $idTab1R5, $idTab1R6, $idTab1R7, _
			     $idTab1S1, $idTab1S2, $idTab1S3, $idTab1S4, $idTab1S5
			       HandleTab1Events( $iMsg )

			; Control events on Tab2
			Case $idTab2R1, $idTab2R2, $idTab2R3, $idTab2R4, $idTab2R5
				HandleTab2Events( $iMsg )

			; Control events on Tab3
			Case $idTab3RW78 To $idTab3RW78 + 7, _
			     $idTab3RVst To $idTab3RVst + 6, _
			     $idTab3RWXP To $idTab3RWXP + 4
			       HandleTab3Events( $iMsg )

			; Close GUI
			Case $GUI_EVENT_CLOSE
				ExitLoop

		EndSwitch

		; Does Explorer still exists, or is it closed?
		If Not WinExists( $hExplorer ) Then Close()

	WEnd

	Cleanup()
	Exit

EndFunc
