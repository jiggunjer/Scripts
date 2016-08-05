#include-once

; --- Create child windows with controls for Tab items ---

; The child window fills out the free area of a Tab item
; It acts as a container for the controls on the Tab item
; This ensures that controls are properly refreshed when necessary
; And it's easy to show/hide all controls when a tab is active/inactive

Func CreateTab1( $hGui, $xPos0, $yPos0, $width, $height ) ; Pos and size of Tab control
	$hTab1 = GUICreate( "", $width-4, $height-24, $xPos0+1, $yPos0+21, $WS_POPUP, $WS_EX_MDICHILD, $hGui )
	$idTab1Update = GUICtrlCreateDummy()
	GUISetBkColor( 0xFFFFFF )
	Local $xPos = 7, $yPos = 15

	;Local $sCount = " (" & CountItems() & ")"
	$idTab1G1 = GUICtrlCreateGroup( "Get items", $xPos, $yPos-1, 120, 100 )
	$idTab1R1 = GUICtrlCreateRadio( "All",     $xPos+10, $yPos+20, 120-20, 20 )
	$idTab1R2 = GUICtrlCreateRadio( "Files",   $xPos+10, $yPos+45, 120-20, 20 )
	$idTab1R3 = GUICtrlCreateRadio( "Folders", $xPos+10, $yPos+70, 120-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	$xPos += 140
	;$sCount = " (" & CountItems( True ) & ")"
	$idTab1G2 = GUICtrlCreateGroup( "Get selected", $xPos, $yPos-1, 120, 100 )
	$idTab1R4 = GUICtrlCreateRadio( "All",     $xPos+10, $yPos+20, 120-20, 20 )
	$idTab1R5 = GUICtrlCreateRadio( "Files",   $xPos+10, $yPos+45, 120-20, 20 )
	$idTab1R6 = GUICtrlCreateRadio( "Folders", $xPos+10, $yPos+70, 120-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	$xPos += 140
	GUICtrlCreateGroup( "Open selected items", $xPos, $yPos-1, 120, 100 )
	$idTab1R7 = GUICtrlCreateRadio( "Open", $xPos+10, $yPos+20, 120-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	$xPos += 140
	Local $iLVheight
	Local $rows = ( @OSVersion = "WIN_XP" ) ? 12 : 11
	CreateListView( $xPos, $yPos+5, $width-4-$xPos-7, $rows, $idLVFilesFolders, $hLVFilesFolders, $iLVheight )

	$xPos = 7
	$yPos = $iLVheight-80
	GUICtrlCreateGroup( "Set items selected", $xPos, $yPos, 120, 100 )
	$idTab1S1 = GUICtrlCreateRadio( "All",     $xPos+10, $yPos+20, 120-20, 20 )
	$idTab1S2 = GUICtrlCreateRadio( "None",    $xPos+10, $yPos+45, 120-20, 20 )
	$idTab1S3 = GUICtrlCreateRadio( "Reverse", $xPos+10, $yPos+70, 120-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	$xPos += 140
	GUICtrlCreateGroup( "Set items focused or selected", $xPos, $yPos, 2*140-20, 100 )
	GUICtrlCreateLabel( "Select one/more items in listview and click button:",        $xPos+10, $yPos+20, 2*140-20-20, 20 )
	$idTab1S4 = GUICtrlCreateRadio( "Set focused item (first item if more selected)", $xPos+10, $yPos+45, 2*140-20-20, 20 )
	$idTab1S5 = GUICtrlCreateRadio( "Set selected items (add items to selection)",    $xPos+10, $yPos+70, 2*140-20-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	GUISetState()
EndFunc

Func CreateTab2( $hGui, $xPos0, $yPos0, $width, $height ) ; Pos and size of Tab control
	$hTab2 = GUICreate( "", $width-4, $height-24, $xPos0+1, $yPos0+21, $WS_POPUP, $WS_EX_MDICHILD, $hGui )
	$idTab2Update = GUICtrlCreateDummy()
	GUISetBkColor( 0xFFFFFF )
	Local $xPos = 7, $yPos = 15

	$xPos += 2*140
	Local $iLVheight
	Local $rows = ( @OSVersion = "WIN_XP" ) ? 12 : 11
	Local $iStyle = $LVS_SINGLESEL + $LVS_NOCOLUMNHEADER
	CreateListView( $xPos, $yPos+5, $width-4-$xPos-7, $rows, $idLVFolders, $hLVFolders, $iLVheight, $iStyle )
	$iGheight = $iLVheight

	$xPos -= 2*140
	GUICtrlCreateGroup( "Browse to specific folder", $xPos, $yPos-1, 2*140-20, 100 )
	$idTab2R1 = GUICtrlCreateRadio( "Desktop folder",      $xPos+10, $yPos+20, 2*140-20-20, 20 )
	$idTab2R2 = GUICtrlCreateRadio( "Computer folder",     $xPos+10, $yPos+45, 2*140-20-20, 20 )
	$idTab2R3 = GUICtrlCreateRadio( "Home/startup folder", $xPos+10, $yPos+70, 2*140-20-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	$yPos = $iLVheight-80
	GUICtrlCreateGroup( "Browse to relative folder", $xPos, $yPos, 2*140-20, 100 )
	$idTab2R4 = GUICtrlCreateRadio( "Parent folder",                               $xPos+10, $yPos+20, 2*140-20-20, 20 )
	GUICtrlCreateLabel( "Select child folder in listview and click radio button:", $xPos+10, $yPos+45, 2*140-20-20, 20 )
	$idTab2R5 = GUICtrlCreateRadio( "Child folder",                                $xPos+10, $yPos+70, 2*140-20-20, 20 )
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )
EndFunc

Func CreateTab3( $hGui, $xPos0, $yPos0, $width, $height ) ; Pos and size of Tab control
	$hTab3 = GUICreate( "", $width-4, $height-24, $xPos0+1, $yPos0+21, $WS_POPUP, $WS_EX_MDICHILD, $hGui )
	$idTab3Update = GUICtrlCreateDummy()
	GUISetBkColor( 0xFFFFFF )
	Local $xPos = 7, $yPos = 15

	Local $aTexts = [ "Medium size icons (48x48)", "Small icons (16x16)",  "List view (16x16)",   "Details view (16x16)", _
	                  "Tiles view (48x48)",        "Content view (32x32)", "Large icons (96x96)", "Extra large icons (256x256)" ]

	; --- Windows 7, 8 ---

	$width = ( $width - 4 - $xPos - 7 - 40 ) / 3
	GUICtrlCreateGroup( "Windows 7, 8", $xPos, $yPos-1, $width, $iGheight+6 )
	Local $fDisable = ( @OSVersion = "WIN_XP" Or @OSVersion = "WIN_VISTA" ) * $GUI_DISABLE
	$idTab3RW78 = GUICtrlCreateRadio( $aTexts[$aW78ViewModeIndex[0]],  $xPos+10, $yPos+20,       $width-20, 20 )
	If $fDisable Then GUICtrlSetState( $idTab3RW78, $fDisable )
	For $i = 1 To 7
	              GUICtrlCreateRadio( $aTexts[$aW78ViewModeIndex[$i]], $xPos+10, $yPos+20+$i*23, $width-20, 20 )
	              If $fDisable Then GUICtrlSetState( $idTab3RW78+$i, $fDisable )
	Next
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	; --- Windows Vista ---

	$xPos += $width+20
	GUICtrlCreateGroup( "Windows Vista", $xPos, $yPos-1, $width, $iGheight+6 )
	$fDisable = ( @OSVersion <> "WIN_VISTA" ) * $GUI_DISABLE
	$idTab3RVst = GUICtrlCreateRadio( $aTexts[$aVstViewModeIndex[0]],  $xPos+10, $yPos+20,       $width-20, 20 )
	If $fDisable Then GUICtrlSetState( $idTab3RVst, $fDisable )
	For $i = 1 To 6
	              GUICtrlCreateRadio( $aTexts[$aVstViewModeIndex[$i]], $xPos+10, $yPos+20+$i*23, $width-20, 20 )
	              If $fDisable Then GUICtrlSetState( $idTab3RVst+$i, $fDisable )
	Next
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )

	; --- Windows XP ---

	Local $aTextsXP = [ "Icons view (32x32)", "List view (16x16)", "Details view (16x16)", "Thumbnail icons", "Tiles view (48x48)" ]

	$xPos += $width+20
	GUICtrlCreateGroup( "Windows XP", $xPos, $yPos-1, $width+1, $iGheight+6 )
	$fDisable = ( @OSVersion <> "WIN_XP" ) * $GUI_DISABLE
	$idTab3RWXP = GUICtrlCreateRadio( $aTextsXP[$aWXPViewModeIndex[0]],  $xPos+10, $yPos+20,       $width-20, 20 )
	If $fDisable Then GUICtrlSetState( $idTab3RWXP, $fDisable )
	For $i = 1 To 4
	              GUICtrlCreateRadio( $aTextsXP[$aWXPViewModeIndex[$i]], $xPos+10, $yPos+20+$i*23, $width-20, 20 )
	              If $fDisable Then GUICtrlSetState( $idTab3RWXP+$i, $fDisable )
	Next
	GUICtrlCreateGroup( "", -99, -99, 1, 1 )
EndFunc
