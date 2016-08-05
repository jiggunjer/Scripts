;
; === Functions ===
;
; Get Windows Explorer interface objects
;   GetIShellBrowser
;   GetShellInterfaces
;   IsNewShellFolder
;
; Get/set information in Windows Explorer
;   GetCurrentFolder
;   SetCurrentFolder
;
;   CountItems
;   GetItems
;   GetFiles
;   GetFolders
;   GetPidls
;   GetFocusedItem
;
;   SetFocusedItem
;   SetSelectedItem
;
;   GetIconView
;   SetIconView
;
; Helper functions
;   GetDisplayNameAbs
;   GetDisplayNameRel
;   GetParsingNameRel
;
; === Internal functions ===
;
;   GetExplorerItems
;   GetExplorerItemsByType
;   GetExplorerPidls
;
; === Free memory ===
;
; Depending on parameters most functions can return PIDLs. Some functions return only PIDLs. In these cases
; you must free memory (_WinAPI_CoTaskMemFree) used by the PIDLs, when you have finished using the PIDLs.
;

#include-once

#include "ShellInterfaces.au3"
#include "ShellFunctions.au3"
#include "ContextMenu.au3"
#include "Hourglass.au3"

Global $oIShellBrowser, $pIShellView, $oIShellView, $oIFolderView, $oIFolderView2, $oIPersistFolder2, $oIShellFolder


; --- Get Windows Explorer interface objects ---

; For examples see top of Example.au3

Func GetIShellBrowser( $hExplorer )
	; IShellWindows interface
	Local $pIShellWindows, $oIShellWindows
	CoCreateInstance( $tCLSID_ShellWindows, $NULL, $CLSCTX_ALL, $tRIID_IShellWindows, $pIShellWindows )
	$oIShellWindows = ObjCreateInterface( $pIShellWindows, $sIID_IShellWindows, $dtag_IShellWindows )

	; Number of shell windows
	Local $iWindows
	$oIShellWindows.get_Count( $iWindows )

	; Get an IWebBrowserApp object for each window
	; This is done in two steps:
	; 1. Get an IDispatch object for the window
	; 2. Get the IWebBrowserApp interface
	; Check if it's the right window
	Local $pIDispatch, $oIDispatch
	Local $pIWebBrowserApp, $oIWebBrowserApp, $hWnd
	For $i = 0 To $iWindows - 1
		$oIShellWindows.Item( $i, $pIDispatch )
		If $pIDispatch Then
			$oIDispatch = ObjCreateInterface( $pIDispatch, $sIID_IDispatch, $dtag_IDispatch )
			$oIDispatch.QueryInterface( $tRIID_IWebBrowserApp, $pIWebBrowserApp )
			If $pIWebBrowserApp Then
				$oIWebBrowserApp = ObjCreateInterface( $pIWebBrowserApp, $sIID_IWebBrowserApp, $dtag_IWebBrowserApp )
				$oIWebBrowserApp.get_HWND( $hWnd )
				If $hWnd = $hExplorer Then ExitLoop
			EndIf
		EndIf
	Next

	; IServiceProvider interface
	Local $pIServiceProvider, $oIServiceProvider
	$oIWebBrowserApp.QueryInterface( $tRIID_IServiceProvider, $pIServiceProvider )
	$oIServiceProvider = ObjCreateInterface( $pIServiceProvider, $sIID_IServiceProvider, $dtag_IServiceProvider )

	; IShellBrowser interface
	Local $pIShellBrowser
	$oIServiceProvider.QueryService( $tRIID_STopLevelBrowser, $tRIID_IShellBrowser, $pIShellBrowser )
	$oIShellBrowser = ObjCreateInterface( $pIShellBrowser, $sIID_IShellBrowser, $dtag_IShellBrowser )
EndFunc

Func GetShellInterfaces()
	Local $pIFolderView, $pIFolderView2, $pIPersistFolder2, $pIShellFolder, $pPidlFolder, $pPidlRel, $i = 0

	; IShellView interface
	$oIShellBrowser.QueryActiveShellView( $pIShellView )
	$oIShellView = ObjCreateInterface( $pIShellView, $sIID_IShellView, $dtag_IShellView )

	; IFolderView interface
	$oIShellView.QueryInterface( $tRIID_IFolderView, $pIFolderView )
	$oIFolderView = ObjCreateInterface( $pIFolderView, $sIID_IFolderView, $dtag_IFolderView )

	If @OSVersion <> "WIN_XP" Then
		; IFolderView2 interface (Vista and later)
		$oIShellView.QueryInterface( $tRIID_IFolderView2, $pIFolderView2 )
		$oIFolderView2 = ObjCreateInterface( $pIFolderView2, $sIID_IFolderView2, $dtag_IFolderView2 )
	EndIf

	; IPersistFolder2 interface
	$oIFolderView.GetFolder( $tRIID_IPersistFolder2, $pIPersistFolder2 )
	$oIPersistFolder2 = ObjCreateInterface( $pIPersistFolder2, $sIID_IPersistFolder2, $dtag_IPersistFolder2 )
	$oIPersistFolder2.GetCurFolder( $pPidlFolder )

	; IShellFolder interface
	If ILIsEqual( $pPidlFolder, $pPidlAbsDesktop ) Then
		SHGetDesktopFolder( $pIShellFolder )
	Else
		Local $pIParentFolder, $oIParentFolder, $pPidlRel
		SHBindToParent( $pPidlFolder, DllStructGetPtr( $tRIID_IShellFolder ), $pIParentFolder, $pPidlRel )
		$oIParentFolder = ObjCreateInterface( $pIParentFolder, $sIID_IShellFolder, $dtag_IShellFolder )
		$oIParentFolder.BindToObject( $pPidlRel, $NULL, $tRIID_IShellFolder, $pIShellFolder )
	EndIf
	$oIShellFolder = ObjCreateInterface( $pIShellFolder, $sIID_IShellFolder, $dtag_IShellFolder )

	; Free memory used by $pPidlFolder
	_WinAPI_CoTaskMemFree( $pPidlFolder )

	; Wait for Explorer to refresh
	$pPidlRel = GetFocusedItem()
	While Not $pPidlRel And $i < 10
		Sleep( 25 )
		$pPidlRel = GetFocusedItem()
		$i += 1
	WEnd

	; Free memory used by $pPidlRel
	If $pPidlRel Then _
		_WinAPI_CoTaskMemFree( $pPidlRel )
EndFunc

Func IsNewShellFolder()
	; For a new folder, we'll get a new
	; pointer for the IShellView interface.

	Local $pIShellViewNew, $i = 0
	$oIShellBrowser.QueryActiveShellView( $pIShellViewNew )

	; Wait for Explorer to refresh
	While $pIShellViewNew = $pIShellView And $i < 10
		Sleep( 10 )
		$oIShellBrowser.QueryActiveShellView( $pIShellViewNew )
		$i += 1
	WEnd

	Return ( $pIShellViewNew <> $pIShellView )
EndFunc


; --- Get/set information in Windows Explorer ---

; For examples search the functions in Example.au3 and Example\*.au3.

Func GetCurrentFolder()
	Local $pPidlAbs
	$oIPersistFolder2.GetCurFolder( $pPidlAbs )
	Return $pPidlAbs
EndFunc

; After this command $oIShellBrowser is the only valid interface object. To
; be able to use the other interfaces you must execute GetShellInterfaces().
Func SetCurrentFolder( $pPidl, $fFlag )
	$oIShellBrowser.BrowseObject( $pPidl, BitOR( $SBSP_DEFBROWSER, $fFlag ) )
EndFunc


Func CountItems( $fSelected = False )
	Local $iCount, $fSelection = $fSelected ? $SVGIO_SELECTION : $SVGIO_ALLVIEW
	$oIFolderView.ItemCount( $fSelection, $iCount )
	Return $iCount
EndFunc

Func GetItems( $fSelected = False, $fFullPath = False, $fPidl = False, $iMax = 0 )
	Local $fSelection = $fSelected ? $SVGIO_SELECTION : $SVGIO_ALLVIEW
	Local $aItems, $fItemName = $fFullPath ? $SHGDN_FORPARSING : $SHGDN_NORMAL
	GetExplorerItems( $oIFolderView, $oIShellFolder, $fSelection, $fItemName, $aItems, $fFullPath, $fPidl, $iMax )
	Return $aItems
EndFunc

Func GetFiles( $fSelected = False, $fFullPath = False, $fPidl = False, $iMax = 0 )
	Local $fSelection = $fSelected ? $SVGIO_SELECTION : $SVGIO_ALLVIEW
	Local $aItems, $fItemName = $fFullPath ? $SHGDN_FORPARSING : $SHGDN_NORMAL
	GetExplorerItemsByType( $oIFolderView, $oIShellFolder, $fSelection, $fItemName, $aItems, True, $fFullPath, $fPidl, $iMax )
	Return $aItems
EndFunc

Func GetFolders( $fSelected = False, $fFullPath = False, $fPidl = False, $iMax = 0 )
	Local $fSelection = $fSelected ? $SVGIO_SELECTION : $SVGIO_ALLVIEW
	Local $aItems, $fItemName = $fFullPath ? $SHGDN_FORPARSING : $SHGDN_NORMAL
	GetExplorerItemsByType( $oIFolderView, $oIShellFolder, $fSelection, $fItemName, $aItems, False, $fFullPath, $fPidl, $iMax )
	Return $aItems
EndFunc

; This function is much faster than the three functions above, because the names are not calculated.
; The function does not distinguish between files and folders, it gets all pidls or all selected pidls.
Func GetPidls( $fSelected = False )
	Local $aPidls, $fSelection = $fSelected ? $SVGIO_SELECTION : $SVGIO_ALLVIEW
	GetExplorerPidls( $oIFolderView, $oIShellFolder, $fSelection, $aPidls )
	Return $aPidls
EndFunc

Func GetFocusedItem()
	Local $iFocus, $pPidlRel
	$oIFolderView.GetFocusedItem( $iFocus )
	If $iFocus = -1 Then Return 0
	$oIFolderView.Item( $iFocus, $pPidlRel )
	Return $pPidlRel
EndFunc


Func SetFocusedItem( $item, $fPidl = False )
	If $fPidl Then
		$oIShellView.SelectItem( $item, $SVSI_FOCUSED )  ; $item = $pPidlRel
	Else
		$oIFolderView.SelectItem( $item, $SVSI_FOCUSED ) ; $item = $iIndex
	EndIf
EndFunc

Func SetSelectedItem( $item, $fSelected = True, $fPidl = False )
	Local $fSelection = $fSelected ? $SVSI_SELECT : $SVSI_DESELECT
	If $fPidl Then
		$oIShellView.SelectItem( $item, $fSelection )  ; $item = $pPidlRel
	Else
		$oIFolderView.SelectItem( $item, $fSelection ) ; $item = $iIndex
	EndIf
EndFunc


Func GetIconView()
	Local $iViewMode, $iIconSize
	If @OSVersion = "WIN_XP" Then
		$oIFolderView.GetCurrentViewMode( $iViewMode )
		Return $iViewMode
	Else
		$oIFolderView2.GetViewModeAndIconSize( $iViewMode, $iIconSize )
		Local $aView = [ $iViewMode, $iIconSize ]
		Return $aView
	EndIf
EndFunc

Func SetIconView( $iViewMode, $iIconSize = 0 )
	If @OSVersion <> "WIN_XP" Then
		$oIFolderView.SetCurrentViewMode( $iViewMode )
	Else
		$oIFolderView2.SetViewModeAndIconSize( $iViewMode, $iIconSize )
	EndIf
EndFunc


; --- Helper functions ---

; For examples search Example.au3 and Example\*.au3.

; Returns a nice name (not full path) relative to Desk-
; top - without GUID-strings for system files/folders.
; $pPidlAbs must be a child PIDL relative to Desktop.
; (A child PIDL relative to Desktop is also absolut.)
Func GetDisplayNameAbs( $pPidlAbs )
	Local Static $pIDesktopFolder, $oIDesktopFolder
	If Not $pIDesktopFolder Then
		SHGetDesktopFolder( $pIDesktopFolder )
		$oIDesktopFolder = ObjCreateInterface( $pIDesktopFolder, $sIID_IShellFolder, $dtag_IShellFolder )
	EndIf
	; Note that this code is the same as the code below,
	; just $oIDesktopFolder instead of $oIShellFolder.
	Local $tSTRRET = DllStructCreate( $tagSTRRET ), $sName
	$oIDesktopFolder.GetDisplayNameOf( $pPidlAbs, $SHGDN_NORMAL, $tSTRRET )
	StrRetToBuf( DllStructGetPtr( $tSTRRET ), $NULL, $sName )
	Return $sName
EndFunc

; Returns a nice name (not full path) relative to current
; folder - without GUID-strings for system files/folders.
Func GetDisplayNameRel( $pPidlRel )
	; Note that this code is the same as the code above,
	; just $oIShellFolder instead of $oIDesktopFolder.
	Local $tSTRRET = DllStructCreate( $tagSTRRET ), $sName
	$oIShellFolder.GetDisplayNameOf( $pPidlRel, $SHGDN_NORMAL, $tSTRRET )
	StrRetToBuf( DllStructGetPtr( $tSTRRET ), $NULL, $sName )
	Return $sName
EndFunc

; Returns the full path for a file/folder. Can
; include GUID-strings for system files/folders.
Func GetParsingNameRel( $pPidlRel )
	; Note that this code is the same as the code above,
	; just $SHGDN_FORPARSING instead of $SHGDN_NORMAL.
	Local $tSTRRET = DllStructCreate( $tagSTRRET ), $sName
	$oIShellFolder.GetDisplayNameOf( $pPidlRel, $SHGDN_FORPARSING, $tSTRRET )
	StrRetToBuf( DllStructGetPtr( $tSTRRET ), $NULL, $sName )
	Return $sName
EndFunc


; --- Internal functions ---

; You can call these functions directly, but you have to take care of all the parameters, which includes
; interface objects and constants like $SVGIO_ALLVIEW, $SVGIO_SELECTION, $SHGDN_NORMAL and $SHGDN_FORPARSING.

; Retrieve all/selected items
; $fSelection = $SVGIO_ALLVIEW/$SVGIO_SELECTION
; $fName = $SHGDN_NORMAL/$SHGDN_FORPARSING
Func GetExplorerItems( $oIFolderView, $oIShellFolder, $fSelection, $fName, ByRef $aItems, $fFullPath = True, $fPidl = False, $iMax = 0 )
	; Number of items
	Local $iItems
	$fSelection = BitOR( $fSelection, $SVGIO_FLAG_VIEWORDER )
	$oIFolderView.ItemCount( $fSelection, $iItems )
	If $iItems = 0 Then Return 0
	Hourglass( True )

	If $iMax Then
		If $iItems > $iMax Then
			$iItems = $iMax
		Else
			$iMax = $iItems
		EndIf
	Else
		$iMax = $iItems
	EndIf

	; Enumeration object
	Local $pIEnumIDList, $oIEnumIDList
	$oIFolderView.Items( $fSelection, $tRIID_IEnumIDList, $pIEnumIDList )
	$oIEnumIDList = ObjCreateInterface( $pIEnumIDList, $sIID_IEnumIDList, $dtag_IEnumIDList )

	; Name format
	If Not $fFullPath Then
		If $fName = $SHGDN_NORMAL Then
			$fName = $SHGDN_INFOLDER
		Else ; $fName = $SHGDN_FORPARSING
			$fName = BitOr( $SHGDN_INFOLDER, $SHGDN_FORPARSING )
		EndIf
	EndIf

	Local $aItemsAr[$iMax], $pidlRel, $iFetched, $n = 0
	If $fPidl Then ReDim $aItemsAr[$iMax][2]
	Local $tSTRRET = DllStructCreate( $tagSTRRET ), $sName
	While $oIEnumIDList.Next( 1, $pidlRel, $iFetched ) = $S_OK And $n < $iMax
		$oIShellFolder.GetDisplayNameOf( $pidlRel, $fName, $tSTRRET )
		StrRetToBuf( DllStructGetPtr( $tSTRRET ), $NULL, $sName )
		If $fPidl Then
			$aItemsAr[$n][0] = $sName
			$aItemsAr[$n][1] = $pidlRel
		Else
			_WinAPI_CoTaskMemFree( $pidlRel )
			$aItemsAr[$n] = $sName
		EndIf
		$n += 1
	WEnd

	Hourglass( False )
	$aItems = $aItemsAr
	Return $iItems
EndFunc

; Retrieve all/selected items by type
; $fSelection = $SVGIO_ALLVIEW/$SVGIO_SELECTION
; $fName = $SHGDN_NORMAL/$SHGDN_FORPARSING
; Get items by type: $fFiles = True/False
Func GetExplorerItemsByType( $oIFolderView, $oIShellFolder, $fSelection, $fName, ByRef $aItems, $fFiles = True, $fFullPath = True, $fPidl = False, $iMax = 0 )
	; Number of items
	Local $iItems
	$fSelection = BitOR( $fSelection, $SVGIO_FLAG_VIEWORDER )
	$oIFolderView.ItemCount( $fSelection, $iItems )
	If $iItems = 0 Then Return 0
	Hourglass( True )

	If $iMax Then
		If $iItems > $iMax Then
			$iItems = $iMax
		Else
			$iMax = $iItems
		EndIf
	Else
		$iMax = $iItems
	EndIf

	; Enumeration object
	Local $pIEnumIDList, $oIEnumIDList
	$oIFolderView.Items( $fSelection, $tRIID_IEnumIDList, $pIEnumIDList )
	$oIEnumIDList = ObjCreateInterface( $pIEnumIDList, $sIID_IEnumIDList, $dtag_IEnumIDList )

	; Name format
	If Not $fFullPath Then
		If $fName = $SHGDN_NORMAL Then
			$fName = $SHGDN_INFOLDER
		Else ; $fName = $SHGDN_FORPARSING
			$fName = BitOr( $SHGDN_INFOLDER, $SHGDN_FORPARSING )
		EndIf
	EndIf

	; Get Explorer items
	Local $aItemsAr[$iMax], $pidlRel, $iFetched, $n = 0
	If $fPidl Then ReDim $aItemsAr[$iMax][2]
	Local $tSTRRET = DllStructCreate( $tagSTRRET ), $sName
	Local $tArray = DllStructCreate( "ptr" ), $iAttribs
	While $oIEnumIDList.Next( 1, $pidlRel, $iFetched ) = $S_OK And $n < $iMax
		DllStructSetData( $tArray, 1, $pidlRel )
		$iAttribs = BitOR( $SFGAO_FOLDER, $SFGAO_STREAM )
		$oIShellFolder.GetAttributesOf( 1, $tArray, $iAttribs )
		If BitAND( $iAttribs, $SFGAO_FOLDER ) And Not BitAND( $iAttribs, $SFGAO_STREAM ) Then
			; Item is a folder
			If $fFiles Then
				_WinAPI_CoTaskMemFree( $pidlRel )
				ContinueLoop
			EndIf
		Else
			; Item is a file
			If Not $fFiles Then
				_WinAPI_CoTaskMemFree( $pidlRel )
				ContinueLoop
			EndIf
		EndIf
		$oIShellFolder.GetDisplayNameOf( $pidlRel, $fName, $tSTRRET )
		StrRetToBuf( DllStructGetPtr( $tSTRRET ), $NULL, $sName )
		If $fPidl Then
			$aItemsAr[$n][0] = $sName
			$aItemsAr[$n][1] = $pidlRel
		Else
			_WinAPI_CoTaskMemFree( $pidlRel )
			$aItemsAr[$n] = $sName
		EndIf
		$n += 1
	WEnd

	Hourglass( False )
	If $fPidl Then
		ReDim $aItemsAr[$n][2]
	Else
		ReDim $aItemsAr[$n]
	EndIf
	$aItems = $aItemsAr
	Return $iItems
EndFunc

; Retrieve all/selected pidls
; $fSelection = $SVGIO_ALLVIEW/$SVGIO_SELECTION
; This function is much faster than the two functions above, because the names are not calculated
Func GetExplorerPidls( $oIFolderView, $oIShellFolder, $fSelection, ByRef $aPidls )
	; Number of pidls
	Local $iPidls
	$fSelection = BitOR( $fSelection, $SVGIO_FLAG_VIEWORDER )
	$oIFolderView.ItemCount( $fSelection, $iPidls )
	If $iPidls = 0 Then Return 0
	Hourglass( True )

	; Enumeration object
	Local $pIEnumIDList, $oIEnumIDList
	$oIFolderView.Items( $fSelection, $tRIID_IEnumIDList, $pIEnumIDList )
	$oIEnumIDList = ObjCreateInterface( $pIEnumIDList, $sIID_IEnumIDList, $dtag_IEnumIDList )

	Local $aPidlsAr[$iPidls], $pidlRel, $iFetched, $n = 0
	While $oIEnumIDList.Next( 1, $pidlRel, $iFetched ) = $S_OK
		$aPidlsAr[$n] = $pidlRel
		$n += 1
	WEnd

	Hourglass( False )
	$aPidls = $aPidlsAr
	Return $iPidls
EndFunc
