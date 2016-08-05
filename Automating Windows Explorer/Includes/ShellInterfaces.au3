#include-once

#include <WinAPI.au3>


; === IColumnManager interface (Vista and later) ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb776149(v=vs.85).aspx

; Exposes methods that enable inspection and manipulation of columns in the Microsoft Windows Explorer
; Details view. Each column is referenced by a PROPERTYKEY structure, which names a property.

Global Const $sIID_IColumnManager  = "{D8EC27BB-3F3B-4042-B10A-4ACFD924D453}"
Global Const $tRIID_IColumnManager = _WinAPI_GUIDFromString( $sIID_IColumnManager )
Global Const $dtag_IColumnManager  = _
	"SetColumnInfo hresult(struct*;struct*);" & _ ; Sets the state for a specified column.
	"GetColumnInfo hresult(struct*;struct*);" & _ ; Gets information about each column: width, visibility, display name, and state.
	"GetColumnCount hresult(dword;uint*);" & _    ; Gets the column count for either the visible columns or the complete set of columns.
	"GetColumns hresult(dword;struct*;uint);" & _ ; Gets an array of PROPERTYKEY structures that represent the columns that the view supports. Includes either all columns or only those currently visible.
	"SetColumns hresult(struct*;uint);"           ; Sets the collection of columns for the view to display.

; CM_COLUMNINFO structure
Global Const $tagCM_COLUMNINFO = _
	"dword cbSize;" & _        ; The size of the structure, in bytes.
	"dword dwMask;" & _        ; One or more values from the CM_MASK enumeration that specify which members of this structure are valid.
	"dword dwState;" & _       ; One or more values from the CM_STATE enumeration that specify the state of the column.
	"uint  uWidth;" & _        ; One of the members of the CM_SET_WIDTH_VALUE enumeration that specifies the column width.
	"uint  uDefaultWidth;" & _ ; The default width of the column.
	"uint  uIdealWidth;" & _   ; The ideal width of the column.
	"wchar wszName[80]"        ; A buffer of size MAX_COLUMN_NAME_LEN that contains the name of the column as a null-terminated Unicode string.

; PROPERTYKEY Structure
Global Const $tagPROPERTYKEY = $tagGUID & ";dword pid"

; Flags to specify which set of columns are being requested
Global Const $CM_ENUM_ALL                 = 0x00000001
Global Const $CM_ENUM_VISIBLE             = 0x00000002

; CM_MASK constants
Global Const $CM_MASK_WIDTH               = 0x00000001
Global Const $CM_MASK_DEFAULTWIDTH        = 0x00000002
Global Const $CM_MASK_IDEALWIDTH          = 0x00000004
Global Const $CM_MASK_NAME                = 0x00000008
Global Const $CM_MASK_STATE               = 0x00000010

; CM_STATE constants
Global Const $CM_STATE_NONE               = 0x00000000
Global Const $CM_STATE_VISIBLE            = 0x00000001
Global Const $CM_STATE_FIXEDWIDTH         = 0x00000002
Global Const $CM_STATE_NOSORTBYFOLDERNESS = 0x00000004

; CM_SET_WIDTH_VALUE constants
Global Const $CM_WIDTH_USEDEFAULT         = -0x00000001
Global Const $CM_WIDTH_AUTOSIZE           = -0x00000002


; === ICommDlgBrowser interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb776141(v=vs.85).aspx

; ICommDlgBrowser is exposed by the common file dialog boxes to be used when they host a Shell browser.
; If supported, ICommDlgBrowser exposes methods that allow a Shell view to handle several cases that
; require different behavior in a dialog box than in a normal Shell view. You obtain an ICommDlgBrowser
; interface pointer by calling QueryInterface on the IShellBrowser object.

Global Const $sIID_ICommDlgBrowser  = "{000214F1-0000-0000-C000-000000000046}"
Global Const $tRIID_ICommDlgBrowser = _WinAPI_GUIDFromString( $sIID_ICommDlgBrowser )
Global Const $dtag_ICommDlgBrowser  = _
	"OnDefaultCommand hresult(ptr);" & _    ; Called when a user double-clicks in the view or presses the ENTER key.
	"OnStateChange hresult(ptr;ulong);" & _ ; Called after a state, identified by the uChange parameter, has changed in the IShellView interface.
	"IncludeObject hresult(ptr;ptr);"       ; Allows the common dialog box to filter objects that the view displays.


; === IContextMenu interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb776095(v=vs.85).aspx

; Exposes methods that either create or merge a shortcut menu associated with a Shell object.

Global Const $sIID_IContextMenu  = "{000214E4-0000-0000-C000-000000000046}"
Global Const $tRIID_IContextMenu = _WinAPI_GUIDFromString( $sIID_IContextMenu )
Global Const $dtag_IContextMenu  = _
	"QueryContextMenu hresult(handle;uint;uint;uint;uint);" & _ ; Adds commands to a shortcut menu.
	"InvokeCommand hresult(struct*);" & _                       ; Carries out the command associated with a shortcut menu item.
	"GetCommandString hresult(uint;uint;uint*;ptr;uint);"       ; Gets information about a shortcut menu command, including the help string and the language-independent, or canonical, name for the command.

Global Const $CMF_NORMAL        = 0x00000000
Global Const $CMF_DEFAULTONLY   = 0x00000001
Global Const $CMF_EXTENDEDVERBS = 0x00000100
Global Const $CMF_EXPLORE       = 0x00000004

Global Const $SW_SHOWNORMAL = 1

; CMINVOKECOMMANDINFO structure
Global Const $tagCMINVOKECOMMANDINFO = "dword cbSize;dword fMask;hwnd hWnd;ptr lpVerb;ptr lpParameters;" & _
                                       "ptr lpDirectory;int nShow;dword dwHotKey;handle hIcon"


; === IDispatch interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/ms221608(v=vs.85).aspx

; Exposes objects, methods and properties to programming tools and other applications that support Automation.
; COM components implement the IDispatch interface to enable access by Automation clients, such as Visual Basic. 

Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Global Const $dtag_IDispatch = _
	"GetTypeInfoCount hresult(dword*);" & _                 ; Retrieves the number of type information interfaces that an object provides (either 0 or 1).
	"GetTypeInfo hresult(dword;dword;ptr*);" & _            ; Gets the type information for an object.
	"GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _   ; Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
	"Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);" ; Provides access to properties and methods exposed by an object.


; === IEnumIDList interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb761982(v=vs.85).aspx

; Exposes a standard set of methods used to enumerate the pointers to item identifier lists (PIDLs) of the items in a Shell folder.
; When a folder's IShellFolder::EnumObjects method is called, it creates an enumeration object and passes a pointer to the object's
; IEnumIDList interface back to the calling application.

Global Const $sIID_IEnumIDList  = "{000214F2-0000-0000-C000-000000000046}"
Global Const $tRIID_IEnumIDList = _WinAPI_GUIDFromString( $sIID_IEnumIDList )
Global Const $dtag_IEnumIDList  = _
	"Next hresult(ulong;ptr*;ulong*);" & _ ; Retrieves the specified number of item identifiers in the enumeration sequence and advances the current position by the number of items retrieved.
	"Reset hresult();" & _                 ; Returns to the beginning of the enumeration sequence.
	"Skip hresult(ulong);" & _             ; Skips the specified number of elements in the enumeration sequence.
	"Clone hresult(ptr*);"                 ; Creates a new item enumeration object with the same contents and state as the current one.


; === IExplorerBrowser interface (Vista and later) ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb761909(v=vs.85).aspx

; IExplorerBrowser is a browser object that can either be navigated or host a view of a data object.
; As a full-featured browser object, IExplorerBrowser also supports an automatic travel log.

If @AutoItX64 Then
	Global $sSetRect = "SetRect hresult(ptr;struct*);"
Else
	Global $sSetRect = "SetRect hresult(ptr;long;long;long;long);"
EndIf

Global Const $sCLSID_ExplorerBrowser = "{71F96385-DDD6-48D3-A0C1-AE06E8B055FB}"
Global Const $tCLSID_ExplorerBrowser = _WinAPI_GUIDFromString( $sCLSID_ExplorerBrowser )
Global Const $sIID_IExplorerBrowser  = "{DFD3B6B5-C10C-4BE9-85F6-A66969F402F6}"
Global Const $tRIID_IExplorerBrowser = _WinAPI_GUIDFromString( $sIID_IExplorerBrowser )
Global Const $dtag_IExplorerBrowser  = _
	"Initialize hresult(hwnd;struct*;struct*);" & _ ; Prepares the browser to be navigated.
	"Destroy hresult();" & _                        ; Destroys the browser.
	$sSetRect & _                                   ; Sets the size and position of the view windows created by the browser.
	"SetPropertyBag hresult(ptr);" & _              ; Sets the name of the property bag.
	"SetEmptyText hresult(ptr);" & _                ; Sets the default empty text.
	"SetFolderSettings hresult(ptr);" & _           ; Sets the folder settings for the current view.
	"Advise hresult(ptr;dword*);" & _               ; Initiates a connection with IExplorerBrowser for event callbacks.
	"Unadvise hresult(dword);" & _                  ; Terminates an advisory connection.
	"SetOptions hresult(dword);" & _                ; Sets the current browser options.
	"GetOptions hresult(dword*);" & _               ; Gets the current browser options.
	"BrowseToIDList hresult(ptr;uint);" & _         ; Browses to a pointer to an item identifier list (PIDL).
	"BrowseToObject hresult(ptr;uint);" & _         ; Browses to an object.
	"FillFromObject hresult(ptr;dword);" & _        ; Creates a results folder and fills it with items.
	"RemoveAll hresult();" & _                      ; Removes all items from the results folder.
	"GetCurrentView hresult(struct*;ptr*);"         ; Gets an interface for the current view of the browser.

; --- EXPLORER_BROWSER_OPTIONS for SetOptions and GetOptions dwFlag ---

Global Const $EBO_NONE               = 0x00000000 ; No options.
Global Const $EBO_NAVIGATEONCE       = 0x00000001 ; Do not navigate further than the initial navigation.
Global Const $EBO_SHOWFRAMES         = 0x00000002 ; Use the following standard panes: Commands Module pane, Navigation pane, Details pane, and Preview pane.
Global Const $EBO_ALWAYSNAVIGATE     = 0x00000004 ; Always navigate, even if you are attempting to navigate to the current folder.
Global Const $EBO_NOTRAVELLOG        = 0x00000008 ; Do not update the travel log.
Global Const $EBO_NOWRAPPERWINDOW    = 0x00000010 ; Do not use a wrapper window. This flag is used with legacy clients that need the browser parented directly on themselves.
Global Const $EBO_HTMLSHAREPOINTVIEW = 0x00000020 ; Show WebView for sharepoint sites.


; === IExplorerBrowserEvents interface (Vista and later) ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb761883(v=vs.85).aspx

; Exposes methods for notification of Explorer browser navigation and view creation events.

Global Const $sIID_IExplorerBrowserEvents = "{361BBDC7-E6EE-4E13-BE58-58E2240C810F}"
Global Const $dtag_IExplorerBrowserEvents = _
	"OnNavigationPending hresult(ptr);" & _  ; Notifies clients of a pending Explorer browser navigation to a Shell folder.
	"OnViewCreated hresult(ptr);" & _        ; Notifies clients that the view of the Explorer browser has been created and can be modified.
	"OnNavigationComplete hresult(ptr);" & _ ; Notifies clients that the Explorer browser has successfully navigated to a Shell folder.
	"OnNavigationFailed hresult(ptr);"       ; Notifies clients that the Explorer browser has failed to navigate to a Shell folder.


; === IExplorerPaneVisibility interface (Vista and later) ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb761858(v=vs.85).aspx

Global Const $sIID_IExplorerPaneVisibility  = "{E07010EC-BC17-44C0-97B0-46C7C95B9EDC}"
Global Const $tRIID_IExplorerPaneVisibility = _WinAPI_GUIDFromString( $sIID_IExplorerPaneVisibility )
Global Const $dtag_IExplorerPaneVisibility  = _
	"GetPaneState hresult(ptr;uint*);" ; Gets the visibility state of the given Windows Explorer pane.


; === IFolderFilter interface ===

Global Const $sIID_IFolderFilter = "{9CC22886-DC8E-11D2-B1D0-00C04F8EEB3E}"
Global Const $dtag_IFolderFilter = _
	"ShouldShow hresult(ptr;ptr;ptr);" & _
	"GetEnumFlags hresult(ptr;ptr;hwnd;dword*);"


; === IFolderFilterSite interface ===

Global Const $sIID_IFolderFilterSite  = "{C0A651F5-B48B-11D2-B5ED-006097C686F6}"
Global Const $tRIID_IFolderFilterSite = _WinAPI_GUIDFromString( $sIID_IFolderFilterSite )
Global Const $dtag_IFolderFilterSite  = "SetFilter hresult(ptr);"


; === IFolderView interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775606(v=vs.85).aspx

; Exposes methods that retrieve information about a folder's display options,
; select specified items in that folder, and set the folder's view mode.

Global Const $sIID_IFolderView  = "{CDE725B0-CCC9-4519-917E-325D72FAB4CE}"
Global Const $tRIID_IFolderView = _WinAPI_GUIDFromString( $sIID_IFolderView )
Global Const $dtag_IFolderView  = _
	"GetCurrentViewMode hresult(uint*);" & _              ; Gets an address containing a value representing the folder's current view mode.
	"SetCurrentViewMode hresult(uint);" & _               ; Sets the selected folder's view mode.
	"GetFolder hresult(clsid;ptr*);" & _                  ; Gets the folder object.
	"Item hresult(int;ptr*);" & _                         ; Gets the identifier of a specific item in the folder view, by index.
	"ItemCount hresult(uint;int*);" & _                   ; Gets a pointer to the number of items in the folder.
	"Items hresult(uint;clsid;ptr*);" & _                 ; Gets the address of an enumeration object based on the collection of items in the folder view.
	"GetSelectionMarkedItem hresult(int*);" & _           ; Gets the index of an item in the folder's view which has been marked by using the SVSI_SELECTIONMARK in IFolderView::SelectItem.
	"GetFocusedItem hresult(int*);" & _                   ; Gets the index of the item that currently has focus in the folder's view.
	"GetItemPosition hresult(ptr;ptr);" & _               ; Gets the position of an item in the folder's view.
	"GetSpacing hresult(ptr);" & _                        ; Gets a POINT structure containing the width (x) and height (y) dimensions, including the surrounding white space, of an item.
	"GetDefaultSpacing hresult(ptr);" & _                 ; Gets a pointer to a POINT structure containing the default width (x) and height (y) measurements of an item, including the surrounding white space.
	"GetAutoArrange hresult();" & _                       ; Gets the current state of the folder's Auto Arrange mode.
	"SelectItem hresult(int;dword);" & _                  ; Selects an item in the folder's view.
	"SelectAndPositionItems hresult(uint;ptr;ptr;dword);" ; Allows the selection and positioning of items visible in the folder's view.

; --- ItemCount uFlags ---

Global Const $SVGIO_BACKGROUND     = 0x00000000 ; Refers to the background of the view. It is used with IID_IContextMenu to get a shortcut menu for the background of the view.
Global Const $SVGIO_SELECTION      = 0x00000001 ; Refers to the currently selected items. Used with IID_IDataObject to retrieve a data object that represents the selected items.
Global Const $SVGIO_ALLVIEW        = 0x00000002 ; Used in the same way as SVGIO_SELECTION but refers to all items in the view.
Global Const $SVGIO_CHECKED        = 0x00000003 ; Used in the same way as SVGIO_SELECTION but refers to checked items in views where checked mode is supported. For more details on checked mode, see FOLDERFLAGS.
Global Const $SVGIO_TYPE_MASK      = 0x0000000F ; Masks all bits but those corresponding to the SVGIO flags.
Global Const $SVGIO_FLAG_VIEWORDER = 0x80000000 ; Returns the items in the order they appear in the view. If this flag is not set, the selected item will be listed first.

; --- SelectItem dwFlags ---

Global Const $SVSI_DESELECT	      = 0          ; Deselect the item.
Global Const $SVSI_SELECT	        = 0x1        ; Select the item.
Global Const $SVSI_EDIT	          = 0x3        ; Put the name of the item into rename mode.
Global Const $SVSI_DESELECTOTHERS	= 0x4        ; Deselect all but the selected item. If the item parameter is NULL, deselect all items.
Global Const $SVSI_ENSUREVISIBLE	= 0x8        ; In the case of a folder that cannot display all of its contents on one screen, display the portion that contains the selected item.
Global Const $SVSI_FOCUSED	      = 0x10       ; Give the selected item the focus when multiple items are selected, placing the item first in any list of the collection returned by a method.
Global Const $SVSI_TRANSLATEPT	  = 0x20       ; Convert the input point from screen coordinates to the list-view client coordinates.
Global Const $SVSI_SELECTIONMARK	= 0x40       ; Mark the item so that it can be queried using IFolderView::GetSelectionMarkedItem.
Global Const $SVSI_POSITIONITEM	  = 0x80       ; Allows the window's default view to position the item. In most cases, this will place the item in the first available position. However, if the call comes during the processing of a mouse-positioned context menu, the position of the context menu is used to position the item.
Global Const $SVSI_CHECK	        = 0x100      ; The item should be checked. This flag is used with items in views where the checked mode is supported.
Global Const $SVSI_CHECK2	        = 0x200      ; The second check state when the view is in tri-check mode, in which there are three values for the checked state. You can indicate tri-check mode by specifying FWF_TRICHECKSELECT in IFolderView2::SetCurrentFolderFlags. The 3 states for FWF_TRICHECKSELECT are unchecked, SVSI_CHECK and SVSI_CHECK2.
Global Const $SVSI_KEYBOARDSELECT	= 0x401      ; Selects the item and marks it as selected by the keyboard. Includes the SVSI_SELECT flag.
Global Const $SVSI_NOTAKEFOCUS	  = 0x40000000 ; An operation to select or focus an item should not also set focus on the view itself.
Global Const $SVSI_NOSTATECHANGE  = 0x80000000 ; An operation to edit or position an item should not affect the item's focus or selected state.


; === IFolderView2 interface (Vista and later) ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775541(v=vs.85).aspx

; Exposes methods that retrieve information about a folder's display options,
; select specified items in that folder, and set the folder's view mode.

Global Const $sIID_IFolderView2  = "{1AF3A467-214F-4298-908E-06B03E0B39F9}"
Global Const $tRIID_IFolderView2 = _WinAPI_GUIDFromString( $sIID_IFolderView2 )
Global Const $dtag_IFolderView2  = $dtag_IFolderView & _ ; Inherits from IFolderView.
	"SetGroupBy hresult();" & _                       ; Groups the view by the given property key and direction.
	"GetGroupBy hresult();" & _                       ; Retrieves the property and sort order used for grouping items in the folder display.
	"SetViewProperty hresult();" & _                  ; Caches a property for an item in the view's property cache.
	"GetViewProperty hresult();" & _                  ; Gets a property value for a given property key from the view's cache.
	"SetTileViewProperties hresult();" & _            ; Set the list of tile properties for an item.
	"SetExtendedTileViewProperties hresult();" & _    ; Sets the list of extended tile properties for an item.
	"SetText hresult();" & _                          ; Sets the default text to be used when there are no items in the view.
	"SetCurrentFolderFlags hresult();" & _            ; Sets and applies specified folder flags.
	"GetCurrentFolderFlags hresult();" & _            ; Gets the currently applied folder flags.
	"GetSortColumnCount hresult();" & _               ; Gets the count of sort columns currently applied to the view.
	"SetSortColumns hresult();" & _                   ; Sets and sorts the view by the given sort columns.
	"GetSortColumns hresult();" & _                   ; Gets the sort columns currently applied to the view.
	"GetItem hresult();" & _                          ; Retrieves an object that represents a specified item.
	"GetVisibleItem hresult();" & _                   ; Gets the next visible item starting with iStart.
	"GetSelectedItem hresult();" & _                  ; Gets the next selected item starting with iStart.
	"GetSelection hresult();" & _                     ; Gets the current selection as an IShellItemArray.
	"GetSelectionState hresult();" & _                ; Gets the selection state including check state.
	"InvokeVerbOnSelection hresult(str);" & _         ; Invokes the given verb on the current selection.
	"SetViewModeAndIconSize hresult(uint;int);" & _   ; Sets and applies the view mode and image size.
	"GetViewModeAndIconSize hresult(uint*;int*);" & _ ; Gets the current view mode and icon size applied to the view.
	"SetGroupSubsetCount hresult();" & _              ; Turns on group subsetting and sets the number of visible rows of items in each group.
	"GetGroupSubsetCount hresult();" & _              ; Gets the count of visible rows displayed for a group's subset.
	"SetRedraw hresult(int);" & _                     ; Sets redraw on and off.
	"IsMoveInSameFolder hresult();" & _               ; Checks to see if this view sourced the current drag-and-drop or cut-and-paste operation (used by drop target objects).
	"DoRename hresult();"                             ; Starts a rename operation on the current selection.


; === IInputObject interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb761804(v=vs.85).aspx

; Exposes methods that change user interface (UI) activation and process accelerators for a user input object contained in the Shell.

Global Const $sIID_IInputObject  = "{68284FAA-6A48-11D0-8C78-00C04FD918B4}"
Global Const $tRIID_IInputObject = _WinAPI_GUIDFromString( $sIID_IInputObject )
Global Const $dtag_IInputObject  = _
	"UIActivateIO hresult(int;ptr);" & _   ; UI-activates or deactivates the object.
	"HasFocusIO hresult();" & _            ; Determines if one of the object's windows has the keyboard focus.
	"TranslateAcceleratorIO hresult(ptr);" ; Enables the object to process keyboard accelerators.

; --- MSG structure ---

; #STRUCTURE# ===================================================================================================================
; Name...........: $tagMSG
; Description ...: The MSG structure contains message information from a thread's message queue
; Fields ........: hwnd    - Handle to the window whose window procedure receives the message. hwnd is NULL when the message is a thread message.
;                  message - Specifies the message identifier. Applications can only use the low word; the high word is reserved by the system.
;                  wParam  - Specifies additional information about the message. The exact meaning depends on the value of the message member.
;                  lParam  - Specifies additional information about the message. The exact meaning depends on the value of the message member.
;                  time    - Specifies the time at which the message was posted.
;                  X       - Specifies the cursor position, X position in screen coordinates, when the message was posted.
;                  Y       - Y position in screen coordinates.
; ===============================================================================================================================

Global Const $tagMSG = "hwnd hwnd;uint message;wparam wParam;lparam lParam;dword time;int X;int Y"


; === IOleWindow interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/ms680102(v=vs.85).aspx

; The IOleWindow interface provides methods that allow an application to obtain the handle to the various windows that participate in in-place activation, and also to enter and exit context-sensitive help mode.

Global Const $dtag_IOleWindow = _
	"GetWindow hresult(hwnd*);" & _      ; Gets a window handle.
	"ContextSensitiveHelp hresult(int);" ; Controls enabling of context-sensitive help.


; === IPersist interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/ms688695(v=vs.85).aspx

; The IPersist interface defines the single method GetClassID, which is designed to supply the CLSID of an object
; that can be stored persistently in the system. A call to this method can allow the object to specify which object
; handler to use in the client process, as it is used in the OLE default implementation of marshaling.

Global Const $sIID_IPersist  = "{0000010C-0000-0000-C000-000000000046}"
Global Const $tRIID_IPersist = _WinAPI_GUIDFromString( $sIID_IPersist )
Global Const $dtag_IPersist = _
	"GetClassID hresult(ptr*);" ; Instructs a Shell folder object to initialize itself based on the information passed.


; === IPersistFolder interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775346(v=vs.85).aspx

; Exposes a method that initializes Shell folder objects.

Global Const $sIID_IPersistFolder  = "{000214EA-0000-0000-C000-000000000046}"
Global Const $tRIID_IPersistFolder = _WinAPI_GUIDFromString( $sIID_IPersistFolder )
Global Const $dtag_IPersistFolder  = $dtag_IPersist & _ ; Inherits from IPersist.
	"Initialize hresult(ptr);" ; Instructs a Shell folder object to initialize itself based on the information passed.


; === IPersistFolder2 interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775344(v=vs.85).aspx

; Exposes methods that obtain information from Shell folder objects.

Global Const $sIID_IPersistFolder2  = "{1AC3D9F0-175C-11d1-95BE-00609797EA4F}"
Global Const $tRIID_IPersistFolder2 = _WinAPI_GUIDFromString( $sIID_IPersistFolder2 )
Global Const $dtag_IPersistFolder2  = $dtag_IPersistFolder & _ ; Inherits from IPersistFolder.
	"GetCurFolder hresult(ptr*);" ; Gets the ITEMIDLIST for the folder object. 


; === IServiceProvider interface ===

; http://msdn.microsoft.com/en-us/library/cc678965(v=vs.85).aspx

; IServiceProvider provides a generic access mechanism to locate a GUID-identified service.

Global Const $sIID_IServiceProvider  = "{6D5140C1-7436-11CE-8034-00AA006009FA}"
Global Const $tRIID_IServiceProvider = _WinAPI_GUIDFromString( $sIID_IServiceProvider )
Global Const $dtag_IServiceProvider  = _
	"QueryService hresult(struct*;struct*;ptr*);" ; Acts as the factory method for any services exposed through an implementation of IServiceProvider.


; === IShellBrowser interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775123(v=vs.85).aspx

; IShellBrowser is implemented by hosts of Shell views (objects that implement IShellView).
; Exposes methods that provide services for the view it is hosting and other objects that run in the context of the Explorer window.

Global Const $sIID_IShellBrowser  = "{000214E2-0000-0000-C000-000000000046}"
Global Const $tRIID_IShellBrowser = _WinAPI_GUIDFromString( $sIID_IShellBrowser )
Global Const $dtag_IShellBrowser  = $dtag_IOleWindow & _         ; Inherits from IOleWindow.
	"InsertMenusSB hresult(handle;ptr);" & _                       ; Allows the container to insert its menu groups into the composite menu that is displayed when an extended namespace is being viewed or used.
	"SetMenuSB hresult(handle;handle;hwnd);" & _                   ; Installs the composite menu in the view window.
	"RemoveMenusSB hresult(handle);" & _                           ; Permits the container to remove any of its menu elements from the in-place composite menu and to free all associated resources.
	"SetStatusTextSB hresult(ptr);" & _                            ; Sets and displays status text about the in-place object in the container's frame-window status bar.
	"EnableModelessSB hresult(int);" & _                           ; Tells Windows Explorer to enable or disable its modeless dialog boxes.
	"TranslateAcceleratorSB hresult(ptr;word);" & _                ; Translates accelerator keystrokes intended for the browser's frame while the view is active.
	"BrowseObject hresult(ptr;uint);" & _                          ; Informs Microsoft Windows Explorer to browse to another folder.
	"GetViewStateStream hresult(dword;ptr*);" & _                  ; Gets an IStream interface that can be used for storage of view-specific state information.
	"GetControlWindow hresult(uint;hwnd);" & _                     ; Gets the window handle to a browser control.
	"SendControlMsg hresult(uint;uint;wparam;lparam;lresult);" & _ ; Sends control messages to either the toolbar or the status bar in a Windows Explorer window.
	"QueryActiveShellView hresult(ptr*);" & _                      ; Retrieves the currently active (displayed) Shell view object.
	"OnViewWindowActive hresult(ptr);" & _                         ; Called by the Shell view when the view window or one of its child windows gets the focus or becomes active.
	"SetToolbarItems hresult(ptr;uint;uint);"                      ; Adds toolbar items to Windows Explorer's toolbar.

Global Const $sSID_STopLevelBrowser  = "{4C96BE40-915C-11CF-99D3-00AA004AE837}"
Global Const $tRIID_STopLevelBrowser = _WinAPI_GUIDFromString( $sSID_STopLevelBrowser )

; --- BrowseObject wFlags ---

; Flags specifying the folder to be browsed. It can be zero or one or more of the following values.

; These flags specify whether another window is to be created.

Global Const $SBSP_DEFBROWSER            = 0x0000     ; Use default behavior, which respects the view option (the user setting to create new windows or to browse in place). In most cases, calling applications should use this flag.
Global Const $SBSP_SAMEBROWSER           = 0x0001     ; Browse to another folder with the same Windows Explorer window.
Global Const $SBSP_NEWBROWSER            = 0x0002     ; Creates another window for the specified folder.

; The following flags specify the mode. These values are ignored if SBSP_SAMEBROWSER is specified or if SBSP_DEFBROWSER is specified and the user has selected Browse In Place.

Global Const $SBSP_DEFMODE               = 0x0000     ; Use the current window.
Global Const $SBSP_OPENMODE              = 0x0010     ; Specifies no folder tree for the new browse window. If the current browser does not match the SBSP_OPENMODE of the browse object call, a new window is opened.
Global Const $SBSP_EXPLOREMODE           = 0x0020     ; Specifies a folder tree for the new browse window. If the current browser does not match the SBSP_EXPLOREMODE of the browse object call, a new window is opened.
Global Const $SBSP_HELPMODE              = 0x0040     ; Not supported. Do not use.
Global Const $SBSP_NOTRANSFERHIST        = 0x0080     ; Do not transfer the browsing history to the new window.

; The following flags specify the category of the pidl parameter.

Global Const $SBSP_ABSOLUTE              = 0x0000     ; An absolute pointer to an item identifier list (PIDL), relative to the desktop.
Global Const $SBSP_RELATIVE              = 0x1000     ; A relative PIDL, relative to the current folder.
Global Const $SBSP_PARENT                = 0x2000     ; Browse the parent folder, ignore the PIDL.
Global Const $SBSP_NAVIGATEBACK          = 0x4000     ; Navigate back, ignore the PIDL.
Global Const $SBSP_NAVIGATEFORWARD       = 0x8000     ; Navigate forward, ignore the PIDL.

Global Const $SBSP_ALLOW_AUTONAVIGATE    = 0x00010000 ; Enable auto-navigation.

; The following flags specify mode.

Global Const $SBSP_KEEPSAMETEMPLATE      = 0x00020000 ; Microsoft Windows Vista. Not supported. Do not use.
Global Const $SBSP_KEEPWORDWHEELTEXT     = 0x00040000 ; Microsoft Windows Vista. Navigate without clearing the search entry field.
Global Const $SBSP_ACTIVATE_NOFOCUS      = 0x00080000 ; Microsoft Windows Vista. Navigate without the default behavior of setting focus into the new view.

; The following flags control how history is manipulated as a result of navigation.

Global Const $SBSP_CREATENOHISTORY       = 0x00100000 ; Windows 7 and later. Do not add a new entry to the travel log. When the user enters a search term in the search box and subsequently refines the query, the browser navigates forward but does not add an additional travel log entry.
Global Const $SBSP_PLAYNOSOUND           = 0x00200000 ; Windows 7 and later. Do not make the navigation complete sound for each keystroke in the search box.
Global Const $SBSP_CALLERUNTRUSTED       = 0x00800000 ; Microsoft Internet Explorer 6 Service Pack 2 (SP2) and later. The navigation was possibly initiated by a Web page with scripting code already present on the local system.
Global Const $SBSP_TRUSTFIRSTDOWNLOAD    = 0x01000000 ; Microsoft Internet Explorer 6 Service Pack 2 (SP2) and later. The new window is the result of a user initiated action. Trust the new window if it immediately attempts to download content.
Global Const $SBSP_UNTRUSTEDFORDOWNLOAD  = 0x02000000 ; Microsoft Internet Explorer 6 Service Pack 2 (SP2) and later. The window is navigating to an untrusted, non-HTML file. If the user attempts to download the file, do not allow the download.
Global Const $SBSP_NOAUTOSELECT          = 0x04000000 ; Suppress selection in the history pane.
Global Const $SBSP_WRITENOHISTORY        = 0x08000000 ; Write no history of this navigation in the history Shell folder.
Global Const $SBSP_TRUSTEDFORACTIVEX     = 0x10000000 ; Microsoft Internet Explorer 6 Service Pack 2 (SP2) and later. The navigate should allow ActiveX prompts.
Global Const $SBSP_FEEDNAVIGATION        = 0x20000000 ; Microsoft Internet Explorer 7 and later. If allowed by current registry settings, give the browser a destination to navigate to.
Global Const $SBSP_REDIRECT              = 0x40000000 ; Enables redirection to another URL.
Global Const $SBSP_INITIATEDBYHLINKFRAME = 0x80000000


; === IShellFolder interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb775075(v=vs.85).aspx

; IShellFolder is exposed by all Shell namespace folder objects. Its methods are used to manage folders.

Global Const $sIID_IShellFolder  = "{000214E6-0000-0000-C000-000000000046}"
Global Const $tRIID_IShellFolder = _WinAPI_GUIDFromString( $sIID_IShellFolder )
Global Const $dtag_IShellFolder  = _
	"ParseDisplayName hresult(hwnd;ptr;wstr;ulong*;ptr*;ulong*);" & _ ; Translates the display name of a file object or a folder into an item identifier list.
	"EnumObjects hresult(hwnd;dword;ptr*);" & _                       ; Enables a client to determine the contents of a folder by creating an item identifier enumeration object and returning its IEnumIDList interface. The methods supported by that interface can then be used to enumerate the folder's contents.
	"BindToObject hresult(ptr;ptr;clsid;ptr*);" & _                   ; Retrieves a handler, typically the Shell folder object that implements IShellFolder for a particular item. Optional parameters that control the construction of the handler are passed in the bind context.
	"BindToStorage hresult(ptr;ptr;clsid;ptr*);" & _                  ; Requests a pointer to an object's storage interface.
	"CompareIDs hresult(lparam;ptr;ptr);" & _                         ; Determines the relative order of two file objects or folders, given their item identifier lists.
	"CreateViewObject hresult(hwnd;clsid;ptr*);" & _                  ; Requests an object that can be used to obtain information from or interact with a folder object.
	"GetAttributesOf hresult(uint;struct*;ulong*);" & _               ; Gets the attributes of one or more file or folder objects contained in the object represented by IShellFolder.
	"GetUIObjectOf hresult(hwnd;uint;struct*;clsid;uint*;ptr*);" & _  ; Gets an OLE interface that can be used to carry out actions on the specified file objects or folders.
	"GetDisplayNameOf hresult(ptr;dword;struct*);" & _                ; Retrieves the display name for the specified file object or subfolder.
	"SetNameOf hresult(hwnd;ptr;wstr;dword;ptr*);"                    ; Sets the display name of a file object or subfolder, changing the item identifier in the process.

; --- EnumObjects grfFlags ---

Global Const $SHCONTF_CHECKING_FOR_CHILDREN = 0x00010 ; Windows 7 and later. The calling application is checking for the existence of child items in the folder.
Global Const $SHCONTF_FOLDERS               = 0x00020 ; Include items that are folders in the enumeration.
Global Const $SHCONTF_NONFOLDERS            = 0x00040 ; Include items that are not folders in the enumeration.
Global Const $SHCONTF_INCLUDEHIDDEN         = 0x00080 ; Include hidden items in the enumeration. This does not include hidden system items. (To include hidden system items, use SHCONTF_INCLUDESUPERHIDDEN.)
Global Const $SHCONTF_INIT_ON_FIRST_NEXT    = 0x00100 ; No longer used; always assumed. IShellFolder::EnumObjects can return without validating the enumeration object. Validation can be postponed until the first call to IEnumIDList::Next. Use this flag when a user interface might be displayed prior to the first IEnumIDList::Next call. For a user interface to be presented, hwnd must be set to a valid window handle.
Global Const $SHCONTF_NETPRINTERSRCH        = 0x00200 ; The calling application is looking for printer objects.
Global Const $SHCONTF_SHAREABLE             = 0x00400 ; The calling application is looking for resources that can be shared.
Global Const $SHCONTF_STORAGE               = 0x00800 ; Include items with accessible storage and their ancestors, including hidden items.
Global Const $SHCONTF_NAVIGATION_ENUM       = 0x01000 ; Windows 7 and later. Child folders should provide a navigation enumeration.
Global Const $SHCONTF_FASTITEMS             = 0x02000 ; Windows Vista and later. The calling application is looking for resources that can be enumerated quickly.
Global Const $SHCONTF_FLATLIST              = 0x04000 ; Windows Vista and later. Enumerate items as a simple list even if the folder itself is not structured in that way.
Global Const $SHCONTF_ENABLE_ASYNC          = 0x08000 ; Windows Vista and later. The calling application is monitoring for change notifications. This means that the enumerator does not have to return all results. Items can be reported through change notifications.
Global Const $SHCONTF_INCLUDESUPERHIDDEN    = 0x10000 ; Windows 7 and later. Include hidden system items in the enumeration. This value does not include hidden non-system items. (To include hidden non-system items, use SHCONTF_INCLUDEHIDDEN.)

; --- GetDisplayNameOf uFlags ---

Global Const $SHGDN_NORMAL        = 0x0000 ; When not combined with another flag, return the parent-relative name that identifies the item, suitable for displaying to the user. This name often does not include extra information such as the file name extension and does not need to be unique. This name might include information that identifies the folder that contains the item. For instance, this flag could cause IShellFolder::GetDisplayNameOf to return the string "username (on Machine)" for a particular user's folder.
Global Const $SHGDN_INFOLDER      = 0x0001 ; The name is relative to the folder from which the request was made. This is the name display to the user when used in the context of the folder. For example, it is used in the view and in the address bar path segment for the folder. This name should not include disambiguation information—for instance "username" instead of "username (on Machine)" for a particular user's folder. Use this flag in combinations with SHGDN_FORPARSING and SHGDN_FOREDITING.
Global Const $SHGDN_FOREDITING    = 0x1000 ; The name is used for in-place editing when the user renames the item.
Global Const $SHGDN_FORADDRESSBAR = 0x4000 ; The name is displayed in an address bar combo box.
Global Const $SHGDN_FORPARSING    = 0x8000 ; The name is used for parsing. That is, it can be passed to IShellFolder::ParseDisplayName to recover the object's PIDL. The form this name takes depends on the particular object. When SHGDN_FORPARSING is used alone, the name is relative to the desktop. When combined with SHGDN_INFOLDER, the name is relative to the folder from which the request was made.

Global Const $tagSTRRET = "uint uType;ptr data;"


; === IShellView interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774834(v=vs.85).aspx

; IShellView exposes methods that present a view in the Microsoft Windows Explorer or folder windows.

Global Const $sIID_IShellView  = "{000214E3-0000-0000-C000-000000000046}"
Global Const $tRIID_IShellView = _WinAPI_GUIDFromString( $sIID_IShellView )
Global Const $dtag_IShellView  = $dtag_IOleWindow & _    ; Inherits from IOleWindow.
	"TranslateAccelerator hresult(ptr);" & _               ; Translates accelerator key strokes when a namespace extension's view has the focus.
	"EnableModeless hresult(int);" & _                     ; Enables or disables modeless dialog boxes. This method is currently not implemented.
	"UIActivate hresult(uint);" & _                        ; Called when the activation state of the view window is changed by an event that is not caused by the shell view itself. For example, if the TAB key is pressed when the tree has the focus, the view should be given the focus.
	"Refresh hresult();" & _                               ; Refreshes the view's contents in response to user input.
	"CreateViewWindow hresult(ptr;ptr;ptr;ptr;hwnd*);" & _ ; Creates a view window. This can be either the right pane of Windows Explorer or the client window of a folder window.
	"DestroyViewWindow hresult();" & _                     ; Destroys the view window.
	"GetCurrentInfo hresult(ptr*);" & _                    ; Retrieves the current folder settings.
	"AddPropertySheetPages hresult(dword;ptr;lparam);" & _ ; Allows the view to add pages to the Options property sheet from the View menu.
	"SaveViewState hresult();" & _                         ; Saves the shell's view settings so the current state can be restored during a subsequent browsing session.
	"SelectItem hresult(ptr;uint);" & _                    ; Changes the selection state of one or more items within the shell view window.
	"GetItemObject hresult(uint;struct*;ptr*);"            ; Retrieves an interface that refers to data presented in the view.

; --- CreateViewWindow FOLDERSETTINGS structure ---

Global Const $tagFOLDERSETTINGS = _
	"uint ViewMode;" & _ ; Folder view mode. This can be one of the FOLDERVIEWMODE values.
	"uint fFlags"        ; A set of flags that indicate the options for the folder. This can be zero or a combination of the FOLDERFLAGS values.

; FOLDERVIEWMODE values
Global Const $FVM_FIRST      = 1 ; Specifies a convenience constant equal to the first constant in FOLDERVIEWMODE.
Global Const $FVM_ICON       = 1 ; The view should display medium-size icons.
Global Const $FVM_SMALLICON  = 2 ; The view should display small icons.
Global Const $FVM_LIST       = 3 ; Object names are displayed in a list view.
Global Const $FVM_DETAILS    = 4 ; Object names and other selected information, such as the size or date last updated, are shown.
Global Const $FVM_THUMBNAIL  = 5 ; The view should display thumbnail icons.
Global Const $FVM_TILE       = 6 ; The view should display large icons.
Global Const $FVM_THUMBSTRIP = 7 ; The view should display icons in a filmstrip format.
Global Const $FVM_CONTENT    = 8 ; Windows 7 and later. The view should display content mode.
Global Const $FVM_LAST       = 8 ; Specifies a convenience constant equal to the last constant in FOLDERVIEWMODE.

; FOLDERFLAGS values
Global Const $FWF_AUTOARRANGE         = 0x00000001 ; Automatically arrange the elements in the view. This implies LVS_AUTOARRANGE if the list view control is used to implement the view.
Global Const $FWF_ABBREVIATEDNAMES    = 0x00000002 ; Abbreviate names. This flag is not currently supported.
Global Const $FWF_NOBACKBROWSING      = 0x00000002 ; Do not allow back (upward) browsing with Backspace key.
Global Const $FWF_SNAPTOGRID          = 0x00000004 ; Arrange items on a grid. This flag is not currently supported.
Global Const $FWF_OWNERDATA           = 0x00000008 ; This flag is not currently supported.
Global Const $FWF_BESTFITWINDOW       = 0x00000008 ; Enable the best-fit window mode. Let the view size the window so that its contents fit inside the view window in the best possible manner.
Global Const $FWF_DESKTOP             = 0x00000020 ; Make the folder behave like the desktop. This value applies only to the desktop view and is not used for typical Shell folders.
Global Const $FWF_SINGLESEL           = 0x00000040 ; Do not allow more than a single item to be selected. This is used in the common dialog boxes.
Global Const $FWF_NOSUBFOLDERS        = 0x00000080 ; Do not show subfolders.
Global Const $FWF_TRANSPARENT         = 0x00000100 ; Draw transparently. This is used only for the desktop.
Global Const $FWF_NOCLIENTEDGE        = 0x00000200 ; This flag is not currently supported.
Global Const $FWF_NOSCROLL            = 0x00000400 ; Do not add scroll bars. This is used only for the desktop.
Global Const $FWF_ALIGNLEFT           = 0x00000800 ; The view should be left-aligned. This implies LVS_ALIGNLEFT if the list view control is used to implement the view.
Global Const $FWF_NOICONS             = 0x00001000 ; The view should not display icons.
Global Const $FWF_SHOWSELALWAYS       = 0x00002000 ; Always show the selection.
Global Const $FWF_NOVISIBLE           = 0x00004000 ; This flag is not currently supported.
Global Const $FWF_SINGLECLICKACTIVATE = 0x00008000 ; This flag is not currently supported.
Global Const $FWF_NOWEBVIEW           = 0x00010000 ; The view should not be shown as a Web view.
Global Const $FWF_HIDEFILENAMES       = 0x00020000 ; The view should not display file names.
Global Const $FWF_CHECKSELECT         = 0x00040000 ; Turns on the check mode for the view.
Global Const $FWF_NOENUMREFRESH       = 0x00080000 ; Windows Vista: Do not re-enumerate the view (or drop the current contents of the view) when the view is refreshed.
Global Const $FWF_NOGROUPING          = 0x00100000 ; Windows Vista: Do not allow grouping in the view
Global Const $FWF_FULLROWSELECT       = 0x00200000 ; Windows Vista: When an item is selected, the item and all its sub-items are highlighted.
Global Const $FWF_NOFILTERS           = 0x00400000 ; Windows Vista: Do not display filters in the view.
Global Const $FWF_NOCOLUMNHEADER      = 0x01000000 ; Windows Vista: Do not display a column header in the view in any view mode.
Global Const $FWF_NOHEADERINALLVIEWS  = 0x02000000 ; Windows Vista: Only show the column header in details view mode.
Global Const $FWF_EXTENDEDTILES       = 0x01000000 ; Windows Vista: When the view is in "tile view mode" the layout of a single item should be extended to the width of the view.
Global Const $FWF_TRICHECKSELECT      = 0x02000000 ; Windows Vista: Check boxes have 3 modes: unchecked, SVSI_CHECK, SVSI_CHECK2.
Global Const $FWF_AUTOCHECKSELECT     = 0x04000000 ; Windows Vista: Items can be selected using checkboxes.
Global Const $FWF_NOBROWSERVIEWSTATE  = 0x08000000 ; Windows Vista: The view should not save view state in the browser.
Global Const $FWF_SUBSETGROUPS        = 0x10000000 ; Windows Vista: The view should list the number of items displayed in each group. To be used with IFolderView2::SetGroupSubsetCount.
Global Const $FWF_USESEARCHFOLDER     = 0x40000000 ; Windows Vista: Use the search folder for stacking and searching.
Global Const $FWF_ALLOWRTLREADING     = 0x80000000 ; Windows Vista: Allow right-to-left reading.

; --- UIActivate uState ---

; Flag specifying the activation state of the window. This parameter can be one of the following values.

Global Const $SVUIA_DEACTIVATE       = 0 ; Windows Explorer is about to destroy the Shell view window. The Shell view should remove all extended user interfaces. These are typically merged menus and merged modeless pop-up windows.
Global Const $SVUIA_ACTIVATE_NOFOCUS = 1 ; The Shell view is losing the input focus, or it has just been created without the input focus. The Shell view should be able to set menu items appropriate for the nonfocused state. This means no selection-specific items should be added.
Global Const $SVUIA_ACTIVATE_FOCUS   = 2 ; Microsoft Windows Explorer has just created the view window with the input focus. This means the Shell view should be able to set menu items appropriate for the focused state.
Global Const $SVUIA_INPLACEACTIVATE  = 3 ; The Shell view is active without focus. This flag is only used when UIActivate is exposed through the IShellView2 interface.


; === IShellView2 interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774824(v=vs.85).aspx

; Extends the capabilities of IShellView.

Global Const $sIID_IShellView2  = "{88E39E80-3578-11CF-AE69-08002B2E1262}"
Global Const $tRIID_IShellView2 = _WinAPI_GUIDFromString( $sIID_IShellView2 )
Global Const $dtag_IShellView2  = $dtag_IShellView & _ ; Inherits from IShellView.
	"GetView hresult();" & _           ; Requests the current or default Shell view, together with all other valid view identifiers (VIDs) supported by this implementation of IShellView2.
	"CreateViewWindow2 hresult();" & _ ; Used to request the creation of a new Shell view window. It can be either the right pane of Microsoft Windows Explorer or the client window of a folder window.
	"HandleRename hresult();" & _      ; Used to change an item's identifier.
	"SelectAndPositionItem hresult();" ; Selects and positions an item in a Shell View.


; === IShellView3 interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774820(v=vs.85).aspx

; Extends the capabilities of IShellView2 by providing a method to replace IShellView2::CreateViewWindow2.

Global Const $sIID_IShellView3  = "{EC39FA88-F8AF-41C5-8421-38BED28F4673}"
Global Const $tRIID_IShellView3 = _WinAPI_GUIDFromString( $sIID_IShellView3 )
Global Const $dtag_IShellView3  = $dtag_IShellView2 & _                     ; Inherits from IShellView2.
	"CreateViewWindow3 hresult(ptr;ptr;dword;dword;dword;uint;ptr;ptr;ptr*);" ; Requests the creation of a new Shell view window. The view can be either the right pane of Microsoft Windows Explorer or the client window of a folder window.


; === IShellWindows interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/cc836570(v=vs.85).aspx

; Provides access to the collection of open Shell windows.

Global Const $sCLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
Global Const $tCLSID_ShellWindows = _WinAPI_GUIDFromString( $sCLSID_ShellWindows )
Global Const $sIID_IShellWindows  = "{85CB6900-4D95-11CF-960C-0080C7F4EE85}"
Global Const $tRIID_IShellWindows = _WinAPI_GUIDFromString( $sIID_IShellWindows )
Global Const $dtag_IShellWindows  = $dtag_IDispatch & _ ; Inherits from IDispatch.
 	"get_Count hresult(long*);" & _   ; Gets the number of windows in the Shell windows collection.
 	"Item hresult(variant;ptr*);" & _ ; Returns the registered Shell window for a specified index.
	"_NewEnum hresult();" & _         ; Retrieves an enumerator for the collection of Shell windows.
 	"Register hresult();" & _         ; Registers an open window as a Shell window; the window is specified by handle.
 	"RegisterPending hresult();" & _  ; Registers a pending window as a Shell window; the window is specified by an absolute PIDL.
 	"Revoke hresult();" & _           ; Revokes a Shell window's registration and removes the window from the Shell windows collection.
 	"OnNavigate hresult();" & _       ; Occurs when a Shell window is navigated to a new location.
 	"OnActivated hresult();" & _      ; Occurs when a Shell window's activation state changes.
 	"FindWindowSW hresult();" & _     ; Finds a window in the Shell windows collection and returns the window's handle and IDispatch interface.
 	"OnCreated hresult();" & _        ; Occurs when a new Shell window is created for a frame.
 	"ProcessAttachDetach hresult();"  ; Deprecated. Always returns S_OK.


; === IUnknown interface ===

; http://msdn.microsoft.com/en-us/library/windows/desktop/ms680509(v=vs.85).aspx

; The IUnknown interface lets clients get pointers to other interfaces on a given object through the QueryInterface method, and manage the
; existence of the object through the IUnknown::AddRef and IUnknown::Release methods. All other COM interfaces are inherited, directly or
; indirectly, from IUnknown. Therefore, the three methods in IUnknown are the first entries in the VTable for every interface.

Global Const $sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Global Const $dtag_IUnknown = _
	"QueryInterface hresult(ptr;ptr*);" & _ ; Returns pointers to supported interfaces.
	"AddRef ulong();" & _                   ; Increments reference count.
	"Release ulong();"                      ; Decrements reference count.


; === IWebBrowser interface ===

Global Const $sIID_IWebBrowser = "{EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B}"
Global Const $dtag_IWebBrowser  = $dtag_IDispatch & _ ; Inherits from IDispatch.
	"GoBack hresult();" & _
	"GoForward hresult();" & _
	"GoHome hresult();" & _
	"GoSearch hresult();" & _
	"Navigate hresult();" & _
	"Refresh hresult();" & _
	"Refresh2 hresult();" & _
	"Stop hresult();" & _
	"get_Application hresult();" & _
	"get_Parent hresult();" & _
	"get_Container hresult();" & _
	"get_Document hresult();" & _
	"get_TopLevelContainer hresult();" & _
	"get_Type hresult();" & _
	"get_Left hresult();" & _
	"put_Left hresult();" & _
	"get_Top hresult();" & _
	"put_Top hresult();" & _
	"get_Width hresult();" & _
	"put_Width hresult();" & _
	"get_Height hresult();" & _
	"put_Height hresult();" & _
	"get_LocationName hresult();" & _
	"get_LocationURL hresult();" & _
	"get_Busy hresult();"


; === IWebBrowserApp interface ===

Global Const $sIID_IWebBrowserApp  = "{0002DF05-0000-0000-C000-000000000046}"
Global Const $tRIID_IWebBrowserApp = _WinAPI_GUIDFromString( $sIID_IWebBrowserApp )
Global Const $dtag_IWebBrowserApp  = $dtag_IWebBrowser & _ ; Inherits from IWebBrowser.
	"Quit hresult();" & _
	"ClientToWindow hresult();" & _
	"PutProperty hresult();" & _
	"GetProperty hresult();" & _
	"get_Name hresult();" & _
	"get_HWND hresult(hwnd*);" & _
	"get_FullName hresult();" & _
	"get_Path hresult();" & _
	"get_Visible hresult();" & _
	"put_Visible hresult();" & _
	"get_StatusBar hresult();" & _
	"put_StatusBar hresult();" & _
	"get_StatusText hresult();" & _
	"put_StatusText hresult();" & _
	"get_ToolBar hresult();" & _
	"put_ToolBar hresult();" & _
	"get_MenuBar hresult();" & _
	"put_MenuBar hresult();" & _
	"get_FullScreen hresult();" & _
	"put_FullScreen hresult();"
