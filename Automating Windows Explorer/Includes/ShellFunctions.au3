#include-once

#include <FileConstants.au3>
#include <WinAPIShellEx.au3>

Global Const $NULL    = 0x00000000
Global Const $S_FALSE = 0x00000001

Global Const $MAX_PATH = 260

; PIDL for Desktop. Empty PIDL only containing the terminating 2-byte NULL.
Global Const $pPidlAbsDesktop = _WinAPI_ShellGetSpecialFolderLocation( $CSIDL_DESKTOP )

; Flag used with CoCreateInstance
Global Const $CLSCTX_INPROC_SERVER = 0x1 ; The code that creates and manages objects of this class is a DLL that runs in the same process as the caller of the function specifying the class context.
Global Const $CLSCTX_ALL = 0x17

Global Const $dllOle32   = DllOpen( "ole32.dll" )
Global Const $dllShell32 = DllOpen( "shell32.dll" )
Global Const $dllShlwapi = DllOpen( "shlwapi.dll" )
Global Const $dllUser32  = DllOpen( "user32.dll" )

OnAutoItExitRegister( "ShellFunctionsClose" )


; --- Shell functions ---

Func CoCreateInstance( $rclsid, $pUnkOuter, $ClsContext, $riid, ByRef $ppv )
	Local $aRet = DllCall( $dllOle32, "long_ptr", "CoCreateInstance", "ptr", DllStructGetPtr($rclsid), "ptr", $pUnkOuter, "dword", $ClsContext, "ptr", DllStructGetPtr($riid), "ptr*", 0 )
	If @error Then Return SetError(1,0,0)
	$ppv = $aRet[5]
	Return $aRet[0]
EndFunc

Func ILAppendID( $pidl, $pmkid, $fAppend )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILAppendID", "ptr", $pidl, "ptr", $pmkid, "int", $fAppend )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILClone( $pidl )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILClone", "ptr", $pidl )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILCombine( $pidlAbs, $pidlRel )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILCombine", "ptr", $pidlAbs, "ptr", $pidlRel )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILCreateFromPath( $sPath )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILCreateFromPathW", "wstr", $sPath )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILFindChild( $pidlParent, $pidlChild )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILFindChild", "ptr", $pidlParent, "ptr", $pidlChild )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILFindLastID( $pidl )
	Local $aRet = DllCall( $dllShell32, "ptr", "ILFindLastID", "ptr", $pidl )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILGetSize( $pidl )
	Local $aRet = DllCall( $dllShell32, "uint", "ILGetSize", "ptr", $pidl )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILIsEqual( $pidl1, $pidl2 )
	Local $aRet = DllCall( $dllShell32, "int", "ILIsEqual", "ptr", $pidl1, "ptr", $pidl2 )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILIsParent( $pidlParent, $pidlChild, $fImmediate )
	Local $aRet = DllCall( $dllShell32, "int", "ILIsParent", "ptr", $pidlParent, "ptr", $pidlChild, "int", $fImmediate )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func ILRemoveLastID( ByRef $pidl )
	Local $aRet = DllCall( $dllShell32, "int", "ILRemoveLastID", "ptr", $pidl )
	If @error Then Return SetError(1, 0, 0)
	$pidl = $aRet[1]
	Return $aRet[0]
EndFunc

Func OleInitialize()
	Local $aRet = DllCall($dllOle32, "long", "OleInitialize", "ptr", 0)
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func OleUninitialize()
	DllCall($dllOle32, "none", "OleUninitialize")
	If @error Then Return SetError(1, 0, 0)
	Return 0
EndFunc

Func SHBindToParent( $pidl, $riid, ByRef $pIShellFolder, ByRef $pidlRel )
	Local $aRet = DllCall( $dllShell32, "long", "SHBindToParent", "ptr", $pidl, "ptr", $riid, "ptr*", 0, "ptr*", 0 )
	If @error Then Return SetError(1, 0, 0)
	$pIShellFolder = $aRet[3]
	$pidlRel = $aRet[4]
	Return $aRet[0]
EndFunc

; Copied from WinAPIEx.au3 by Yashied and modified to identify
; the file or folder through a pointer to a fully qualified PIDL.
Func ShellGetFileInfo($pidl, $iFlags, $iAttributes, ByRef $tSHFILEINFO)
	Local $aRet = DllCall($dllShell32, 'dword_ptr', 'SHGetFileInfoW', 'ptr', $pidl, 'dword', $iAttributes, 'ptr', DllStructGetPtr($tSHFILEINFO), 'uint', DllStructGetSize($tSHFILEINFO), 'uint', $iFlags)
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func SHGetDesktopFolder( ByRef $pIShellFolder )
	Local $aRet = DllCall( $dllShell32, "uint", "SHGetDesktopFolder", "ptr*", 0 )
	If @error Then Return SetError(1, 0, 0)
	$pIShellFolder = $aRet[1]
	Return $aRet[0]
EndFunc

Func SHGetPathFromIDList( $pidl, ByRef $sPath )
	Local $stPath = DllStructCreate( "wchar[" & $MAX_PATH & "]" )
	Local $aRet = DllCall( $dllShell32, "int", "SHGetPathFromIDListW", "ptr", $pidl, "ptr", DllStructGetPtr( $stPath ) )
	If @error Then Return SetError(1, 0, 0)
	$sPath = DllStructGetData( $stPath, 1 )
	Return $aRet[0]
EndFunc

Func SHGetSpecialFolderLocation( $hwndOwner, $nFolder, ByRef $ppidl )
	Local $aRet = DllCall( $dllShell32, "uint", "SHGetSpecialFolderLocation", "hwnd", $hwndOwner, "int", $nFolder, "ptr*", 0 )
	If @error Then Return SetError(1, 0, 0)
	$ppidl = $aRet[3]
	Return $aRet[0]
EndFunc

Func StrRetToBuf( $pSTRRET, $pidl, ByRef $sBuf, $iBuf = 512 )
	Local $aRet = DllCall( $dllShlwapi, "long", "StrRetToBufW", "ptr", $pSTRRET, "ptr", $pidl, "wstr", $sBuf, "uint", $iBuf )
	If @error Then Return SetError(1, 0, 0)
	$sBuf = $aRet[3]
	Return $aRet[0]
EndFunc


; --- Other functions ---

Func GetIconIndex( $sPath )
	Local $fString = IsString( $sPath )
	Local $pidl = $fString ? ILCreateFromPath( $sPath ) : $sPath
	Local $tSHFILEINFO = DllStructCreate( $tagSHFILEINFO )
	Local $iFlags = BitOR( $SHGFI_PIDL, $SHGFI_SYSICONINDEX )
	ShellGetFileInfo( $pidl, $iFlags, 0, $tSHFILEINFO )
	If @error Then Return SetError(1, 0, 0)
	Local $iIcon = DllStructGetData( $tSHFILEINFO, "iIcon" )
	_WinAPI_DestroyIcon( DllStructGetData( $tSHFILEINFO, "hIcon" ) )
	If Not $fString Then Return $iIcon
	_WinAPI_CoTaskMemFree( $pidl )
	Return $iIcon
EndFunc

; Copied from "SMF - Search my Files" by KaFu
Func SystemImageList( $bLargeIcons = False )
	Local $tSHFILEINFO = DllStructCreate( $tagSHFILEINFO )
	Local $dwFlags = BitOR( $SHGFI_USEFILEATTRIBUTES, $SHGFI_SYSICONINDEX )
	If Not $bLargeIcons Then $dwFlags = BitOR( $dwFlags, $SHGFI_SMALLICON )
	Local $hIml = _WinAPI_ShellGetFileInfo( ".txt", $dwFlags, $FILE_ATTRIBUTE_NORMAL, $tSHFILEINFO )
	If @error Then Return SetError( @error, 0, 0 )
	Return $hIml
EndFunc

Func ObjErrFunc( $oError )
	ConsoleWrite( "Err.Number is:         " & $oError.number & @CRLF & _
	              "Err.Windescription is: " & StringStripWS( $oError.windescription, 2 ) & @CRLF & _
	              "Err.Description is:    " & $oError.description & @CRLF & _
	              "Err.Source is:         " & $oError.source & @CRLF & _
	              "Err.Helpfile is:       " & $oError.helpfile & @CRLF & _
	              "Err.Helpcontext is:    " & $oError.helpcontext & @CRLF & _
	              "Err.LastDLLerror is:   " & $oError.lastdllerror & @CRLF & _
	              "Err.Scriptline is:     " & $oError.scriptline & @CRLF & _
	              "Err.Retcode is:        " & $oError.retcode & @CRLF & @CRLF )
EndFunc

Func ShellFunctionsClose()
	_WinAPI_CoTaskMemFree( $pPidlAbsDesktop )
	DllClose( $dllOle32 )
	DllClose( $dllShell32 )
	DllClose( $dllShlwapi )
	DllClose( $dllUser32 )
EndFunc
