#include "Includes\AutomatingWindowsExplorer.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
  ; Windows Explorer on XP, Vista, 7, 8
  Local $hExplorer = WinGetHandle( "[REGEXPCLASS:^(Cabinet|Explore)WClass$]" )
  If Not $hExplorer Then
    MsgBox( 0, "Automating Windows Explorer", "Could not find Windows Explorer. Terminating." )
    Return
  EndIf

  ; Get an IShellBrowser interface
  GetIShellBrowser( $hExplorer )
  If Not IsObj( $oIShellBrowser ) Then
    MsgBox( 0, "Automating Windows Explorer", "Could not get an IShellBrowser interface. Terminating." )
    Return
  EndIf

  ; Get other interfaces
  GetShellInterfaces()

  ; Get current folder
  Local $pFolder = GetCurrentFolder(), $sFolder
  SHGetPathFromIDList( $pFolder, $sFolder )
  MsgBox( 0, "Folder", $sFolder )

  ; Free memory
  _WinAPI_CoTaskMemFree( $pFolder )
EndFunc