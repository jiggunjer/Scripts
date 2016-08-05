#Include <APIConstants.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global $File = _WinAPI_SaveFileDlg('', @WorkingDir, 'AutoIt v3 Scripts (*.au3)|All Files (*.*)', 1, 'MyScript.au3', '.', BitOR($OFN_PATHMUSTEXIST, $OFN_OVERWRITEPROMPT))

If Not @error Then
	ConsoleWrite($File & @CR)
EndIf
