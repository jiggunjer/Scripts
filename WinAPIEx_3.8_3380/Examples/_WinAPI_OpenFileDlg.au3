#Include <APIConstants.au3>
#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global $File

; Single file selection (returns the fully qualified path)
$File = _WinAPI_OpenFileDlg('', @WorkingDir, 'AutoIt v3 Scripts (*.au3)|All Files (*.*)', 1, '', '', BitOR($OFN_PATHMUSTEXIST, $OFN_FILEMUSTEXIST, $OFN_HIDEREADONLY))
If Not @error Then
	ConsoleWrite($File & @CR)
EndIf

; Multiple file selection (returns an array)
$File = _WinAPI_OpenFileDlg('', @WorkingDir, 'AutoIt v3 Scripts (*.au3)|All Files (*.*)', 1, '', '', BitOR($OFN_PATHMUSTEXIST, $OFN_FILEMUSTEXIST, $OFN_HIDEREADONLY, $OFN_ALLOWMULTISELECT, $OFN_EXPLORER))
If Not @error Then
	ConsoleWrite('--------------------------------------------------' & @CR)
	For $i = 1 To $File[0]
		ConsoleWrite($File[$i] & @CR)
	Next
EndIf
