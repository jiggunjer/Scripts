#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global Const $File = 'Notepad.exe'

ConsoleWrite(_WinAPI_SearchPath($File) & @CR)
ConsoleWrite(_WinAPI_SearchPath($File, @SystemDir) & @CR)
