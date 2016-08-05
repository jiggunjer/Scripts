#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global $Value = -1

ConsoleWrite(_WinAPI_ShortToWord($Value) & @CR)
