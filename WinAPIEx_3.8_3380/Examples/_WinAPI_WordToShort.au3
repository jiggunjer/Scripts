#Include <WinAPIEx.au3>

Opt('MustDeclareVars', 1)

Global $Value = 65535

ConsoleWrite(_WinAPI_WordToShort($Value) & @CR)
